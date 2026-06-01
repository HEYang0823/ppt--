#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT Agent — 自然语言驱动的 PPT 全自动生产流水线

用法:
  python3 ppt_agent.py "帮我做一个关于数据合规的PPT，20页，参考文件在 ~/Downloads/refs/"

工作流:
  自然语言输入 → 读取参考PPTX → 调用 LLM API → 生成 slides_config.json → 调用 make_ppt.py → 输出 PPTX + Prompts

支持的 LLM:
  - Anthropic Claude (默认, 需要 ANTHROPIC_API_KEY)
  - OpenAI GPT (需要 OPENAI_API_KEY)
  - DeepSeek (需要 DEEPSEEK_API_KEY)

API Key 设置:
  方式1: export ANTHROPIC_API_KEY="sk-ant-..."
  方式2: 写在 .env 文件中
"""

import os, sys, json, argparse, re
from pathlib import Path

# ============================================================
# 0. 尝试加载 .env
# ============================================================
def load_dotenv():
    env_path = Path(__file__).parent / '.env'
    if env_path.exists():
        with open(env_path) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, val = line.split('=', 1)
                    os.environ[key.strip()] = val.strip().strip('"').strip("'")

load_dotenv()

# ============================================================
# 1. 读取参考 PPTX 文件
# ============================================================
def read_reference_files(paths):
    """读取参考 PPTX 文件，返回结构化摘要"""
    try:
        from pptx import Presentation
    except ImportError:
        print("⚠️  未安装 python-pptx, 跳过参考文件读取")
        print("   安装: pip3 install python-pptx")
        return []

    refs = []
    for p in paths:
        p = Path(os.path.expanduser(p))
        if p.is_file() and p.suffix in ('.pptx', '.ppt'):
            files = [p]
        elif p.is_dir():
            files = sorted([f for f in p.iterdir() if f.suffix in ('.pptx', '.ppt') and not f.name.startswith('~')])
        else:
            print(f"  ⚠️  路径不存在: {p}")
            continue

        for f in files:
            try:
                prs = Presentation(str(f))
                slides = []
                for i, slide in enumerate(prs.slides, 1):
                    texts = []
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            for para in shape.text_frame.paragraphs:
                                t = para.text.strip()
                                if t and len(t) > 3:
                                    texts.append(t[:200])
                                    if len(texts) >= 3:
                                        break
                        if len(texts) >= 3:
                            break
                    if texts:
                        slides.append({'n': i, 'title': texts[0][:100], 'content': texts[0][:200]})

                refs.append({
                    'file': f.name,
                    'slides': len(prs.slides),
                    'sample': slides[:15]  # 最多取前15页
                })
                print(f"  ✅ {f.name} ({len(prs.slides)} slides)")
            except Exception as e:
                print(f"  ⚠️  无法读取 {f.name}: {e}")

    return refs

# ============================================================
# 2. 加载 Schema 和示例
# ============================================================
def load_schema_and_example():
    """加载 slides_config.json 的 schema 描述和示例"""
    engine_dir = Path(__file__).parent

    # 尝试加载示例配置
    example_path = engine_dir / 'examples' / 'slides_compliance.json'
    example_json = None
    if example_path.exists():
        with open(example_path, 'r', encoding='utf-8') as f:
            example_json = f.read()

    schema = """
## slides_config.json 格式说明

每页是一个 JSON 对象，支持以下 type：

| type | 用途 | 示例 |
|------|------|------|
| cover | 封面 | {"type":"cover","title":"...","subtitle":"...","icons":["🏭","🌾"],"image_hint":"..."} |
| content | 通用内容 | {"type":"content","title":"...","subtitle":"...","cards":[...],"info_boxes":[...]} |
| cards | 多卡片并列 | {"type":"cards","title":"...","cards":[{"title":"...","icon":"...","color":"ACC_BLUE","lines":["..."]}]} |
| table | 对比表格 | {"type":"table","title":"...","table":{"headers":["列1","列2"],"rows":[["...","..."]],"widths":[2.0,3.0]}} |
| layers | 多层架构 | {"type":"layers","title":"...","layers":[{"label":"层名","desc":"说明","color":"ACC_BLUE"}]} |
| comparison | 左右对比 | {"type":"comparison","title":"...","left":{"header":"✘ 左","items":[["标题","说明"]]},"right":{"header":"✔ 右","items":["要点"]}} |
| summary | 编号总结 | {"type":"summary","title":"...","items":[{"q":"问题?","detail":"说明","color":"ACC_BLUE"}],"next":"下节课预告"} |
| dalle | AI图像页 | {"type":"dalle","title":"...","subtitle":"...","cards":[...],"dalle_desc":"描述要生成的图","image_hint":"配图建议"} |

颜色常量: ACC_BLUE, GREEN_OK, WARM_ACC, RED_WARN, PURPLE, PRI_DARK, TXT_MED
"""
    return schema, example_json

# ============================================================
# 3. 构建 LLM Prompt
# ============================================================
def build_agent_prompt(user_request, refs, schema, example_json):
    """构建发给 LLM 的完整 prompt"""
    ref_text = ""
    if refs:
        ref_text = "\n## 参考文件内容（从本地 PPTX 提取）\n\n"
        for r in refs:
            ref_text += f"### {r['file']} ({r['slides']} 页)\n"
            for s in r['sample'][:8]:
                ref_text += f"- [页{s['n']}] {s['title']}\n"
            ref_text += "\n"

    prompt = f"""你是一个高校教学 PPT 生成助手。请根据用户的需求和提供的参考文件内容，生成一份完整的 slides_config.json。

{schema}

## 用户需求

{user_request}

{ref_text}
## 输出要求

1. 只输出合法的 JSON，用 ```json ``` 包裹
2. meta.output_name 使用英文（作为文件名）
3. 每页根据内容选择合适的 type：
   - 需要流程架构图/原理图/示意图 → type: "dalle"
   - 文字+案例 → type: "content"
   - 多概念并列比较 → type: "cards"
   - 表格对比 → type: "table"
   - 左右对照（优缺点/不要vs要） → type: "comparison"
   - 编号问答总结 → type: "summary"
4. 约30%的页设为 "dalle" type（关键视觉页）
5. info_boxes 格式: [left, top, width, height, [lines], bg_color, text_color, font_size]
6. cards 格式: [left, top, width, height, title, [lines], icon, color]
7. 颜色使用常量名: "ACC_BLUE", "GREEN_OK", "WARM_ACC", "RED_WARN", "PURPLE"
8. dalle_desc 用中文描述要生成的图像（100字以内）

## 示例参考（20页配置的片段）

{example_json[:3000] if example_json else ""}

请生成完整的 slides_config.json："""
    return prompt

# ============================================================
# 4. 调用 LLM API
# ============================================================
def call_llm(prompt, provider="anthropic"):
    """调用 LLM API 生成 JSON 配置"""
    api_key = None

    if provider == "anthropic":
        api_key = os.environ.get("ANTHROPIC_API_KEY")
        if not api_key:
            raise ValueError("请设置 ANTHROPIC_API_KEY 环境变量")
        return call_anthropic(prompt, api_key)

    elif provider == "openai":
        api_key = os.environ.get("OPENAI_API_KEY")
        if not api_key:
            raise ValueError("请设置 OPENAI_API_KEY 环境变量")
        return call_openai(prompt, api_key)

    elif provider == "deepseek":
        api_key = os.environ.get("DEEPSEEK_API_KEY")
        if not api_key:
            raise ValueError("请设置 DEEPSEEK_API_KEY 环境变量")
        return call_deepseek(prompt, api_key)

    else:
        raise ValueError(f"Unknown provider: {provider}")

def call_anthropic(prompt, api_key):
    """调用 Anthropic Claude API"""
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
        print("  🤖 正在调用 Claude API...")
        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=8192,
            messages=[{"role": "user", "content": prompt}]
        )
        return message.content[0].text
    except ImportError:
        print("  ⚠️  未安装 anthropic SDK, 尝试 HTTP 调用...")
        return call_anthropic_http(prompt, api_key)

def call_anthropic_http(prompt, api_key):
    """HTTP 方式调用 Anthropic API (无需 SDK)"""
    import urllib.request, urllib.error
    url = "https://api.anthropic.com/v1/messages"
    data = json.dumps({
        "model": "claude-sonnet-4-6",
        "max_tokens": 8192,
        "messages": [{"role": "user", "content": prompt}]
    }).encode('utf-8')

    req = urllib.request.Request(url, data=data, method='POST')
    req.add_header('x-api-key', api_key)
    req.add_header('anthropic-version', '2023-06-01')
    req.add_header('content-type', 'application/json')

    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            result = json.loads(resp.read())
            return result['content'][0]['text']
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"API 调用失败: {e.code} {e.reason}")

def call_openai(prompt, api_key):
    """调用 OpenAI API"""
    import urllib.request, urllib.error
    url = "https://api.openai.com/v1/chat/completions"
    data = json.dumps({
        "model": "gpt-4.1-mini",
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": 8192
    }).encode('utf-8')

    req = urllib.request.Request(url, data=data, method='POST')
    req.add_header('Authorization', f'Bearer {api_key}')
    req.add_header('Content-Type', 'application/json')

    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            result = json.loads(resp.read())
            return result['choices'][0]['message']['content']
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"API 调用失败: {e.code} {e.reason}")

def call_deepseek(prompt, api_key):
    """调用 DeepSeek API"""
    import urllib.request, urllib.error
    url = "https://api.deepseek.com/v1/chat/completions"
    data = json.dumps({
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": 8192
    }).encode('utf-8')

    req = urllib.request.Request(url, data=data, method='POST')
    req.add_header('Authorization', f'Bearer {api_key}')
    req.add_header('Content-Type', 'application/json')

    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            result = json.loads(resp.read())
            return result['choices'][0]['message']['content']
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"API 调用失败: {e.code} {e.reason}")

# ============================================================
# 5. 解析 LLM 返回的 JSON
# ============================================================
def extract_json(response_text):
    """从 LLM 回复中提取 JSON"""
    # 尝试找 ```json ``` 包裹的内容
    m = re.search(r'```json\s*(.*?)\s*```', response_text, re.DOTALL)
    if m:
        return m.group(1)
    # 尝试找 ``` ``` 包裹的内容
    m = re.search(r'```\s*(\{.*?\})\s*```', response_text, re.DOTALL)
    if m:
        return m.group(1)
    # 尝试直接找 JSON 对象
    m = re.search(r'\{.*"slides".*\}', response_text, re.DOTALL)
    if m:
        return m.group(0)
    return response_text

# ============================================================
# 6. 主函数
# ============================================================
def main():
    parser = argparse.ArgumentParser(
        description='PPT Agent — 自然语言驱动的PPT全自动流水线',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python3 ppt_agent.py "做一个关于数据交易合规的PPT，20页"
  python3 ppt_agent.py "数据要素X案例研讨，15页" --refs ~/Downloads/refs/
  python3 ppt_agent.py "..." --provider openai
  python3 ppt_agent.py "..." --provider deepseek
  python3 ppt_agent.py "..." --dry-run  # 只生成配置，不渲染PPT
        """
    )
    parser.add_argument('request', help='自然语言需求描述')
    parser.add_argument('--refs', nargs='+', default=[], help='参考PPTX文件或目录路径')
    parser.add_argument('--provider', default='anthropic', choices=['anthropic', 'openai', 'deepseek'], help='LLM提供商 (默认: anthropic)')
    parser.add_argument('--pages', type=int, default=20, help='期望页数 (默认: 20)')
    parser.add_argument('--dry-run', action='store_true', help='只生成配置JSON，不渲染PPT')
    parser.add_argument('--output', default=None, help='输出文件名 (不含扩展名)')

    args = parser.parse_args()

    print("=" * 60)
    print("📊 PPT Agent — 全自动 PPT 生产流水线")
    print("=" * 60)

    # Step 1: 读取参考文件
    print("\n📂 Step 1: 读取参考文件...")
    refs = []
    if args.refs:
        refs = read_reference_files(args.refs)
        if not refs:
            print("  (未找到 PPTX 文件, 继续...)")

    # Step 2: 构建 prompt
    print("\n📝 Step 2: 构建 Agent prompt...")
    schema, example = load_schema_and_example()
    user_req = args.request
    if args.pages:
        user_req = f"{args.request}\n\n(期望页数: {args.pages}页)"
    prompt = build_agent_prompt(user_req, refs, schema, example)
    print(f"  Prompt 长度: {len(prompt)} 字符")

    # Step 3: 调用 LLM
    print(f"\n🧠 Step 3: 调用 {args.provider.upper()} API...")
    try:
        response = call_llm(prompt, args.provider)
        print(f"  回复长度: {len(response)} 字符")
    except Exception as e:
        print(f"\n❌ LLM 调用失败: {e}")
        print("\n💡 请检查:")
        print("  1. API Key 是否已设置 (export XX_API_KEY=...)")
        print("  2. 网络是否能访问 API")
        print("  3. API 账户是否有余额")
        print("\n💡 也可以使用 --dry-run 先生成 prompt，手动复制到 Claude 对话中生成 JSON")
        sys.exit(1)

    # Step 4: 提取并保存 JSON
    print("\n💾 Step 4: 解析 JSON 配置...")
    json_str = extract_json(response)
    try:
        config = json.loads(json_str)
    except json.JSONDecodeError:
        print("  ⚠️  JSON 解析失败, 保存原始回复供检查")
        with open('llm_response.txt', 'w', encoding='utf-8') as f:
            f.write(response)
        print("  原始回复已保存到: llm_response.txt")
        sys.exit(1)

    out_name = args.output or config.get('meta', {}).get('output_name', 'output')
    config_path = f'slides_{out_name}.json'
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    print(f"  配置已保存: {config_path}")

    if args.dry_run:
        print(f"\n✅ Dry-run 完成! 配置: {config_path}")
        return

    # Step 5: 调用 make_ppt.py
    print("\n🎨 Step 5: 渲染 PPTX + 生成 DALL-E prompts...")
    engine_path = Path(__file__).parent / 'make_ppt.py'
    if not engine_path.exists():
        print(f"  ❌ 找不到 make_ppt.py, 请确保它在同一目录")
        sys.exit(1)

    import subprocess
    result = subprocess.run([sys.executable, str(engine_path), config_path], capture_output=True, text=True)
    print(result.stdout)
    if result.returncode != 0:
        print(f"  ⚠️  渲染警告:\n{result.stderr}")

    print("\n" + "=" * 60)
    print("🎉 全部完成!")
    dalle_count = sum(1 for s in config.get('slides', []) if s.get('type') == 'dalle')
    print(f"   总页数: {len(config.get('slides', []))} | DALL-E 页: {dalle_count} | 文字页: {len(config.get('slides', [])) - dalle_count}")
    print("=" * 60)

if __name__ == '__main__':
    main()
