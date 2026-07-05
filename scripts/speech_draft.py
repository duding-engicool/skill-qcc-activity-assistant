#!/usr/bin/env python3
"""
讲稿生成脚本
为每页生成讲稿骨架
"""

import argparse
import json
import sys
from pathlib import Path

try:
    import yaml
except ImportError:
    print(json.dumps({
        "status": "error",
        "message": "pyyaml未安装，请运行: pip install pyyaml"
    }, ensure_ascii=False))
    sys.exit(1)


# 转折语库
TRANSITIONS = {
    "opening": "各位评委、各位同事，大家好。今天由我代表{circle_name}汇报我们的QCC课题。",
    "next_section": "接下来，让我们进入{section}部分。",
    "data_intro": "首先来看数据。",
    "conclusion": "由此我们可以得出结论：",
    "transition": "这就引出了我们的下一个问题——",
    "closing": "以上就是我们的汇报内容，感谢各位评委的聆听与指导！"
}


def generate_speech_for_page(page_data, prev_page=None):
    """为单页生成讲稿"""
    title = page_data.get("title", "")
    layout = page_data.get("layout", "")
    blocks = page_data.get("blocks", [])
    est_seconds = page_data.get("est_seconds", 60)
    
    speech_parts = []
    
    # 转折/衔接
    if prev_page is None:
        speech_parts.append(f"**【开场】** {TRANSITIONS['opening'].format(circle_name='品管圈')}")
    elif layout == "section_break":
        speech_parts.append(f"**【章节过渡】** {TRANSITIONS['next_section'].format(section=title)}")
    else:
        speech_parts.append(f"**【衔接】** {TRANSITIONS['transition']}")
    
    # 数据念读
    for block in blocks:
        if block.get("type") == "bullet":
            items = block.get("items", [])
            if items:
                speech_parts.append("**【数据念读】**")
                for item in items:
                    # 高亮数字
                    speech_parts.append(f"  - {item}")
        elif block.get("type") == "paragraph":
            content = block.get("content", "")
            speech_parts.append(f"**【要点】** {content}")
        elif block.get("type") == "table":
            speech_parts.append("**【表格说明】** 请看表格中的数据...")
    
    # 关键判断
    if "效果" in title or "确认" in title:
        speech_parts.append(f"**【关键结论】** {TRANSITIONS['conclusion']}目标达成情况如数据所示。")
    elif "目标" in title:
        speech_parts.append(f"**【目标说明】** 我们的目标值是基于SMART原则设定的。")
    
    # 控时标注
    speech_parts.append(f"\n*[预计时长: {est_seconds}秒]*")
    
    return "\n".join(speech_parts)


def generate_speech_draft(outline_path, target_duration, speaker_style):
    """生成完整讲稿"""
    with open(outline_path, 'r', encoding='utf-8') as f:
        outline = yaml.safe_load(f)
    
    pages = outline.get("pages", [])
    metadata = outline.get("metadata", {})
    
    draft = []
    draft.append(f"# QCC发表讲稿\n")
    draft.append(f"**课题**: {metadata.get('project_name', 'N/A')}\n")
    draft.append(f"**圈名**: {metadata.get('circle_name', 'N/A')}\n")
    draft.append(f"**目标时长**: {target_duration}分钟\n")
    draft.append(f"**风格**: {speaker_style}\n")
    draft.append("---\n")
    
    # 控时表
    draft.append("## 控时表\n")
    draft.append("| 页码 | 标题 | 预计时长 | 累计时间 |")
    draft.append("|------|------|----------|----------|")
    
    cumulative = 0
    for page in pages:
        est = page.get("est_seconds", 60)
        cumulative += est
        draft.append(f"| P{page.get('page_no')} | {page.get('title', '')} | {est}秒 | {cumulative//60}分{cumulative%60}秒 |")
    
    draft.append(f"\n**总时长**: {cumulative//60}分{cumulative%60}秒\n")
    draft.append("---\n")
    
    # 逐页讲稿
    draft.append("## 逐页讲稿\n")
    
    prev_page = None
    for page in pages:
        draft.append(f"### P{page.get('page_no')}: {page.get('title', '')}\n")
        speech = generate_speech_for_page(page, prev_page)
        draft.append(speech)
        draft.append("\n---\n")
        prev_page = page
    
    # 结尾
    draft.append("## 结尾\n")
    draft.append(f"**【结束语】** {TRANSITIONS['closing']}")
    
    return "\n".join(draft), cumulative


def main():
    parser = argparse.ArgumentParser(description='生成QCC发表讲稿')
    parser.add_argument('--outline', required=True, help='大纲YAML路径')
    parser.add_argument('--target-duration', type=int, default=15, help='目标时长(分钟)')
    parser.add_argument('--speaker-style', choices=['稳重', '平实', '有节奏感'], default='平实', help='演讲风格')
    parser.add_argument('--output', default='发表讲稿.md', help='输出讲稿路径')
    
    args = parser.parse_args()
    
    outline_path = Path(args.outline)
    if not outline_path.exists():
        print(json.dumps({
            "status": "error",
            "message": f"大纲文件不存在: {args.outline}"
        }, ensure_ascii=False))
        sys.exit(1)
    
    # 生成讲稿
    draft, total_seconds = generate_speech_draft(outline_path, args.target_duration, args.speaker_style)
    
    # 保存
    output_path = Path(args.output)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(draft)
    
    # 输出结果
    result = {
        "status": "success",
        "message": f"讲稿已生成: {output_path}",
        "data": {
            "total_duration_seconds": total_seconds,
            "target_duration_seconds": args.target_duration * 60,
            "deviation_percent": round((total_seconds - args.target_duration * 60) / (args.target_duration * 60) * 100, 1)
        }
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
