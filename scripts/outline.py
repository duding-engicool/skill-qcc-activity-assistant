#!/usr/bin/env python3
"""
PPT大纲生成脚本
按QCC大赛标准结构生成PPT大纲YAML
"""

import argparse
import json
import sys
import os
from pathlib import Path
from datetime import datetime

try:
    import yaml
except ImportError:
    print(json.dumps({
        "status": "error",
        "message": "pyyaml未安装，请运行: pip install pyyaml"
    }, ensure_ascii=False))
    sys.exit(1)


# 14章节标准结构
CHAPTERS_14 = [
    {"id": 1, "name": "封面", "required": ["课题名", "圈名", "发表人", "日期"]},
    {"id": 2, "name": "圈情介绍", "required": ["圈名由来", "圈徽", "成员介绍", "活动历程"]},
    {"id": 3, "name": "主题选定", "required": ["选题理由", "评价表", "选定主题"]},
    {"id": 4, "name": "活动计划", "required": ["甘特图", "进度安排"]},
    {"id": 5, "name": "现状把握", "required": ["现状数据", "柏拉图", "层别分析"]},
    {"id": 6, "name": "目标设定", "required": ["目标值", "设定依据", "目标图"]},
    {"id": 7, "name": "原因分析", "required": ["鱼骨图", "原因清单"]},
    {"id": 8, "name": "要因验证", "required": ["验证数据", "要因确认表"]},
    {"id": 9, "name": "对策拟定", "required": ["对策方案", "5W1H表", "评价表"]},
    {"id": 10, "name": "对策实施", "required": ["实施过程", "实施照片", "阶段性成果"]},
    {"id": 11, "name": "效果确认", "required": ["效果对比", "目标达成", "经济效益"]},
    {"id": 12, "name": "标准化", "required": ["标准化文件", "作业指导书"]},
    {"id": 13, "name": "检讨与改进", "required": ["活动检讨", "遗留问题", "未来计划"]},
    {"id": 14, "name": "致谢", "required": ["致谢词"]}
]

# 8章节精简版合并规则
CHAPTERS_8 = [
    {"id": 1, "name": "封面", "merge_from": [1]},
    {"id": 2, "name": "圈情与计划", "merge_from": [2, 4]},
    {"id": 3, "name": "主题选定与现状", "merge_from": [3, 5]},
    {"id": 4, "name": "目标与原因", "merge_from": [6, 7]},
    {"id": 5, "name": "要因验证与对策", "merge_from": [8, 9]},
    {"id": 6, "name": "实施与效果", "merge_from": [10, 11]},
    {"id": 7, "name": "标准化与检讨", "merge_from": [12, 13]},
    {"id": 8, "name": "致谢", "merge_from": [14]}
]


def load_project_info(project_root):
    """加载项目基本信息"""
    info = {
        "project_name": "QCC课题",
        "circle_name": "品管圈",
        "presenter": "发表人",
        "baseline_value": "N/A",
        "target_value": "N/A",
        "actual_value": "N/A",
        "improvement_rate": "N/A"
    }
    
    main_file = project_root / "00_项目主档案.md"
    if main_file.exists():
        try:
            with open(main_file, 'r', encoding='utf-8') as f:
                content = f.read()
                # 简单解析关键字段
                for line in content.split('\n'):
                    if '课题' in line or '主题' in line:
                        info["project_name"] = line.split(':')[-1].strip() if ':' in line else line.strip()
                    elif '圈名' in line:
                        info["circle_name"] = line.split(':')[-1].strip() if ':' in line else line.strip()
                    elif '现状' in line and ('值' in line or '%' in line):
                        info["baseline_value"] = line.split(':')[-1].strip() if ':' in line else line.strip()
                    elif '目标' in line and ('值' in line or '%' in line):
                        info["target_value"] = line.split(':')[-1].strip() if ':' in line else line.strip()
        except:
            pass
    
    return info


def find_images_in_step(project_root, step_dir):
    """查找步骤目录中的图片，返回带analysis占位符的图片信息"""
    images = []
    step_path = project_root / step_dir
    if step_path.exists():
        for f in step_path.iterdir():
            if f.suffix.lower() in ['.png', '.jpg', '.jpeg']:
                images.append({
                    "path": str(f.relative_to(project_root)),
                    "position": "right",
                    "width_cm": 14,
                    "analysis": ""  # 占位符，由智能体填充图片解析内容
                })
    return images


def generate_outline(project_root, template_type, target_duration):
    """生成PPT大纲"""
    info = load_project_info(project_root)
    chapters = CHAPTERS_14 if template_type == "完整版" else CHAPTERS_8
    
    pages = []
    page_no = 1
    
    for chapter in chapters:
        # 封面页
        if chapter["id"] == 1 or (template_type == "精简版" and chapter["id"] == 1):
            pages.append({
                "page_no": page_no,
                "title": info["project_name"],
                "layout": "cover",
                "blocks": [
                    {
                        "type": "paragraph",
                        "font_size": 24,
                        "content": info["project_name"]
                    },
                    {
                        "type": "paragraph",
                        "font_size": 16,
                        "content": f"圈名: {info['circle_name']}"
                    }
                ],
                "images": [],
                "speaker_notes_outline": f"介绍课题名称和圈名",
                "est_seconds": 30
            })
            page_no += 1
        
        # 章节过渡页
        pages.append({
            "page_no": page_no,
            "title": chapter["name"],
            "layout": "section_break",
            "blocks": [
                {
                    "type": "paragraph",
                    "font_size": 20,
                    "content": f"第{chapter['id']}章 {chapter['name']}"
                }
            ],
            "images": [],
            "speaker_notes_outline": f"进入{chapter['name']}章节",
            "est_seconds": 15
        })
        page_no += 1
        
        # 内容页（根据章节生成）
        if chapter["id"] == 5:  # 现状把握
            images = find_images_in_step(project_root, "03_Step3_现状把握")
            pages.append({
                "page_no": page_no,
                "title": "现状把握 · 数据分析",
                "layout": "text_left_image_right",
                "blocks": [
                    {
                        "type": "bullet",
                        "font_size": 16,
                        "items": [
                            f"现状值: {info['baseline_value']}",
                            "数据收集周期: 30天",
                            "样本量: N=500+"
                        ]
                    }
                ],
                "images": images[:2],  # images已经是字典列表
                "speaker_notes_outline": "说明现状数据收集方法和主要发现",
                "est_seconds": 90
            })
            page_no += 1
        
        elif chapter["id"] == 6:  # 目标设定
            pages.append({
                "page_no": page_no,
                "title": "目标设定",
                "layout": "text_single",
                "blocks": [
                    {
                        "type": "bullet",
                        "font_size": 16,
                        "items": [
                            f"现状值: {info['baseline_value']}",
                            f"目标值: {info['target_value']}",
                            "目标设定依据: SMART原则"
                        ]
                    }
                ],
                "images": [],
                "speaker_notes_outline": "说明目标值设定依据",
                "est_seconds": 60
            })
            page_no += 1
        
        elif chapter["id"] == 11:  # 效果确认
            images = find_images_in_step(project_root, "09_Step9_效果确认")
            pages.append({
                "page_no": page_no,
                "title": "效果确认",
                "layout": "text_left_image_right",
                "blocks": [
                    {
                        "type": "bullet",
                        "font_size": 16,
                        "items": [
                            f"改善前: {info['baseline_value']}",
                            f"改善后: {info['actual_value']}",
                            f"目标达成率: {info['improvement_rate']}"
                        ]
                    }
                ],
                "images": images[:2],  # images已经是字典列表
                "speaker_notes_outline": "对比改善前后数据，说明目标达成情况",
                "est_seconds": 120
            })
            page_no += 1
    
    # 致谢页
    pages.append({
        "page_no": page_no,
        "title": "致谢",
        "layout": "acknowledgement",
        "blocks": [
            {
                "type": "paragraph",
                "font_size": 18,
                "content": "感谢各位评委的聆听与指导！"
            }
        ],
        "images": [],
        "speaker_notes_outline": "致谢结束",
        "est_seconds": 15
    })
    
    # 计算总时长
    total_seconds = sum(p["est_seconds"] for p in pages)
    
    outline = {
        "metadata": {
            "project_name": info["project_name"],
            "circle_name": info["circle_name"],
            "template": template_type,
            "target_duration_minutes": target_duration,
            "estimated_duration_seconds": total_seconds,
            "total_pages": len(pages),
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        },
        "pages": pages
    }
    
    return outline


def generate_review_markdown(outline, output_path):
    """生成大纲检视Markdown"""
    md = []
    md.append("# PPT大纲检视\n")
    md.append(f"**课题**: {outline['metadata']['project_name']}\n")
    md.append(f"**圈名**: {outline['metadata']['circle_name']}\n")
    md.append(f"**模板**: {outline['metadata']['template']}\n")
    md.append(f"**总页数**: {outline['metadata']['total_pages']}\n")
    md.append(f"**预估时长**: {outline['metadata']['estimated_duration_seconds']//60}分{outline['metadata']['estimated_duration_seconds']%60}秒\n")
    md.append("---\n")
    
    md.append("## 目录结构\n")
    for page in outline["pages"]:
        md.append(f"- P{page['page_no']}: {page['title']} ({page['layout']})")
    
    md.append("\n---\n")
    
    # 提醒
    if outline["metadata"]["total_pages"] > 100:
        md.append("## ⚠️ 提醒\n")
        md.append(f"总页数 {outline['metadata']['total_pages']} 超过100页上限，建议精简。\n")
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(md))


def main():
    parser = argparse.ArgumentParser(description='生成QCC发表PPT大纲')
    parser.add_argument('--project-root', required=True, help='项目根目录路径')
    parser.add_argument('--template', choices=['完整版', '精简版'], default='完整版', help='大纲模板')
    parser.add_argument('--target-duration', type=int, default=15, help='目标时长(分钟)')
    parser.add_argument('--output', default='发表PPT大纲.yaml', help='输出YAML路径')
    parser.add_argument('--review-output', default='大纲检视.md', help='输出检视Markdown路径')
    
    args = parser.parse_args()
    
    project_root = Path(args.project_root)
    if not project_root.exists():
        print(json.dumps({
            "status": "error",
            "message": f"项目目录不存在: {args.project_root}"
        }, ensure_ascii=False))
        sys.exit(1)
    
    # 生成大纲
    outline = generate_outline(project_root, args.template, args.target_duration)
    
    # 保存YAML
    output_path = Path(args.output)
    with open(output_path, 'w', encoding='utf-8') as f:
        yaml.dump(outline, f, allow_unicode=True, default_flow_style=False, sort_keys=False)
    
    # 生成检视
    generate_review_markdown(outline, args.review_output)
    
    # 输出结果
    result = {
        "status": "success",
        "message": f"大纲已生成: {output_path}",
        "data": {
            "total_pages": outline["metadata"]["total_pages"],
            "estimated_duration_seconds": outline["metadata"]["estimated_duration_seconds"],
            "template": args.template
        }
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
