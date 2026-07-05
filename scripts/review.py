#!/usr/bin/env python3
"""
发表前自查脚本
检查数据口径一致性、图分辨率、页数上限等
"""

import argparse
import json
import sys
from pathlib import Path

try:
    import yaml
    from pptx import Presentation
except ImportError as e:
    print(json.dumps({
        "status": "error",
        "message": f"依赖缺失: {e}，请运行: pip install python-pptx pyyaml"
    }, ensure_ascii=False))
    sys.exit(1)


def check_data_consistency(project_root):
    """检查数据口径一致性"""
    issues = []
    
    # 读取主档案
    main_file = project_root / "00_项目主档案.md"
    if not main_file.exists():
        issues.append({"level": "❌", "item": "主档案缺失", "suggestion": "请补充00_项目主档案.md"})
        return issues
    
    # 简单检查现状值/目标值/实绩值是否存在
    with open(main_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    if '现状' not in content and 'baseline' not in content.lower():
        issues.append({"level": "⚠️", "item": "主档案中未找到现状值", "suggestion": "请补充现状数据"})
    
    if '目标' not in content and 'target' not in content.lower():
        issues.append({"level": "⚠️", "item": "主档案中未找到目标值", "suggestion": "请补充目标数据"})
    
    return issues


def check_image_resolution(project_root):
    """检查图片分辨率"""
    issues = []
    
    try:
        from PIL import Image
    except ImportError:
        issues.append({"level": "⚠️", "item": "Pillow未安装，跳过图片检查", "suggestion": "pip install Pillow"})
        return issues
    
    # 扫描所有图片
    for img_file in project_root.rglob("*"):
        if img_file.suffix.lower() in ['.png', '.jpg', '.jpeg']:
            try:
                img = Image.open(img_file)
                dpi = img.info.get('dpi', (0, 0))
                if dpi[0] < 300 or dpi[1] < 300:
                    issues.append({
                        "level": "⚠️",
                        "item": f"图片分辨率不足: {img_file.name} ({dpi[0]}x{dpi[1]}dpi)",
                        "suggestion": "请使用300dpi以上的高清图片"
                    })
            except Exception as e:
                issues.append({"level": "⚠️", "item": f"无法读取图片: {img_file.name}", "suggestion": str(e)})
    
    return issues


def check_pptx(ppt_path):
    """检查PPT"""
    issues = []
    
    if not ppt_path.exists():
        issues.append({"level": "❌", "item": "PPT文件不存在", "suggestion": "请先生成PPT"})
        return issues
    
    try:
        prs = Presentation(str(ppt_path))
        total_pages = len(prs.slides)
        
        # 检查页数
        if total_pages > 100:
            issues.append({
                "level": "❌",
                "item": f"总页数 {total_pages} 超过100页上限",
                "suggestion": "请精简PPT内容"
            })
        
        # 检查必备元素
        # 简单检查第一页和最后一页
        if total_pages >= 2:
            first_slide = prs.slides[0]
            last_slide = prs.slides[-1]
            
            # 检查封面是否有标题
            has_title = False
            for shape in first_slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    has_title = True
                    break
            if not has_title:
                issues.append({"level": "⚠️", "item": "封面页缺少标题", "suggestion": "请在封面添加课题名称"})
        
        issues.append({"level": "✅", "item": f"PPT页数: {total_pages}", "suggestion": ""})
        
    except Exception as e:
        issues.append({"level": "❌", "item": f"PPT解析失败: {e}", "suggestion": "请检查PPT文件格式"})
    
    return issues


def check_speech(speech_path, ppt_path):
    """检查讲稿"""
    issues = []
    
    if not speech_path.exists():
        issues.append({"level": "⚠️", "item": "讲稿文件不存在", "suggestion": "请先生成讲稿"})
        return issues
    
    with open(speech_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 检查关键数字
    if '现状' not in content and '目标' not in content:
        issues.append({"level": "⚠️", "item": "讲稿中未找到关键数据", "suggestion": "请确保讲稿包含现状值、目标值、实绩值"})
    
    issues.append({"level": "✅", "item": "讲稿文件存在", "suggestion": ""})
    
    return issues


def generate_review_report(project_root, ppt_path, speech_path, output_path):
    """生成自查报告"""
    all_issues = []
    
    # 数据一致性检查
    all_issues.extend(check_data_consistency(project_root))
    
    # 图片分辨率检查
    all_issues.extend(check_image_resolution(project_root))
    
    # PPT检查
    if ppt_path:
        all_issues.extend(check_pptx(ppt_path))
    
    # 讲稿检查
    if speech_path:
        all_issues.extend(check_speech(speech_path, ppt_path))
    
    # 生成报告
    report = []
    report.append("# QCC发表前自查报告\n")
    
    # 统计
    error_count = sum(1 for i in all_issues if "❌" in i["level"])
    warning_count = sum(1 for i in all_issues if "⚠️" in i["level"])
    pass_count = sum(1 for i in all_issues if "✅" in i["level"])
    
    report.append(f"**统计**: 通过 {pass_count} | 警告 {warning_count} | 不通过 {error_count}\n")
    report.append("---\n")
    
    # 阻断项
    if error_count > 0:
        report.append("## ❌ 阻断项（必须修复）\n")
        for issue in all_issues:
            if "❌" in issue["level"]:
                report.append(f"- **{issue['item']}**\n  - 建议: {issue['suggestion']}")
        report.append("\n")
    
    # 警告项
    if warning_count > 0:
        report.append("## ⚠️ 警告项（建议修复）\n")
        for issue in all_issues:
            if "⚠️" in issue["level"]:
                report.append(f"- **{issue['item']}**\n  - 建议: {issue['suggestion']}")
        report.append("\n")
    
    # 通过项
    report.append("## ✅ 通过项\n")
    for issue in all_issues:
        if "✅" in issue["level"]:
            report.append(f"- {issue['item']}")
    report.append("\n")
    
    # 结论
    report.append("---\n")
    report.append("## 结论\n")
    if error_count > 0:
        report.append("**⛔ 不通过**: 存在阻断项，请修复后重新检查。\n")
    elif warning_count > 0:
        report.append("**⚠️ 有条件通过**: 建议修复警告项后发表。\n")
    else:
        report.append("**✅ 通过**: 可以发表。\n")
    
    # 写入文件
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(report))
    
    return error_count, warning_count, pass_count


def main():
    parser = argparse.ArgumentParser(description='QCC发表前自查')
    parser.add_argument('--project-root', required=True, help='项目根目录路径')
    parser.add_argument('--ppt', help='PPT文件路径')
    parser.add_argument('--speech', help='讲稿文件路径')
    parser.add_argument('--output', default='自查报告.md', help='输出报告路径')
    
    args = parser.parse_args()
    
    project_root = Path(args.project_root)
    if not project_root.exists():
        print(json.dumps({
            "status": "error",
            "message": f"项目目录不存在: {args.project_root}"
        }, ensure_ascii=False))
        sys.exit(2)
    
    ppt_path = Path(args.ppt) if args.ppt else None
    speech_path = Path(args.speech) if args.speech else None
    output_path = Path(args.output)
    
    # 生成报告
    error_count, warning_count, pass_count = generate_review_report(
        project_root, ppt_path, speech_path, output_path
    )
    
    # 输出结果
    result = {
        "status": "success",
        "message": f"自查报告已生成: {output_path}",
        "data": {
            "error_count": error_count,
            "warning_count": warning_count,
            "pass_count": pass_count,
            "blocked": error_count > 0
        }
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    
    # 退出码
    if error_count > 0:
        sys.exit(2)
    elif warning_count > 0:
        sys.exit(1)
    else:
        sys.exit(0)


if __name__ == '__main__':
    main()
