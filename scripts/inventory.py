#!/usr/bin/env python3
"""
项目产物盘点脚本
扫描QCC项目目录，对照标准产物清单识别缺失项
"""

import argparse
import json
import sys
import os
from pathlib import Path
from datetime import datetime


# 标准产物清单（10步骤）
STANDARD_PRODUCTS = {
    "01_Step1_主题选定": ["选题理由.md", "主题选定评价表.md"],
    "02_Step2_活动计划": ["活动计划表.md", "甘特图.md"],
    "03_Step3_现状把握": ["现状数据.md", "柏拉图.png", "层别分析.md"],
    "04_Step4_目标设定": ["目标设定.md", "目标值计算.md"],
    "05_Step5_原因分析": ["鱼骨图.png", "原因分析.md"],
    "06_Step6_要因验证": ["要因验证表.md", "验证数据.md"],
    "07_Step7_对策拟定": ["对策计划表.md", "5W1H分析.md"],
    "08_Step8_对策实施": ["实施记录.md", "实施过程照片/"],
    "09_Step9_效果确认": ["效果确认.md", "效果对比图.png"],
    "10_Step10_标准化": ["标准化文件.md", "检讨与改进.md"]
}


def check_file_quality(filepath):
    """检查文件质量"""
    warnings = []
    
    if not filepath.exists():
        return ["❌ 文件不存在"]
    
    # 检查图片分辨率
    if filepath.suffix.lower() in ['.png', '.jpg', '.jpeg']:
        try:
            from PIL import Image
            img = Image.open(filepath)
            dpi = img.info.get('dpi', (0, 0))
            if dpi[0] < 300 or dpi[1] < 300:
                warnings.append(f"⚠️ 图片分辨率不足300dpi (当前: {dpi[0]}x{dpi[1]})")
        except Exception as e:
            warnings.append(f"⚠️ 无法读取图片信息: {e}")
    
    # 检查CSV数据量
    if filepath.suffix.lower() == '.csv':
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                if len(lines) < 125:
                    warnings.append(f"⚠️ CSV样本量不足125行 (当前: {len(lines)}行)")
        except Exception as e:
            warnings.append(f"⚠️ 无法读取CSV: {e}")
    
    return warnings if warnings else ["✅ 通过"]


def scan_project(project_root):
    """扫描项目目录"""
    results = {
        "project_name": "",
        "scan_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "steps": {},
        "missing_count": 0,
        "warning_count": 0,
        "pass_count": 0
    }
    
    # 读取项目主档案
    main_file = project_root / "00_项目主档案.md"
    if main_file.exists():
        try:
            with open(main_file, 'r', encoding='utf-8') as f:
                content = f.read()
                # 简单提取课题名
                for line in content.split('\n'):
                    if '课题' in line or '主题' in line:
                        results["project_name"] = line.split(':')[-1].strip() if ':' in line else line.strip()
                        break
        except:
            pass
    
    # 扫描各步骤目录
    for step_dir, required_files in STANDARD_PRODUCTS.items():
        step_path = project_root / step_dir
        step_result = {
            "exists": step_path.exists(),
            "files": {}
        }
        
        if step_path.exists():
            for req_file in required_files:
                file_path = step_path / req_file
                if req_file.endswith('/'):
                    # 目录检查
                    step_result["files"][req_file] = {
                        "status": "✅ 存在" if file_path.exists() else "❌ 缺失",
                        "warnings": []
                    }
                else:
                    # 文件检查
                    if file_path.exists():
                        warnings = check_file_quality(file_path)
                        status = "✅ 通过" if all("✅" in w for w in warnings) else "⚠️ 警告"
                        step_result["files"][req_file] = {
                            "status": status,
                            "warnings": warnings
                        }
                    else:
                        step_result["files"][req_file] = {
                            "status": "❌ 缺失",
                            "warnings": ["❌ 文件不存在"]
                        }
        
        # 统计
        for file_info in step_result["files"].values():
            if "✅" in file_info["status"]:
                results["pass_count"] += 1
            elif "⚠️" in file_info["status"]:
                results["warning_count"] += 1
            else:
                results["missing_count"] += 1
        
        results["steps"][step_dir] = step_result
    
    return results


def generate_report(results, output_path):
    """生成产物完整性报告"""
    report = []
    report.append("# QCC项目产物完整性报告\n")
    report.append(f"**扫描时间**: {results['scan_time']}\n")
    report.append(f"**课题名称**: {results['project_name'] or '未识别'}\n")
    report.append(f"**统计**: 通过 {results['pass_count']} | 警告 {results['warning_count']} | 缺失 {results['missing_count']}\n")
    report.append("---\n")
    
    # 阻断项汇总
    if results["missing_count"] >= 3:
        report.append("## ❌ 阻断项（缺失≥3，无法继续）\n")
        for step_dir, step_info in results["steps"].items():
            for file_name, file_info in step_info["files"].items():
                if "❌" in file_info["status"]:
                    report.append(f"- {step_dir}/{file_name}")
        report.append("\n")
    
    # 详细清单
    report.append("## 详细产物清单\n")
    report.append("| 步骤 | 产物 | 状态 | 备注 |")
    report.append("|------|------|------|------|")
    
    for step_dir, step_info in results["steps"].items():
        for file_name, file_info in step_info["files"].items():
            status = file_info["status"]
            warnings = "; ".join(file_info["warnings"]) if file_info["warnings"] else "-"
            report.append(f"| {step_dir} | {file_name} | {status} | {warnings} |")
    
    report.append("\n---\n")
    report.append("## 建议\n")
    
    if results["missing_count"] >= 3:
        report.append("**⛔ 阻断**: 缺失产物过多，请先补全后再进行发表准备。\n")
    elif results["warning_count"] > 0:
        report.append("**⚠️ 警告**: 部分产物质量不达标，建议优化后再发表。\n")
    else:
        report.append("**✅ 通过**: 产物完整，可以继续生成发表PPT。\n")
    
    # 写入文件
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(report))
    
    return results["missing_count"]


def main():
    parser = argparse.ArgumentParser(description='QCC项目产物盘点')
    parser.add_argument('--project-root', required=True, help='项目根目录路径')
    parser.add_argument('--output', default='产物完整性报告.md', help='输出报告路径')
    
    args = parser.parse_args()
    
    project_root = Path(args.project_root)
    if not project_root.exists():
        print(json.dumps({
            "status": "error",
            "message": f"项目目录不存在: {args.project_root}"
        }, ensure_ascii=False))
        sys.exit(2)
    
    # 扫描项目
    results = scan_project(project_root)
    
    # 生成报告
    output_path = Path(args.output)
    missing_count = generate_report(results, output_path)
    
    # 输出结果
    result = {
        "status": "success",
        "message": f"产物盘点完成，报告已生成: {output_path}",
        "data": {
            "project_name": results["project_name"],
            "pass_count": results["pass_count"],
            "warning_count": results["warning_count"],
            "missing_count": results["missing_count"],
            "blocked": missing_count >= 3
        }
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    
    # 退出码
    if missing_count >= 3:
        sys.exit(2)
    elif results["warning_count"] > 0:
        sys.exit(1)
    else:
        sys.exit(0)


if __name__ == '__main__':
    main()
