#!/usr/bin/env python3
"""
导出归档脚本
将PPT、讲稿、关键产物打包归档
"""

import argparse
import json
import sys
import zipfile
from pathlib import Path
from datetime import datetime


def collect_files(project_root, ppt_path, speech_path):
    """收集需要归档的文件"""
    files = []
    
    # PPT
    if ppt_path and ppt_path.exists():
        files.append(("发表PPT", ppt_path))
    
    # 讲稿
    if speech_path and speech_path.exists():
        files.append(("发表讲稿", speech_path))
    
    # 主档案
    main_file = project_root / "00_项目主档案.md"
    if main_file.exists():
        files.append(("项目主档案", main_file))
    
    # 关键步骤产物
    key_steps = [
        ("03_Step3_现状把握", "现状把握产物"),
        ("06_Step6_要因验证", "要因验证产物"),
        ("07_Step7_对策拟定", "对策拟定产物"),
        ("09_Step9_效果确认", "效果确认产物")
    ]
    
    for step_dir, desc in key_steps:
        step_path = project_root / step_dir
        if step_path.exists():
            for f in step_path.iterdir():
                if f.is_file() and f.suffix.lower() in ['.md', '.png', '.jpg', '.csv', '.xlsx']:
                    files.append((f"{desc}/{f.name}", f))
    
    return files


def generate_archive_manifest(files, output_path):
    """生成归档清单"""
    manifest = []
    manifest.append("# QCC发表归档清单\n")
    manifest.append(f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    manifest.append("---\n")
    
    manifest.append("## 归档文件列表\n")
    manifest.append("| 序号 | 文件说明 | 文件路径 |")
    manifest.append("|------|----------|----------|")
    
    for i, (desc, path) in enumerate(files, 1):
        manifest.append(f"| {i} | {desc} | {path} |")
    
    manifest.append(f"\n**总计**: {len(files)} 个文件\n")
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(manifest))


def create_archive(files, manifest_path, output_path):
    """创建归档zip包"""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        # 添加归档清单
        zf.write(manifest_path, "归档清单.md")
        
        # 添加文件
        for desc, path in files:
            try:
                arcname = f"发表归档/{desc}"
                zf.write(path, arcname)
            except Exception as e:
                print(f"警告: 无法添加文件 {path}: {e}", file=sys.stderr)


def main():
    parser = argparse.ArgumentParser(description='导出QCC发表归档包')
    parser.add_argument('--project-root', required=True, help='项目根目录路径')
    parser.add_argument('--ppt', help='PPT文件路径')
    parser.add_argument('--speech', help='讲稿文件路径')
    parser.add_argument('--output', default='发表归档.zip', help='输出zip路径')
    
    args = parser.parse_args()
    
    project_root = Path(args.project_root)
    if not project_root.exists():
        print(json.dumps({
            "status": "error",
            "message": f"项目目录不存在: {args.project_root}"
        }, ensure_ascii=False))
        sys.exit(1)
    
    ppt_path = Path(args.ppt) if args.ppt else None
    speech_path = Path(args.speech) if args.speech else None
    
    # 收集文件
    files = collect_files(project_root, ppt_path, speech_path)
    
    if not files:
        print(json.dumps({
            "status": "error",
            "message": "未找到任何可归档的文件"
        }, ensure_ascii=False))
        sys.exit(1)
    
    # 生成归档清单
    manifest_path = Path("归档清单.md")
    generate_archive_manifest(files, manifest_path)
    
    # 创建zip包
    output_path = Path(args.output)
    create_archive(files, manifest_path, output_path)
    
    # 清理临时清单
    manifest_path.unlink()
    
    # 输出结果
    result = {
        "status": "success",
        "message": f"归档包已生成: {output_path}",
        "data": {
            "total_files": len(files),
            "output_file": str(output_path)
        }
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
