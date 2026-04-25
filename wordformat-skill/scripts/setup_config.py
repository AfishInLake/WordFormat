#!/usr/bin/env python3
"""
WordFormat 配置准备脚本

统一处理配置文件的准备流程：
1. 查找已有预设配置
2. 复制模板生成新配置
3. 验证配置合法性
4. 保存到预设库

用法：
    # 查找已有预设
    python setup_config.py --list-presets

    # 从预设库复制配置
    python setup_config.py --use "清华大学_计算机学院_本科" --output config.yaml

    # 从模板创建新配置（复制模板到工作目录）
    python setup_config.py --create --output config.yaml

    # 验证配置文件
    python setup_config.py --validate --config config.yaml

    # 保存配置到预设库
    python setup_config.py --save --config config.yaml --name "XX大学_XX学院_本科"
"""

import argparse
import os
import shutil
import sys
from pathlib import Path

try:
    import yaml
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyyaml", "--break-system-packages", "-q"])
    import yaml

# 预设目录：当前工作目录下的 presets/
PRESETS_DIR = Path("presets")

# 获取脚本所在目录（用于定位模板和验证脚本）
SCRIPT_DIR = Path(__file__).resolve().parent
TEMPLATE_FILE = SCRIPT_DIR.parent / "data" / "config.yaml"


def get_presets_dir() -> Path:
    """确保预设目录存在"""
    PRESETS_DIR.mkdir(parents=True, exist_ok=True)
    return PRESETS_DIR


def list_presets():
    """列出所有已保存的预设配置"""
    presets_dir = get_presets_dir()
    presets = sorted(presets_dir.glob("*.yaml"))
    if not presets:
        print("暂无已保存的预设配置。")
        print(f"预设目录: {presets_dir}")
        return
    print(f"已保存的预设配置（共 {len(presets)} 个）：")
    print(f"目录: {presets_dir}")
    print("-" * 50)
    for p in presets:
        print(f"  {p.stem}")


def use_preset(name: str, output: str):
    """从预设库复制配置到工作目录"""
    presets_dir = get_presets_dir()
    # 支持不带 .yaml 后缀
    if not name.endswith(".yaml"):
        name = name + ".yaml"
    preset_path = presets_dir / name

    if not preset_path.exists():
        # 模糊匹配：如果用户输入的是部分名称，尝试查找
        matches = [p for p in presets_dir.glob("*.yaml") if name.replace(".yaml", "") in p.stem]
        if len(matches) == 1:
            preset_path = matches[0]
            print(f"模糊匹配到: {preset_path.stem}")
        elif len(matches) > 1:
            print(f"找到多个匹配的预设：")
            for m in matches:
                print(f"  - {m.stem}")
            print("请指定更精确的名称。")
            sys.exit(1)
        else:
            print(f"未找到预设配置: {name.replace('.yaml', '')}")
            print("可用预设：")
            list_presets()
            sys.exit(1)

    shutil.copy2(preset_path, output)
    print(f"已复制预设配置: {preset_path.stem} -> {output}")


def create_from_template(output: str):
    """从模板创建新配置"""
    if not TEMPLATE_FILE.exists():
        print(f"模板文件不存在: {TEMPLATE_FILE}")
        sys.exit(1)
    shutil.copy2(TEMPLATE_FILE, output)
    print(f"已从模板创建配置: {output}")
    print("请根据格式要求编辑此文件，然后运行 --validate 验证。")


def save_preset(config_path: str, name: str):
    """保存配置到预设库"""
    presets_dir = get_presets_dir()
    if not name.endswith(".yaml"):
        name = name + ".yaml"
    dest = presets_dir / name

    if not os.path.exists(config_path):
        print(f"配置文件不存在: {config_path}")
        sys.exit(1)

    # 先验证
    errors = validate(config_path)
    if errors:
        print(f"配置文件有 {len(errors)} 个问题，请先修复后再保存。")
        sys.exit(1)

    shutil.copy2(config_path, dest)
    print(f"已保存预设配置: {dest}")


def validate(config_path: str) -> list[str]:
    """验证配置文件，返回错误列表"""
    # 导入验证逻辑
    validate_script = SCRIPT_DIR / "validate_config.py"
    if validate_script.exists():
        # 动态导入验证函数
        import importlib.util
        spec = importlib.util.spec_from_file_location("validate_config", validate_script)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        try:
            config = yaml.safe_load(open(config_path, "r", encoding="utf-8"))
        except yaml.YAMLError as e:
            return [f"YAML 语法错误: {e}"]
        if not isinstance(config, dict):
            return ["配置文件根节点必须是字典"]
        return mod.validate_config(config)
    else:
        # 简单验证：检查 YAML 语法
        try:
            config = yaml.safe_load(open(config_path, "r", encoding="utf-8"))
        except yaml.YAMLError as e:
            return [f"YAML 语法错误: {e}"]
        return []


def main():
    parser = argparse.ArgumentParser(
        description="WordFormat 配置准备工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python setup_config.py --list-presets
  python setup_config.py --use "清华大学_计算机学院_本科" --output config.yaml
  python setup_config.py --create --output config.yaml
  python setup_config.py --validate --config config.yaml
  python setup_config.py --save --config config.yaml --name "XX大学_XX学院_本科"
"""
    )

    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--list-presets", action="store_true", help="列出所有已保存的预设配置")
    group.add_argument("--use", metavar="NAME", help="从预设库复制配置（支持模糊匹配）")
    group.add_argument("--create", action="store_true", help="从模板创建新配置")
    group.add_argument("--validate", action="store_true", help="验证配置文件")
    group.add_argument("--save", action="store_true", help="保存配置到预设库")

    parser.add_argument("--output", "-o", default="config.yaml", help="输出文件路径（默认: config.yaml）")
    parser.add_argument("--config", "-c", default="config.yaml", help="要验证/保存的配置文件路径")
    parser.add_argument("--name", "-n", help="保存预设时的名称（格式: 学校_学院_论文类型）")

    args = parser.parse_args()

    if args.list_presets:
        list_presets()
    elif args.use:
        use_preset(args.use, args.output)
    elif args.create:
        create_from_template(args.output)
    elif args.validate:
        errors = validate(args.config)
        if not errors:
            print(f"✅ 配置文件验证通过: {args.config}")
        else:
            print(f"❌ 发现 {len(errors)} 个问题:")
            for err in errors:
                print(f"  {err}")
            sys.exit(1)
    elif args.save:
        if not args.name:
            print("错误: --save 需要 --name 参数（格式: 学校_学院_论文类型）")
            sys.exit(1)
        save_preset(args.config, args.name)


if __name__ == "__main__":
    main()
