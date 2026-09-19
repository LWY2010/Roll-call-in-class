"""
密码重置工具
=============
双击运行，输入新密码即可。
主程序下次启动时自动读取新密码。

也可以打包成 exe：
    pyinstaller -F -n "密码重置" reset_password.py
（注意不加 -w，因为要显示命令行）
"""
import json, os, sys
from pathlib import Path

APP_NAME = "ClassRollCall"
DATA_DIR = Path(os.environ.get('APPDATA', Path.home())) / APP_NAME
DATA_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = DATA_DIR / "config.json"
DEFAULT_PASSWORD = "01180204"


def load_json(path, default):
    if path.exists():
        try: return json.loads(path.read_text(encoding='utf-8'))
        except Exception: pass
    return default

def save_json(path, data):
    try: path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    except Exception as e:
        print("保存失败：", e)


def main():
    print()
    print("=" * 52)
    print("      课堂随机点名 · 密码重置工具")
    print("=" * 52)
    print()

    cfg = load_json(CONFIG_FILE, {})
    cur = cfg.get('password', '')
    if cur:
        print(f"  当前密码：{cur}")
    else:
        print(f"  当前密码：未设置（使用默认密码 {DEFAULT_PASSWORD}）")
    print()

    while True:
        new_pwd = input("  请输入新密码（直接回车取消）：").strip()
        if not new_pwd:
            print()
            print("  已取消，未修改。")
            input("\n  按回车键退出...")
            return

        if len(new_pwd) < 4:
            print("  ⚠ 密码至少 4 位，请重新输入。\n")
            continue

        confirm = input("  请再次输入确认：").strip()
        if new_pwd != confirm:
            print("  ⚠ 两次输入不一致，请重新输入。\n")
            continue

        break

    cfg['password'] = new_pwd
    save_json(CONFIG_FILE, cfg)

    print()
    print("  ✅ 密码已重置成功！")
    print(f"     新密码：{new_pwd}")
    print(f"     配置文件：{CONFIG_FILE}")
    print()
    print("  主程序下次启动时自动生效。")
    input("\n  按回车键退出...")


if __name__ == "__main__":
    main()
