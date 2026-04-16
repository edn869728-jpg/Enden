"""Enden — 一些方便的程式系統 (Convenient Program Systems)

使用方法:
    python main.py
"""

from calculator import run_calculator_cli
from converter import run_converter_cli
from timer import run_timer_cli
from todo import run_todo_cli

MENU = {
    "1": ("待辦事項管理器", run_todo_cli),
    "2": ("計算機", run_calculator_cli),
    "3": ("單位換算器", run_converter_cli),
    "4": ("計時器", run_timer_cli),
    "0": ("離開", None),
}


def main() -> None:
    print("╔══════════════════════════════╗")
    print("║   Enden — 方便程式系統        ║")
    print("╚══════════════════════════════╝")
    while True:
        print("\n主選單:")
        for key, (name, _) in MENU.items():
            print(f"  {key}. {name}")
        choice = input("請選擇: ").strip()
        if choice == "0":
            print("再見！")
            break
        if choice in MENU and MENU[choice][1] is not None:
            MENU[choice][1]()
        else:
            print("  ✘ 無效選項，請輸入 0-4")


if __name__ == "__main__":
    main()
