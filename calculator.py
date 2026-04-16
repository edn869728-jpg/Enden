"""計算機 (Calculator)"""

import re


def calculate(expression: str) -> float:
    """
    計算一個數學運算式並回傳結果。
    支援: +, -, *, /, //, %, **  以及括號與小數點。

    範例:
        calculate("3 + 4 * 2")   -> 11.0
        calculate("10 / 2")      -> 5.0
        calculate("(2 + 3) ** 2") -> 25.0
    """
    if not expression or not re.fullmatch(r"[\d\s\+\-\*\/\%\.\(\)]+", expression):
        raise ValueError(f"不合法的運算式: {expression!r}")
    try:
        result = eval(expression, {"__builtins__": {}})  # noqa: S307
    except ZeroDivisionError:
        raise ZeroDivisionError("除以零錯誤")
    except Exception as exc:
        raise ValueError(f"無法計算運算式: {exc}") from exc
    return float(result)


def run_calculator_cli() -> None:
    """互動式計算機 CLI。"""
    print("\n=== 計算機 ===")
    print("支援運算符: + - * / // % **  以及括號 ( )")
    print("直接按 Enter（空白）返回主選單")
    while True:
        expr = input("\n運算式: ").strip()
        if not expr:
            break
        try:
            result = calculate(expr)
            if result == int(result):
                print(f"  = {int(result)}")
            else:
                print(f"  = {result}")
        except (ValueError, ZeroDivisionError) as e:
            print(f"  錯誤: {e}")
