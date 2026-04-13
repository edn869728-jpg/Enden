"""計算機 (Calculator)"""

import operator
import re


OPERATORS = {
    "+": operator.add,
    "-": operator.sub,
    "*": operator.mul,
    "/": operator.truediv,
    "//": operator.floordiv,
    "%": operator.mod,
    "**": operator.pow,
}


def calculate(expression: str) -> float:
    """
    計算一個數學運算式並回傳結果。
    支援: +, -, *, /, //, %, **

    範例:
        calculate("3 + 4 * 2")  -> 11.0
        calculate("10 / 2")     -> 5.0
    """
    # 只允許數字、空格以及合法的運算符號，避免程式碼注入
    if not re.fullmatch(r"[\d\s\+\-\*\/\%\.\(\)]+", expression):
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
    print("支援運算符: + - * / // % **  (輸入 '0' 返回)")
    while True:
        expr = input("輸入運算式: ").strip()
        if expr == "0":
            break
        if not expr:
            continue
        try:
            result = calculate(expr)
            # 如果結果是整數就不顯示小數點
            if result == int(result):
                print(f"  = {int(result)}")
            else:
                print(f"  = {result}")
        except (ValueError, ZeroDivisionError) as e:
            print(f"  錯誤: {e}")
