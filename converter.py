"""單位換算器 (Unit Converter)"""

# 所有換算都先轉換成 SI 基本單位再換算

_LENGTH = {
    "公里": 1000,
    "km": 1000,
    "公尺": 1,
    "m": 1,
    "公分": 0.01,
    "cm": 0.01,
    "公釐": 0.001,
    "mm": 0.001,
    "英里": 1609.344,
    "mile": 1609.344,
    "英尺": 0.3048,
    "ft": 0.3048,
    "英寸": 0.0254,
    "in": 0.0254,
    "碼": 0.9144,
    "yd": 0.9144,
}

_WEIGHT = {
    "公噸": 1000,
    "t": 1000,
    "公斤": 1,
    "kg": 1,
    "公克": 0.001,
    "g": 0.001,
    "毫克": 0.000001,
    "mg": 0.000001,
    "磅": 0.453592,
    "lb": 0.453592,
    "盎司": 0.0283495,
    "oz": 0.0283495,
}

_TEMPERATURE_UNITS = {"°C", "°F", "K"}

CATEGORIES = {
    "長度": _LENGTH,
    "重量": _WEIGHT,
}


def convert(value: float, from_unit: str, to_unit: str) -> float:
    """
    換算數值。支援長度、重量與溫度。

    範例:
        convert(1, "公里", "公尺")  -> 1000.0
        convert(100, "°C", "°F")   -> 212.0
    """
    # 溫度特殊處理
    if from_unit in _TEMPERATURE_UNITS or to_unit in _TEMPERATURE_UNITS:
        return _convert_temperature(value, from_unit, to_unit)

    for table in CATEGORIES.values():
        if from_unit in table and to_unit in table:
            return value * table[from_unit] / table[to_unit]

    raise ValueError(f"不支援的單位: {from_unit!r} 或 {to_unit!r}")


def _convert_temperature(value: float, from_unit: str, to_unit: str) -> float:
    valid = {"°C", "°F", "K"}
    if from_unit not in valid or to_unit not in valid:
        raise ValueError(f"不支援的溫度單位: {from_unit!r} 或 {to_unit!r}")

    # 先轉換成攝氏
    if from_unit == "°C":
        celsius = value
    elif from_unit == "°F":
        celsius = (value - 32) * 5 / 9
    else:  # K
        celsius = value - 273.15

    # 再從攝氏轉換成目標單位
    if to_unit == "°C":
        return celsius
    elif to_unit == "°F":
        return celsius * 9 / 5 + 32
    else:  # K
        return celsius + 273.15


def run_converter_cli() -> None:
    """互動式單位換算器 CLI。"""
    print("\n=== 單位換算器 ===")

    categories = {
        "1": ("長度", _LENGTH),
        "2": ("重量", _WEIGHT),
        "3": ("溫度", None),
        "0": ("返回", None),
    }

    while True:
        print("\n類別:")
        for k, (name, _) in categories.items():
            print(f"  {k}. {name}")
        choice = input("請選擇: ").strip()

        if choice == "0":
            break
        elif choice not in categories:
            print("無效選項")
            continue

        cat_name, table = categories[choice]

        if choice == "3":
            units = list(_TEMPERATURE_UNITS)
        else:
            units = list(table.keys())

        print(f"\n可用單位 ({cat_name}): {', '.join(units)}")
        try:
            value = float(input("輸入數值: "))
            from_unit = input("從 (單位): ").strip()
            to_unit = input("到 (單位): ").strip()
            result = convert(value, from_unit, to_unit)
            print(f"  {value} {from_unit} = {result:.6g} {to_unit}")
        except ValueError as e:
            print(f"  錯誤: {e}")
