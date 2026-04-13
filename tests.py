"""單元測試 (Unit Tests)"""

import sys
import os
import json
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(__file__))

import todo
import calculator
import converter
import timer


class TestCalculator(unittest.TestCase):
    def test_addition(self):
        self.assertAlmostEqual(calculator.calculate("1 + 2"), 3.0)

    def test_subtraction(self):
        self.assertAlmostEqual(calculator.calculate("10 - 4"), 6.0)

    def test_multiplication(self):
        self.assertAlmostEqual(calculator.calculate("3 * 4"), 12.0)

    def test_division(self):
        self.assertAlmostEqual(calculator.calculate("10 / 4"), 2.5)

    def test_floor_division(self):
        self.assertAlmostEqual(calculator.calculate("10 // 3"), 3.0)

    def test_modulo(self):
        self.assertAlmostEqual(calculator.calculate("10 % 3"), 1.0)

    def test_power(self):
        self.assertAlmostEqual(calculator.calculate("2 ** 8"), 256.0)

    def test_complex_expression(self):
        self.assertAlmostEqual(calculator.calculate("3 + 4 * 2"), 11.0)

    def test_division_by_zero(self):
        with self.assertRaises(ZeroDivisionError):
            calculator.calculate("1 / 0")

    def test_invalid_expression(self):
        with self.assertRaises(ValueError):
            calculator.calculate("import os")


class TestConverter(unittest.TestCase):
    def test_km_to_m(self):
        self.assertAlmostEqual(converter.convert(1, "公里", "公尺"), 1000.0)

    def test_m_to_cm(self):
        self.assertAlmostEqual(converter.convert(1, "公尺", "公分"), 100.0)

    def test_kg_to_g(self):
        self.assertAlmostEqual(converter.convert(1, "公斤", "公克"), 1000.0)

    def test_lb_to_kg(self):
        self.assertAlmostEqual(converter.convert(1, "磅", "公斤"), 0.453592, places=5)

    def test_celsius_to_fahrenheit(self):
        self.assertAlmostEqual(converter.convert(100, "°C", "°F"), 212.0)

    def test_fahrenheit_to_celsius(self):
        self.assertAlmostEqual(converter.convert(32, "°F", "°C"), 0.0)

    def test_celsius_to_kelvin(self):
        self.assertAlmostEqual(converter.convert(0, "°C", "K"), 273.15)

    def test_mile_to_km(self):
        self.assertAlmostEqual(converter.convert(1, "英里", "公里"), 1.609344, places=5)

    def test_unsupported_unit(self):
        with self.assertRaises(ValueError):
            converter.convert(1, "光年", "公尺")


class TestTodo(unittest.TestCase):
    def setUp(self):
        # 用臨時檔案隔離每個測試
        self._tmp = tempfile.NamedTemporaryFile(suffix=".json", delete=False)
        self._tmp.close()
        todo.TODO_FILE = self._tmp.name
        # 清空檔案
        with open(todo.TODO_FILE, "w") as f:
            json.dump([], f)

    def tearDown(self):
        os.unlink(self._tmp.name)

    def test_add_task(self):
        task = todo.add_task("買牛奶")
        self.assertEqual(task["title"], "買牛奶")
        self.assertFalse(task["done"])
        self.assertEqual(task["id"], 1)

    def test_list_tasks(self):
        todo.add_task("任務一")
        todo.add_task("任務二")
        tasks = todo.list_tasks()
        self.assertEqual(len(tasks), 2)

    def test_complete_task(self):
        task = todo.add_task("測試任務")
        result = todo.complete_task(task["id"])
        self.assertTrue(result)
        tasks = todo.list_tasks()
        self.assertTrue(tasks[0]["done"])

    def test_complete_nonexistent_task(self):
        result = todo.complete_task(999)
        self.assertFalse(result)

    def test_delete_task(self):
        task = todo.add_task("刪除我")
        result = todo.delete_task(task["id"])
        self.assertTrue(result)
        tasks = todo.list_tasks()
        self.assertEqual(len(tasks), 0)

    def test_delete_nonexistent_task(self):
        result = todo.delete_task(999)
        self.assertFalse(result)

    def test_list_tasks_hide_done(self):
        t1 = todo.add_task("未完成")
        t2 = todo.add_task("已完成")
        todo.complete_task(t2["id"])
        tasks = todo.list_tasks(show_done=False)
        self.assertEqual(len(tasks), 1)
        self.assertEqual(tasks[0]["id"], t1["id"])

    def test_id_increments(self):
        t1 = todo.add_task("第一")
        t2 = todo.add_task("第二")
        self.assertEqual(t2["id"], t1["id"] + 1)


class TestTimer(unittest.TestCase):
    def test_countdown_calls_tick(self):
        ticks = []
        timer.countdown(3, tick_callback=ticks.append)
        self.assertEqual(ticks, [3, 2, 1, 0])

    def test_countdown_zero(self):
        ticks = []
        timer.countdown(0, tick_callback=ticks.append)
        self.assertEqual(ticks, [0])


if __name__ == "__main__":
    unittest.main()
