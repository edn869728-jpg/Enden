"""待辦事項管理器 (To-Do List Manager)"""

import json
import os
from datetime import datetime

TODO_FILE = os.path.join(os.path.dirname(__file__), "todos.json")


def _load() -> list:
    if os.path.exists(TODO_FILE):
        with open(TODO_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def _save(todos: list) -> None:
    with open(TODO_FILE, "w", encoding="utf-8") as f:
        json.dump(todos, f, ensure_ascii=False, indent=2)


def add_task(title: str) -> dict:
    """新增一個待辦事項。"""
    title = title.strip()
    if not title:
        raise ValueError("任務名稱不能為空")
    todos = _load()
    task = {
        "id": (max((t["id"] for t in todos), default=0) + 1),
        "title": title,
        "done": False,
        "created_at": datetime.now().isoformat(),
    }
    todos.append(task)
    _save(todos)
    return task


def list_tasks(show_done: bool = True) -> list:
    """列出所有待辦事項。"""
    todos = _load()
    if not show_done:
        todos = [t for t in todos if not t["done"]]
    return todos


def complete_task(task_id: int) -> bool:
    """將指定 ID 的待辦事項標記為完成。"""
    todos = _load()
    for task in todos:
        if task["id"] == task_id:
            task["done"] = True
            _save(todos)
            return True
    return False


def delete_task(task_id: int) -> bool:
    """刪除指定 ID 的待辦事項。"""
    todos = _load()
    new_todos = [t for t in todos if t["id"] != task_id]
    if len(new_todos) == len(todos):
        return False
    _save(new_todos)
    return True


def _print_tasks(tasks: list) -> None:
    if not tasks:
        print("  (無任務)")
        return
    for t in tasks:
        status = "✔" if t["done"] else "○"
        print(f"  [{status}] #{t['id']}  {t['title']}  ({t['created_at'][:10]})")


def run_todo_cli() -> None:
    """互動式待辦事項 CLI。"""
    while True:
        print("\n=== 待辦事項管理器 ===")
        print("  1. 新增任務")
        print("  2. 列出所有任務")
        print("  3. 列出未完成任務")
        print("  4. 完成任務")
        print("  5. 刪除任務")
        print("  0. 返回主選單")
        choice = input("請選擇: ").strip()

        if choice == "1":
            title = input("任務名稱: ").strip()
            if not title:
                print("  ✘ 任務名稱不能為空")
                continue
            task = add_task(title)
            print(f"  ✔ 已新增任務 #{task['id']}: {task['title']}")

        elif choice == "2":
            _print_tasks(list_tasks(show_done=True))

        elif choice == "3":
            _print_tasks(list_tasks(show_done=False))

        elif choice == "4":
            try:
                task_id = int(input("任務 ID: ").strip())
            except ValueError:
                print("  ✘ 請輸入有效的數字 ID")
                continue
            if complete_task(task_id):
                print(f"  ✔ 任務 #{task_id} 已標記完成")
            else:
                print(f"  ✘ 找不到任務 #{task_id}")

        elif choice == "5":
            try:
                task_id = int(input("任務 ID: ").strip())
            except ValueError:
                print("  ✘ 請輸入有效的數字 ID")
                continue
            if delete_task(task_id):
                print(f"  ✔ 任務 #{task_id} 已刪除")
            else:
                print(f"  ✘ 找不到任務 #{task_id}")

        elif choice == "0":
            break

        else:
            print("  ✘ 無效選項，請輸入 0-5")
