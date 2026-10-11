"""生成两份小型示例工作簿。"""

from pathlib import Path
from datetime import date
from openpyxl import Workbook


EXAMPLES = Path(__file__).parent / "examples"


def make_file(path: Path, rows: list[tuple]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "订单"
    sheet.append(("订单号", "客户", "金额", "日期"))
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    workbook.close()


if __name__ == "__main__":
    EXAMPLES.mkdir(exist_ok=True)
    make_file(EXAMPLES / "订单_一.xlsx", [
        ("A001", "张三", 120, date(2026, 1, 3)),
        ("A002", "李四", 85, date(2026, 1, 1)),
    ])
    make_file(EXAMPLES / "订单_二.xlsx", [
        ("A001", "张三", 120, date(2026, 1, 3)),
        ("A003", "王五", 200, date(2026, 1, 2)),
    ])
    print(f"已生成示例文件：{EXAMPLES}")
