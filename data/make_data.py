"""실습용 가짜 데이터를 만든다.

여기서 나오는 회사(예제유통), 지점, 품목, 숫자는 모두 지어낸 것이다.
실제 회사나 사람의 정보는 하나도 들어 있지 않다.

    python data/make_data.py

난수 씨앗을 고정해 두었기 때문에 몇 번을 돌려도 같은 파일이 나온다.
"""
import datetime as dt
import random
from pathlib import Path

from openpyxl import Workbook

HERE = Path(__file__).resolve().parent

# 품목코드, 품목명, 품목군, 단가
ITEMS = [
    ("P-101", "A4 복사용지 1박스", "사무용품", 24500),
    ("P-102", "유성펜 12자루", "사무용품", 6800),
    ("P-103", "파일철 10개", "사무용품", 9200),
    ("P-201", "물티슈 10팩", "생활용품", 11900),
    ("P-202", "종이컵 1000개", "생활용품", 15400),
    ("P-203", "주방세제 4개", "생활용품", 8700),
    ("P-301", "믹스커피 100개", "식음료", 13800),
    ("P-302", "생수 2L 12병", "식음료", 7900),
    ("P-303", "컵라면 24개", "식음료", 21600),
    ("P-401", "건전지 AA 20개", "전자소모품", 9900),
    ("P-402", "멀티탭 4구", "전자소모품", 12500),
    ("P-403", "USB 충전 케이블", "전자소모품", 5900),
]

# 지점: (전표 머리글자, 한 주 전표 수, 주간 목표 금액)
BRANCHES = {
    "강남점": ("GN", 62, 7_500_000),
    "수원점": ("SW", 50, 6_500_000),
    "대전점": ("DJ", 44, 5_500_000),
    "부산점": ("BS", 54, 7_500_000),
    "광주점": ("GJ", 40, 4_500_000),
}

# 주마다 지점별로 장사가 얼마나 됐는지(전표 수에 곱한다)
WEEKS = {
    "2026-10-26": {"강남점": 0.95, "수원점": 1.00, "대전점": 1.05, "부산점": 1.10, "광주점": 0.95},
    "2026-11-02": {"강남점": 1.10, "수원점": 1.00, "대전점": 0.90, "부산점": 0.95, "광주점": 1.10},
}

HEADER = ["전표번호", "판매일", "품목코드", "품목명", "수량", "단가", "금액"]
HEADER_DAEJEON = ["전표번호", "일자", "코드", "상품명", "판매수량", "판매단가", "판매금액"]


def make_rows(rng, branch, monday, factor):
    """한 지점의 한 주 판매 내역을 만든다. 월요일부터 토요일까지."""
    prefix, base_count, _ = BRANCHES[branch]
    count = round(base_count * factor)
    days = sorted(rng.choices(range(6), weights=[18, 16, 16, 17, 20, 13], k=count))
    rows = []
    for seq, offset in enumerate(days, start=1):
        day = monday + dt.timedelta(days=offset)
        code, name, _, price = rng.choice(ITEMS)
        qty = rng.choice([1, 2, 2, 3, 3, 4, 5, 5, 6, 8, 10, 10, 12, 15, 20, 24, 30])
        slip = f"{prefix}-{day:%m%d}-{seq:03d}"
        rows.append([slip, day, code, name, qty, price, qty * price])
    return rows


def write_branch_file(path, branch, monday, rows, rng, with_typo):
    wb = Workbook()
    ws = wb.active
    ws.title = "판매"
    sunday = monday + dt.timedelta(days=6)

    if branch == "수원점":
        # 버릇 1: 맨 위에 제목 줄을 넣고, 맨 아래에 합계 줄을 붙인다
        ws.append([f"수원점 주간 판매 현황 ({monday:%m/%d}~{sunday:%m/%d})"])
        ws.append([])
        ws.append(HEADER)
        for r in rows:
            ws.append(r)
        ws.append(["합계", None, None, None, sum(r[4] for r in rows), None, sum(r[6] for r in rows)])

    elif branch == "대전점":
        # 버릇 2: 열 이름을 다르게 쓴다
        ws.append(HEADER_DAEJEON)
        for r in rows:
            ws.append(r)

    elif branch == "부산점":
        # 버릇 3: 날짜와 금액을 글자로 적는다. 품목명 뒤에 빈칸이 붙기도 한다
        ws.append(HEADER)
        typo_at = len(rows) // 2
        for i, r in enumerate(rows):
            slip, day, code, name, qty, price, amount = r
            if with_typo and i == typo_at:
                amount += 90_000  # 손으로 고쳐 적다 틀린 금액 한 건
            if i % 7 == 3:
                name = name + "  "
            ws.append([slip, f"{day:%Y.%m.%d}", code, name, qty, price, f"{amount:,}원"])

    elif branch == "광주점":
        # 버릇 4: 같은 전표를 두 번 붙여 넣고, 중간에 빈 줄을 둔다
        ws.append(HEADER)
        dup_at = set(rng.sample(range(len(rows)), 3))
        blank_at = {len(rows) // 3, 2 * len(rows) // 3}
        for i, r in enumerate(rows):
            if i in blank_at:
                ws.append([])
            ws.append(r)
            if i in dup_at:
                ws.append(r)

    else:
        ws.append(HEADER)
        for r in rows:
            ws.append(r)

    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, dt.date):
                cell.number_format = "yyyy-mm-dd"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def write_table(path, title, header, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = title
    ws.append(header)
    for r in rows:
        ws.append(r)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def main():
    rng = random.Random(20261102)

    for week, factors in WEEKS.items():
        monday = dt.date.fromisoformat(week)
        for branch in BRANCHES:
            rows = make_rows(rng, branch, monday, factors[branch])
            path = HERE / "raw" / week / f"{branch}_주간판매.xlsx"
            write_branch_file(path, branch, monday, rows, rng, with_typo=(week == "2026-11-02"))

    write_table(
        HERE / "master" / "품목.xlsx", "품목",
        ["품목코드", "품목명", "품목군", "기준단가"],
        [list(item) for item in ITEMS],
    )
    write_table(
        HERE / "master" / "지점목표.xlsx", "목표",
        ["지점", "주간목표"],
        [[name, target] for name, (_, _, target) in BRANCHES.items()],
    )

    # 7장 실습용: 열 이름이 영어인 새 지점 파일
    monday = dt.date(2026, 11, 2)
    jeju = make_rows(rng, "광주점", monday, 0.5)
    write_table(
        HERE / "extra" / "제주점_주간판매.xlsx", "sales",
        ["slip_no", "date", "sku", "item", "qty", "unit_price", "amount"],
        [[r[0].replace("GJ", "JJ")] + r[1:] for r in jeju],
    )
    print("데이터를 만들었습니다:", ", ".join(WEEKS), "+ master, extra")


if __name__ == "__main__":
    main()
