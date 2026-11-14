"""주간 판매 보고서 만들기.

2~5장에서 하나씩 만든 함수를 한 파일에 모았다(6장).
지점 파일을 읽어 합치고, 정리하고, 집계해서 서식 있는 엑셀 보고서로 저장한다.

    import weekly_report as wr
    path, notes = wr.make_report("2026-11-02")
"""
import unicodedata
from pathlib import Path

import pandas as pd
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
OUTPUT = BASE / "output"

# ── 내 회사 파일에 맞게 고치는 곳 ─────────────────────────────
# 보고서가 쓰는 표준 열 이름
STANDARD = ["전표번호", "판매일", "품목코드", "품목명", "수량", "단가", "금액"]

# 지점마다 다르게 쓰는 열 이름 → 표준 이름
COLUMN_NAMES = {
    "일자": "판매일",
    "코드": "품목코드",
    "상품명": "품목명",
    "판매수량": "수량",
    "판매단가": "단가",
    "판매금액": "금액",
}

# 전표번호 모양: 영문 두 글자-날짜 네 자리-일련번호 세 자리 (예: GN-1102-001)
SLIP_PATTERN = r"^[A-Z]{2}-\d{4}-\d{3}$"
# ──────────────────────────────────────────────────────────


# 2장: 읽어서 합치기 -------------------------------------------------

def find_header_row(path, max_rows=10):
    """앞쪽 몇 줄을 읽어, 열 이름이 적힌 줄이 몇 번째인지(0부터) 찾는다."""
    known = set(STANDARD) | set(COLUMN_NAMES)
    top = pd.read_excel(path, header=None, nrows=max_rows)
    for i, row in top.iterrows():
        cells = {str(v).strip() for v in row.tolist()}
        if len(cells & known) >= 4:
            return i
    raise ValueError(
        f"{path.name}: 열 이름이 적힌 줄을 찾지 못했습니다. "
        f"위쪽 {max_rows}줄 안에 {STANDARD} 같은 이름이 있어야 합니다."
    )


def read_branch_file(path):
    """지점 파일 하나를 읽어 표준 열 이름으로 맞춘다."""
    path = Path(path)
    df = pd.read_excel(path, header=find_header_row(path))
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns=COLUMN_NAMES)

    missing = [c for c in STANDARD if c not in df.columns]
    if missing:
        raise ValueError(f"{path.name}: 필요한 열이 없습니다 → {missing}")

    df = df[STANDARD].copy()
    name = unicodedata.normalize("NFC", path.stem)  # 맥에서 만든 파일 이름은 자모가 풀려 있을 수 있다
    df.insert(0, "지점", name.split("_")[0])
    df.insert(1, "원본파일", name + path.suffix)
    return df


def read_week(week, data_dir=DATA):
    """raw/<week> 폴더의 엑셀 파일을 모두 읽어 한 표로 합친다."""
    week_dir = Path(data_dir) / "raw" / week
    if not week_dir.is_dir():
        raise FileNotFoundError(f"raw/{week} 폴더가 없습니다. 지점 파일을 이 폴더에 넣었는지 확인하세요.")
    # '~$'로 시작하는 파일은 엑셀이 열려 있을 때 생기는 임시 파일이다
    files = sorted(p for p in week_dir.glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        raise FileNotFoundError(f"raw/{week} 폴더에 엑셀 파일(.xlsx)이 없습니다.")
    return pd.concat([read_branch_file(f) for f in files], ignore_index=True)


# 3장: 정리하기 ------------------------------------------------------

def to_number(series):
    """'1,250,000원' 같은 글자를 숫자로 바꾼다. 못 바꾼 칸은 빈 값이 된다."""
    text = series.astype("string").str.replace(r"[,원\s]", "", regex=True)
    return pd.to_numeric(text, errors="coerce")


def to_date(series):
    """'2026.11.03' 같은 글자를 날짜로 바꾼다. 못 바꾼 칸은 빈 값이 된다."""
    text = series.astype("string").str.strip().str.replace(".", "-", regex=False)
    return pd.to_datetime(text, format="ISO8601", errors="coerce")


def clean(merged, week):
    """합친 표를 정리한다. (정리한 표, 확인이 필요한 줄, 메모 목록)을 돌려준다."""
    notes = []
    df = merged.copy()

    # 1. 전표번호가 전표 모양인 줄만 판매 내역이다. 빈 줄과 합계 줄이 여기서 빠진다
    slip = df["전표번호"].astype("string").str.strip()
    is_sale = slip.str.match(SLIP_PATTERN).fillna(False).astype(bool)
    total_rows = df[slip == "합계"]
    df = df[is_sale].copy()
    df["전표번호"] = slip[is_sale]

    # 2. 글자 앞뒤 빈칸
    for col in ["품목코드", "품목명"]:
        df[col] = df[col].astype("string").str.strip()

    # 3. 날짜와 숫자
    df["판매일"] = to_date(df["판매일"])
    for col in ["수량", "단가", "금액"]:
        df[col] = to_number(df[col])
    broken = df[df[["판매일", "수량", "단가", "금액"]].isna().any(axis=1)]
    if len(broken):
        where = ", ".join(broken["전표번호"].head(5))
        raise ValueError(f"날짜나 숫자로 읽지 못한 줄이 {len(broken)}건 있습니다. 전표번호: {where}")
    for col in ["수량", "단가", "금액"]:
        df[col] = df[col].astype("int64")

    # 4. 같은 지점에 같은 전표번호가 두 번 있으면 한 번만 남긴다
    dup = df.duplicated(subset=["지점", "전표번호"], keep="first")
    if dup.any():
        counts = df[dup].groupby("지점").size()
        notes.append("중복 전표 제거: " + ", ".join(f"{b} {n}건" for b, n in counts.items()))
    df = df[~dup].reset_index(drop=True)

    # 5. 지점이 적어 보낸 합계와 맞는지 검산한다
    for _, row in total_rows.iterrows():
        branch = row["지점"]
        theirs = int(to_number(pd.Series([row["금액"]])).iloc[0])
        ours = int(df.loc[df["지점"] == branch, "금액"].sum())
        mark = "일치" if theirs == ours else "불일치, 확인 필요"
        notes.append(f"{branch} 합계 검산: 파일 {theirs:,}원 / 계산 {ours:,}원 → {mark}")

    # 6. 고치지 않고 표시만 하는 줄: 금액이 수량×단가와 다르거나, 날짜가 그 주가 아닌 줄
    monday = pd.Timestamp(week)
    wrong_amount = df["금액"] != df["수량"] * df["단가"]
    out_of_week = ~df["판매일"].between(monday, monday + pd.Timedelta(days=6))
    check = df[wrong_amount | out_of_week].copy()
    check["사유"] = ""
    check.loc[wrong_amount, "사유"] = "금액이 수량×단가와 다름"
    check.loc[out_of_week, "사유"] = "판매일이 이번 주가 아님"
    if len(check):
        notes.append(f"확인이 필요한 줄 {len(check)}건 ('확인필요' 시트)")

    return df, check, notes


# 4장: 집계하기 ------------------------------------------------------

def summarize(sales, prev_sales, items, targets):
    """(지점별 요약, 지점×품목군 표, 많이 팔린 품목 5개)를 돌려준다."""
    by_branch = sales.groupby("지점", as_index=False)["금액"].sum()
    prev = (
        prev_sales.groupby("지점", as_index=False)["금액"].sum()
        .rename(columns={"금액": "전주금액"})
    )
    # 목표 표를 왼쪽에 두면 지점 순서가 목표 파일 순서로 고정된다
    summary = (
        targets.merge(by_branch, on="지점", how="left", validate="one_to_one")
        .merge(prev, on="지점", how="left", validate="one_to_one")
    )
    summary["금액"] = summary["금액"].fillna(0).astype("int64")
    summary["달성률"] = summary["금액"] / summary["주간목표"]
    summary["전주대비"] = summary["금액"] / summary["전주금액"] - 1
    summary = summary[["지점", "금액", "주간목표", "달성률", "전주금액", "전주대비"]]

    with_group = sales.merge(
        items[["품목코드", "품목군"]], on="품목코드", how="left", validate="many_to_one"
    )
    unknown = with_group.loc[with_group["품목군"].isna(), "품목코드"].unique()
    if len(unknown):
        raise ValueError(f"품목 표에 없는 품목코드가 있습니다 → {list(unknown)}")
    pivot = with_group.pivot_table(
        index="지점", columns="품목군", values="금액",
        aggfunc="sum", fill_value=0, margins=True, margins_name="합계",
    )
    pivot = pivot.reindex(list(targets["지점"]) + ["합계"]).fillna(0).astype("int64")

    top_items = (
        sales.groupby(["품목코드", "품목명"], as_index=False)[["수량", "금액"]].sum()
        .sort_values("금액", ascending=False)
        .head(5)
    )
    return summary, pivot, top_items


# 5장: 엑셀로 내보내기 ------------------------------------------------

HEADER_FILL = PatternFill("solid", fgColor="E7E6E6")
MONEY = "#,##0"
PERCENT = "0.0%"
DATE = "yyyy-mm-dd"


def style_table(ws, header_row, formats=None, widths=None):
    """머리글을 굵게 칠하고, 열마다 숫자 서식과 너비를 준다."""
    for cell in ws[header_row]:
        if cell.value is not None:
            cell.font = Font(bold=True)
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(horizontal="center")
    for col, number_format in (formats or {}).items():
        for cell in ws[col][header_row:]:
            cell.number_format = number_format
    for col, width in (widths or {}).items():
        ws.column_dimensions[col].width = width


def write_report(path, week, summary, pivot, top_items, sales, check):
    """표 다섯 개를 시트 네 장짜리 엑셀 파일로 저장한다."""
    monday = pd.Timestamp(week)
    sunday = monday + pd.Timedelta(days=6)
    n = len(summary)

    total = pd.DataFrame([{
        "지점": "전체",
        "금액": summary["금액"].sum(),
        "주간목표": summary["주간목표"].sum(),
        "달성률": summary["금액"].sum() / summary["주간목표"].sum(),
        "전주금액": summary["전주금액"].sum(),
        "전주대비": summary["금액"].sum() / summary["전주금액"].sum() - 1,
    }])
    summary_out = pd.concat([summary, total], ignore_index=True)
    top_row = n + 7  # 요약 표 아래 두 줄 띄우고 '많이 팔린 품목' 제목

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        summary_out.to_excel(writer, sheet_name="요약", index=False, startrow=2)
        top_items.to_excel(writer, sheet_name="요약", index=False, startrow=top_row)
        pivot.to_excel(writer, sheet_name="품목군별")
        sales.to_excel(writer, sheet_name="상세", index=False)
        check.to_excel(writer, sheet_name="확인필요", index=False)

        # 요약 시트
        ws = writer.sheets["요약"]
        ws["A1"] = f"주간 판매 보고서 ({monday:%m/%d}~{sunday:%m/%d})"
        ws["A1"].font = Font(bold=True, size=14)
        ws.cell(row=top_row, column=1, value="많이 팔린 품목 5개").font = Font(bold=True)
        style_table(
            ws, header_row=3,
            formats={"B": MONEY, "C": MONEY, "D": PERCENT, "E": MONEY, "F": PERCENT},
            widths={"A": 14, "B": 18, "C": 14, "D": 12, "E": 14, "F": 12},
        )
        style_table(ws, header_row=top_row + 1)
        for row in ws.iter_rows(min_row=top_row + 2, max_row=top_row + 1 + len(top_items), min_col=3, max_col=4):
            for cell in row:
                cell.number_format = MONEY
        for cell in ws[3 + n + 1]:  # '전체' 줄
            cell.font = Font(bold=True)
        # 목표에 못 미친 달성률은 빨간 글자
        ws.conditional_formatting.add(
            f"D4:D{3 + n}",
            CellIsRule(operator="lessThan", formula=["1"], font=Font(color="C00000", bold=True)),
        )
        # 지점별 금액과 목표 막대 그래프
        chart = BarChart()
        chart.title = "지점별 판매 금액과 목표"
        chart.title.overlay = False  # 제목과 범례가 막대 위에 겹치지 않게
        chart.y_axis.numFmt = MONEY
        chart.x_axis.delete = False  # 이 두 줄이 없으면 축 글자가 보이지 않는다
        chart.y_axis.delete = False
        chart.legend.position = "b"
        chart.legend.overlay = False
        chart.add_data(Reference(ws, min_col=2, max_col=3, min_row=3, max_row=3 + n), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=4, max_row=3 + n))
        chart.width, chart.height = 16, 8.5
        ws.add_chart(chart, "H3")
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True

        # 품목군별 시트
        ws = writer.sheets["품목군별"]
        cols = [get_column_letter(i) for i in range(2, ws.max_column + 1)]
        style_table(ws, 1, formats={c: MONEY for c in cols}, widths={"A": 12, **{c: 14 for c in cols}})
        for cell in ws[ws.max_row]:
            cell.font = Font(bold=True)

        # 상세, 확인필요 시트
        widths = {"A": 9, "B": 22, "C": 15, "D": 12, "E": 10, "F": 20, "G": 8, "H": 10, "I": 12, "J": 26}
        for name in ["상세", "확인필요"]:
            ws = writer.sheets[name]
            style_table(ws, 1, formats={"D": DATE, "G": MONEY, "H": MONEY, "I": MONEY}, widths=widths)
            ws.freeze_panes = "A2"
        writer.sheets["상세"].auto_filter.ref = writer.sheets["상세"].dimensions


# 6장: 한 번에 실행하기 -----------------------------------------------

def make_report(week, data_dir=DATA, output_dir=OUTPUT):
    """week(그 주 월요일, 'YYYY-MM-DD')의 보고서를 만들고 (파일 경로, 메모 목록)을 돌려준다."""
    data_dir, output_dir = Path(data_dir), Path(output_dir)
    monday = pd.Timestamp(week)
    if monday.weekday() != 0:
        raise ValueError(f"{week}은 월요일이 아닙니다. 보고할 주의 월요일 날짜를 넣으세요.")
    prev_week = f"{monday - pd.Timedelta(days=7):%Y-%m-%d}"

    sales, check, notes = clean(read_week(week, data_dir), week)
    prev_sales, _, _ = clean(read_week(prev_week, data_dir), prev_week)
    items = pd.read_excel(data_dir / "master" / "품목.xlsx")
    targets = pd.read_excel(data_dir / "master" / "지점목표.xlsx")

    absent = [b for b in targets["지점"] if b not in set(sales["지점"])]
    if absent:
        notes.append("파일이 오지 않은 지점: " + ", ".join(absent))

    summary, pivot, top_items = summarize(sales, prev_sales, items, targets)

    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / f"주간판매보고서_{week}.xlsx"
    try:
        write_report(path, week, summary, pivot, top_items, sales, check)
    except PermissionError:
        raise ValueError(f"{path.name} 파일이 엑셀에서 열려 있습니다. 닫고 다시 실행하세요.") from None

    # 마지막 검산: 저장한 파일을 다시 읽어 합계가 같은지 본다
    saved = pd.read_excel(path, sheet_name="상세")
    if int(saved["금액"].sum()) != int(sales["금액"].sum()):
        raise ValueError("저장한 파일의 합계가 계산한 합계와 다릅니다. 보고서를 보내지 마세요.")
    notes.append(f"판매 {len(sales)}건, 합계 {int(sales['금액'].sum()):,}원")
    return path, notes
