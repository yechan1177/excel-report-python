"""주간 판매 보고서를 한 번에 만든다.

    python run_report.py              지난주(월~일) 보고서
    python run_report.py 2026-11-02   그 날짜(월요일)로 시작하는 주의 보고서
"""
import datetime as dt
import sys

import weekly_report as wr


def last_monday(today=None):
    """오늘을 기준으로 지난주 월요일 날짜를 구한다."""
    today = today or dt.date.today()
    this_monday = today - dt.timedelta(days=today.weekday())
    return this_monday - dt.timedelta(days=7)


def write_log(line):
    """예약 실행은 화면을 볼 수 없으니, 결과를 파일에도 한 줄씩 남긴다."""
    wr.OUTPUT.mkdir(parents=True, exist_ok=True)
    with open(wr.OUTPUT / "run_log.txt", "a", encoding="utf-8") as f:
        f.write(f"{dt.datetime.now():%Y-%m-%d %H:%M} {line}\n")


def main(argv):
    week = argv[1] if len(argv) > 1 else f"{last_monday():%Y-%m-%d}"
    try:
        path, notes = wr.make_report(week)
    except (FileNotFoundError, ValueError) as error:
        print("보고서를 만들지 못했습니다.")
        print("이유:", error)
        write_log(f"[실패] {week} {error}")
        return 1

    print(f"{week} 주 보고서를 만들었습니다: output/{path.name}")
    for note in notes:
        print(" -", note)
    write_log(f"[성공] {week} " + " / ".join(notes))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
