"""교육 PC에서 쓸 더미 운영 메일(.eml)을 만듭니다. 교육용 더미 데이터입니다.

업무 PC에서는 Outlook에서 실제 메일을 읽지만, 교육 PC에는 회사 메일이 없습니다.
그래서 실제 메일과 같은 모양의 가짜 메일을 파일로 만들어 두고, 그 폴더를 "메일함" 대신 씁니다.

만드는 메일: 3일치 x 3종류 = 9통
  - exception : 전일 Exception 현황
  - batch     : 배치 오류
  - ams       : AMS 이벤트 현황

날짜별로 내용을 조금씩 다르게 넣어 두었습니다.
  - 매일 반복되는 노이즈 exception (무시 대상)
  - 10/26에 처음 나타난 exception (신규 이슈)
  - 10/25 -> 10/26 건수가 급증한 exception
이렇게 해 두어야 "처음 나온 거야?", "늘었어?" 같은 질문을 평가셋에 넣을 수 있습니다.

실행:
    python make_dummy_mails.py
결과:
    miniprj/data/mails/*.eml
"""

from __future__ import annotations

from datetime import datetime, timedelta
from email.message import EmailMessage
from email.utils import format_datetime
from pathlib import Path

OUT_DIR = Path(__file__).resolve().parent / "data" / "mails"
SENDER = "monitoring@example.com"
RECEIVER = "ops@example.com"
NOTICE = "※ 교육용 더미 데이터입니다. 실제 시스템과 관계없습니다."

# ── exception 메일 본문의 행: (시스템, Exception, 최초 발생 시각, 건수) ──
EXCEPTION_ROWS = {
    "2026-10-24": [
        ("주문시스템", "java.sql.SQLTimeoutException: Connection timed out", "23:41", 18),
        ("회원시스템", "java.lang.NullPointerException at MemberService.getGrade", "10:12", 3),
        ("공통", "org.apache.catalina.connector.ClientAbortException: Broken pipe", "00:05", 210),
    ],
    "2026-10-25": [
        ("주문시스템", "java.sql.SQLTimeoutException: Connection timed out", "22:58", 21),
        ("공통", "org.apache.catalina.connector.ClientAbortException: Broken pipe", "00:03", 198),
    ],
    "2026-10-26": [
        # 전날 21건 -> 142건 급증
        ("주문시스템", "java.sql.SQLTimeoutException: Connection timed out", "23:41", 142),
        # 처음 나타난 exception
        ("결제시스템", "javax.net.ssl.SSLHandshakeException: PKIX path validation failed", "02:17", 37),
        ("공통", "org.apache.catalina.connector.ClientAbortException: Broken pipe", "00:01", 205),
    ],
}

# ── batch 메일 본문의 행: (배치 ID, 배치명, 실패 시각, 오류 코드, 메시지) ──
BATCH_ROWS = {
    "2026-10-24": [],
    "2026-10-25": [
        ("BAT-ORD-010", "주문 일마감 집계", "01:30", "E-DB-001", "Deadlock detected, 재시도 후 성공"),
    ],
    "2026-10-26": [
        ("BAT-PAY-021", "결제 정산 파일 생성", "02:20", "E-NET-404", "외부 정산 서버 연결 실패 (SSL)"),
        ("BAT-MBR-003", "휴면 회원 전환", "03:00", "E-APP-500", "NullPointerException, 0건 처리 후 종료"),
    ],
}

# ── ams 메일 본문의 행: (등급, 호스트, 항목, 값, 임계치, 상태) ──
AMS_ROWS = {
    "2026-10-24": [
        ("Warning", "ord-was-01", "CPU", "82%", "80%", "해소"),
    ],
    "2026-10-25": [
        ("Warning", "ord-was-01", "CPU", "85%", "80%", "해소"),
        ("Info", "batch-01", "Disk /data", "71%", "70%", "지속"),
    ],
    "2026-10-26": [
        ("Critical", "ord-db-01", "DB Connection Pool", "98%", "90%", "지속"),
        ("Warning", "pay-was-02", "Cert Expiry", "D-0", "D-30", "지속"),
        ("Info", "batch-01", "Disk /data", "73%", "70%", "지속"),
    ],
}


def _table(header: str, rows: list[tuple]) -> str:
    """행 목록을 '|'로 구분한 표 문자열로 만듭니다."""
    if not rows:
        return header + "\n(해당 없음)"
    return header + "\n" + "\n".join(" | ".join(str(c) for c in row) for row in rows)


def _write(day: str, kind: str, subject: str, body: str, hour: int) -> Path:
    """메일 한 통을 .eml 파일로 저장합니다. 받은 시각은 다음 날 아침으로 둡니다."""
    msg = EmailMessage()
    msg["From"] = SENDER
    msg["To"] = RECEIVER
    msg["Subject"] = subject
    # 전일 기준 메일이므로 다음 날 아침에 도착한 것으로 둡니다.
    # (day + 1로 계산하면 10/31 같은 월말에 오류가 나므로 timedelta로 하루를 더합니다.)
    sent = (datetime.fromisoformat(day).replace(hour=hour) + timedelta(days=1)).astimezone()
    msg["Date"] = format_datetime(sent)
    msg.set_content(f"{NOTICE}\n\n{body}\n")
    path = OUT_DIR / f"{day}_{kind}.eml"
    path.write_bytes(bytes(msg))
    return path


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for day in EXCEPTION_ROWS:
        _write(
            day, "exception",
            f"[Exception Report] {day} 전일 Exception 현황",
            _table("시스템 | Exception | 최초발생 | 건수", EXCEPTION_ROWS[day]),
            hour=7,
        )
        _write(
            day, "batch",
            f"[Batch Error] {day} 배치 오류 현황",
            _table("배치ID | 배치명 | 실패시각 | 오류코드 | 메시지", BATCH_ROWS[day]),
            hour=6,
        )
        _write(
            day, "ams",
            f"[AMS Event] {day} 이벤트 현황",
            _table("등급 | 호스트 | 항목 | 값 | 임계치 | 상태", AMS_ROWS[day]),
            hour=8,
        )
    print(f"{OUT_DIR} 에 더미 메일 {len(list(OUT_DIR.glob('*.eml')))}통을 만들었습니다.")


if __name__ == "__main__":
    main()
