"""[연동 확인용] Outlook 받은편지함에서 운영 메일을 찾아 목록만 보여 줍니다.

업무 PC에서 Outlook으로 회사 메일을 읽을 수 있는지 확인하는 스크립트입니다.

이 스크립트가 하는 일
  - 받은편지함에서 최근 N일 동안 받은 메일 중 제목이 규칙에 맞는 메일을 찾습니다.
  - 메일 종류, 받은 시각, 제목, 본문 길이만 출력합니다. (본문 내용은 출력하지 않습니다.)

이 스크립트가 하지 않는 일
  - 메일을 지우거나, 옮기거나, 읽음 표시를 바꾸거나, 파일로 저장하지 않습니다.
  - LLM을 호출하지 않습니다. 메일 내용은 이 PC 밖으로 나가지 않습니다.

준비
  1. 업무 PC의 Outlook에 회사 메일 계정이 연결되어 있어야 합니다.
  2. 아래 MAIL_TYPES의 제목을 실제 메일 제목의 "고정된 부분"으로 바꾸세요.

실행 (가상환경 파이썬으로):
    C:\\Users\\SDSPRO\\Desktop\\.venv\\Scripts\\python.exe outlook_check.py
    C:\\Users\\SDSPRO\\Desktop\\.venv\\Scripts\\python.exe outlook_check.py 7   # 최근 7일
"""

from __future__ import annotations

import sys
from datetime import datetime, timedelta

# 메일 종류별 제목의 고정 부분입니다. 날짜를 뺀 나머지를 그대로 적으세요.
#   예) 실제 제목이 "[Exception Report] 2026-10-26 전일 현황" 이면 "[Exception Report]" 만 적습니다.
# 아래 값은 예시입니다. 실제 제목으로 바꿔야 메일을 찾습니다.
MAIL_TYPES = {
    "exception": "[Exception Report]",
    "batch": "[Batch Error]",
    "ams": "[AMS Event]",
}

# Outlook 상수: 받은편지함 폴더 번호, 일반 메일 항목의 종류 번호
OL_FOLDER_INBOX = 6
OL_MAIL_ITEM = 43


def classify_subject(subject: str) -> str | None:
    """제목으로 메일 종류를 정합니다. LLM 없이 규칙만으로 판단합니다."""
    for mail_type, fixed_part in MAIL_TYPES.items():
        if fixed_part in subject:
            return mail_type
    return None


def main(days: int) -> None:
    try:
        import win32com.client
    except ImportError:
        sys.exit("pywin32가 없습니다. 가상환경에서 `pip install pywin32` 후 다시 실행하세요.")

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(OL_FOLDER_INBOX)
    except Exception as exc:  # noqa: BLE001
        sys.exit(
            f"Outlook에 연결하지 못했습니다: {exc}\n"
            "  Outlook이 설치되어 있고 회사 메일 계정이 연결되어 있는지 확인하세요."
        )

    cutoff = (datetime.now() - timedelta(days=days)).date()
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)  # 최신 메일부터 봅니다.

    found: dict[str, list[tuple[str, str, int]]] = {t: [] for t in MAIL_TYPES}
    checked = 0
    for item in items:
        # 회의 요청 같은 메일이 아닌 항목은 건너뜁니다.
        if getattr(item, "Class", None) != OL_MAIL_ITEM:
            continue
        received = item.ReceivedTime
        # 최신순으로 정렬했으므로 기준일보다 오래된 메일이 나오면 더 볼 필요가 없습니다.
        if received.date() < cutoff:
            break
        checked += 1
        mail_type = classify_subject(item.Subject or "")
        if mail_type:
            found[mail_type].append(
                (received.strftime("%Y-%m-%d %H:%M"), item.Subject, len(item.Body or ""))
            )

    print(f"받은편지함 '{inbox.Name}'에서 {cutoff} 이후 메일 {checked}통을 확인했습니다.\n")
    for mail_type, mails in found.items():
        print(f"[{mail_type}] 제목 규칙 '{MAIL_TYPES[mail_type]}' → {len(mails)}통")
        for received, subject, body_len in mails:
            print(f"    {received}  {subject}  (본문 {body_len}자)")
    if not any(found.values()):
        print("\n규칙에 맞는 메일이 없습니다. MAIL_TYPES의 제목을 실제 제목과 비교해 보세요.")


if __name__ == "__main__":
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 1)
