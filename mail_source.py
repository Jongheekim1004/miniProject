"""운영 메일을 읽어 오는 부분입니다. "메일함"을 바꿔 끼울 수 있게 만들어 둡니다.

  - 교육 PC : load_from_eml_folder()  -> 더미 .eml 폴더에서 읽음
  - 업무 PC : load_from_outlook()     -> Outlook 받은편지함에서 읽음 (읽기만 함)

두 함수는 똑같이 Mail 목록을 돌려줍니다. 그래서 뒤에 붙는 분류·요약·RAG 코드는
어느 메일함에서 읽었는지 몰라도 그대로 동작합니다.
(Day 1에서 llm=None으로 진짜 모델과 가짜 모델을 바꿔 끼운 것과 같은 생각입니다.)
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from pathlib import Path

# 메일 종류별 제목의 고정 부분입니다. 제목은 날짜만 다르고 나머지는 항상 같습니다.
# 업무 PC에서는 실제 제목의 고정 부분으로 바꾸세요.
MAIL_TYPES = {
    "exception": "[Exception Report]",
    "batch": "[Batch Error]",
    "ams": "[AMS Event]",
}

DEFAULT_EML_DIR = Path(__file__).resolve().parent / "data" / "mails"


@dataclass
class Mail:
    """메일 한 통. 어느 메일함에서 읽었든 이 모양으로 맞춥니다."""

    mail_type: str       # exception / batch / ams
    received: datetime   # 받은 시각
    subject: str
    body: str


def classify_subject(subject: str) -> str | None:
    """제목으로 메일 종류를 정합니다.

    제목 규칙이 고정이므로 LLM에게 묻지 않고 규칙으로 판단합니다.
    (정해진 답이 있는 판단은 규칙으로 - Day 1 심화와 같은 원칙)
    """
    for mail_type, fixed_part in MAIL_TYPES.items():
        if fixed_part in subject:
            return mail_type
    return None


def load_from_eml_folder(folder: Path = DEFAULT_EML_DIR, since: date | None = None) -> list[Mail]:
    """[교육 PC] 폴더의 .eml 파일을 읽어 운영 메일만 돌려줍니다.

    Args:
        folder: .eml 파일이 있는 폴더
        since: 이 날짜 이후에 받은 메일만 돌려줍니다. None이면 전부.
    """
    mails: list[Mail] = []
    for path in sorted(Path(folder).glob("*.eml")):
        with open(path, "rb") as f:
            msg = BytesParser(policy=policy.default).parse(f)
        mail_type = classify_subject(msg["Subject"] or "")
        if mail_type is None:
            continue
        received = parsedate_to_datetime(msg["Date"])
        if since and received.date() < since:
            continue
        body = msg.get_body(preferencelist=("plain", "html")).get_content()
        mails.append(Mail(mail_type, received, msg["Subject"], body))
    return sorted(mails, key=lambda m: m.received)


def load_from_outlook(days: int = 1) -> list[Mail]:
    """[업무 PC] Outlook 받은편지함에서 최근 days일 동안 받은 운영 메일을 돌려줍니다.

    메일을 읽기만 합니다. 지우기·옮기기·읽음 표시 변경·저장을 하지 않습니다.
    """
    import win32com.client  # 업무 PC에서만 필요하므로 함수 안에서 불러옵니다.

    inbox = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(6)
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)  # 최신 메일부터
    cutoff = (datetime.now() - timedelta(days=days)).date()

    mails: list[Mail] = []
    for item in items:
        if getattr(item, "Class", None) != 43:  # 43 = 일반 메일. 회의 요청 등은 건너뜀
            continue
        received = item.ReceivedTime
        if received.date() < cutoff:
            break  # 최신순이므로 더 오래된 메일은 볼 필요 없음
        mail_type = classify_subject(item.Subject or "")
        if mail_type:
            mails.append(Mail(mail_type, received, item.Subject, item.Body or ""))
    return sorted(mails, key=lambda m: m.received)


if __name__ == "__main__":
    # 교육 PC에서 더미 메일함을 읽어 목록을 보여 줍니다.
    for m in load_from_eml_folder():
        print(f"[{m.mail_type:9}] {m.received:%Y-%m-%d %H:%M}  {m.subject}")
