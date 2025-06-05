from dataclasses import dataclass

@dataclass
class MailMessage:
    id: str
    subject: str
    body: str
    received: str