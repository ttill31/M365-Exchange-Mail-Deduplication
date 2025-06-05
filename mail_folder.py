from dataclasses import dataclass

@dataclass
class MailFolder:
    id: str
    display_name: str
    total_count: int