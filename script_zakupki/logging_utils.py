from datetime import datetime
from pathlib import Path

from .settings import REPORTS_DIR


class TeeOutput:
    def __init__(self, *streams):
        self.streams = streams

    def write(self, message):
        for stream in self.streams:
            stream.write(message)
            stream.flush()

    def flush(self):
        for stream in self.streams:
            stream.flush()


def create_log_path(reports_dir=REPORTS_DIR):
    reports_dir = Path(reports_dir)
    reports_dir.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return reports_dir / f"zakupki_{timestamp}.txt"
