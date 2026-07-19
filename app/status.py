import json
import os
from datetime import date, timedelta

DATA_DIR = os.environ.get('HUGO_DATA_DIR') or os.path.join(
    os.path.expanduser('~'), '.hugo-data'
)
STATUS_FILE = os.path.join(DATA_DIR, 'processed_days.json')


def load_processed_days():
    try:
        with open(STATUS_FILE) as f:
            return set(json.load(f))
    except (OSError, ValueError):
        return set()


def mark_day_processed(day):
    days = load_processed_days()
    days.add(day.isoformat())
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(STATUS_FILE, 'w') as f:
        json.dump(sorted(days), f)


def last_days(n=45):
    today = date.today()
    return [today - timedelta(days=i) for i in range(n - 1, -1, -1)]
