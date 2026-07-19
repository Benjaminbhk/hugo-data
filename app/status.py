import json
import os

import pandas as pd

from app.recap import row_keys

DATA_DIR = os.environ.get('HUGO_DATA_DIR') or os.path.join(
    os.path.expanduser('~'), '.hugo-data'
)
STATUS_FILE = os.path.join(DATA_DIR, 'processed_days.json')
LEGS_DIR = os.path.join(DATA_DIR, 'legs')

# Une même ligne (Ticker, Time, Size, Price) peut légitimement se reproduire
# un autre jour : on ne considère un chevauchement avec un jour déjà traité
# comme réel que si plusieurs lignes coïncident.
MIN_CROSS_DAY_MATCHES = 3


def load_processed_days():
    try:
        with open(STATUS_FILE) as f:
            return set(json.load(f))
    except (OSError, ValueError):
        return set()


def _save_processed_days(days):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(STATUS_FILE, 'w') as f:
        json.dump(sorted(days), f)


def mark_day_processed(day):
    days = load_processed_days()
    days.add(day.isoformat())
    _save_processed_days(days)


def unmark_day_processed(day_iso):
    days = load_processed_days()
    days.discard(day_iso)
    _save_processed_days(days)
    try:
        os.remove(_legs_file(day_iso))
    except OSError:
        pass


def _legs_file(day_iso):
    return os.path.join(LEGS_DIR, f'{day_iso}.csv')


def save_day_legs(day, df):
    os.makedirs(LEGS_DIR, exist_ok=True)
    df.to_csv(_legs_file(day.isoformat()), index=False)


def days_with_legs():
    """Jours ayant des lignes en mémoire, triés du plus récent au plus ancien."""
    if not os.path.isdir(LEGS_DIR):
        return []
    return sorted(
        (f[:-4] for f in os.listdir(LEGS_DIR) if f.endswith('.csv')),
        reverse=True,
    )


def load_day_legs(day):
    """Lignes déjà traitées pour ce jour, ou None."""
    try:
        return pd.read_csv(_legs_file(day.isoformat()))
    except (OSError, ValueError):
        return None


def filter_cross_day_duplicates(df, day):
    """
    Écarte les lignes déjà mémorisées pour un AUTRE jour que `day`.
    Un chevauchement n'est retenu que s'il porte sur au moins
    MIN_CROSS_DAY_MATCHES lignes (une coïncidence isolée est conservée).
    Retourne (df filtré, [(jour, nb de lignes écartées), ...]).
    """
    if not os.path.isdir(LEGS_DIR):
        return df, []
    removals = []
    for fname in sorted(os.listdir(LEGS_DIR)):
        other_day = fname[:-4]
        if not fname.endswith('.csv') or other_day == day.isoformat():
            continue
        try:
            other = pd.read_csv(_legs_file(other_day))
        except (OSError, ValueError):
            continue
        if other.empty or df.empty:
            continue
        mask = row_keys(df).isin(set(row_keys(other)))
        if mask.sum() >= MIN_CROSS_DAY_MATCHES:
            removals.append((other_day, int(mask.sum())))
            df = df[~mask]
    return df.reset_index(drop=True), removals
