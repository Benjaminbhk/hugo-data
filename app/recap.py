import re
import pandas as pd

MONTH_CODES = {
    'F': 1, 'G': 2, 'H': 3, 'J': 4, 'K': 5, 'M': 6,
    'N': 7, 'Q': 8, 'U': 9, 'V': 10, 'X': 11, 'Z': 12,
}
MONTH_LETTERS = {v: k for k, v in MONTH_CODES.items()}

# Ordre d'affichage du recap voulu par Hugo : 4 groupes par importance,
# separes par une ligne vide. Un ticker hors liste est rejete en fin de recap.
TICKER_GROUPS = [
    ['ZTW'],
    ['FKB', 'ZSI', 'FMI', 'ZTO', 'ZSS', 'FPP', 'ZTI'],
    ['ZVL', 'MUR', 'FPO', 'JJY'],
    ['ZSR', 'ZVO', 'ZWO', 'ZVW'],
]
TICKER_RANK = {
    root: (gi, ti)
    for gi, group in enumerate(TICKER_GROUPS)
    for ti, root in enumerate(group)
}
UNRANKED = (len(TICKER_GROUPS), 0)

LEG_RE = re.compile(r'^([A-Z]{3})([FGHJKMNQUVXZ])(\d)$')
L0_RE = re.compile(r'^([A-Z]{3})([FGHJKMNQUVXZ])(\d)([FGHJKMNQUVXZ])(\d)$')

DEDUP_KEY = ['Ticker', 'Time', 'Size', 'Price']


def row_keys(df):
    """
    Clé d'identité d'un trade, robuste aux allers-retours CSV/Excel :
    les numériques sont normalisés (2133 == 2133.0), les textes nettoyés.
    """
    parts = []
    for col in DEDUP_KEY:
        if col not in df.columns:
            continue
        s = df[col]
        if col in ('Size', 'Price'):
            s = pd.to_numeric(s, errors='coerce').round(6)
        parts.append(s.astype(str).str.strip())
    return parts[0].str.cat(parts[1:], sep='|')


def dedupe_dataframes(dataframes):
    """
    Fusionne plusieurs exports Bloomberg en supprimant les doublons.

    Bloomberg plafonne chaque export à 100 lignes : une même journée arrive
    en plusieurs fichiers qui se chevauchent. Un même trade garde (Ticker,
    Time, Size, Price) constants d'un export à l'autre, alors que Volume,
    1DChg et UndPrc dérivent au fil de la journée — on garde la dernière
    occurrence. Retourne (DataFrame fusionné, nombre de doublons supprimés).
    """
    merged = pd.concat(dataframes, ignore_index=True)
    # Normalisation : les lignes reprises de la mémoire (CSV) et celles des
    # exports Excel doivent produire la même clé de déduplication
    if 'Ticker' in merged.columns:
        merged['Ticker'] = merged['Ticker'].astype(str).str.strip()
    if 'Time' in merged.columns:
        merged['Time'] = merged['Time'].astype(str).str.strip()
    before = len(merged)
    merged = merged[~row_keys(merged).duplicated(keep='last')].reset_index(drop=True)
    return merged, before - len(merged)


def _expiry(month_code, year_digit, trade_year):
    year = trade_year - trade_year % 10 + int(year_digit)
    if year < trade_year - 2:
        year += 10
    return year, MONTH_CODES[month_code]


def _code(year, month):
    # Encodage Bloomberg d'une echeance : lettre du mois + dernier chiffre
    # de l'annee (mars 2026 -> H6)
    return f'{MONTH_LETTERS[month]}{str(year)[-1]}'


def _fmt_notional(value):
    if value >= 1e9:
        s = f'{value / 1e9:.1f}'.rstrip('0').rstrip('.')
        return f'${s}bn'
    return f'${value / 1e6:.0f}m'


def _fmt_level(value):
    # Les spreads sont cotés par pas de 0.0025 : arrondir au pas élimine
    # le bruit des prix moyens (1.2499 -> 1.25)
    quantized = round(round(value / 0.0025) * 0.0025, 4)
    return f'{quantized:g}'


def build_recap(df, trade_date):
    """
    Construit le texte 'Recap MSCI Rolls' à partir du DataFrame traité
    par process_data : une ligne par (sous-jacent, paire d'échéances) avec
    les niveaux distincts traités et le notional total (moyenne des legs).
    """
    trade_date = pd.Timestamp(trade_date)
    trade_year = trade_date.year
    groups = {}
    others = []

    def add(key, level, notional):
        entry = groups.setdefault(key, {'levels': [], 'notional': 0.0})
        if level is not None:
            lvl = _fmt_level(level)
            if lvl not in entry['levels'] and lvl != '0':
                entry['levels'].append(lvl)
        entry['notional'] += notional

    rolls = df[df['Structure_ID'].astype(str).str.contains('-R-', na=False)].copy()
    rolls['_roll'] = rolls['Structure_ID'].astype(str).str.extract(r'(\d{8}-R-\d+)')

    for _, grp in rolls.groupby('_roll'):
        legs = grp[grp['Structure'] == 'Leg']
        if len(legs) == 2:
            parsed = []
            for _, leg in legs.iterrows():
                m = LEG_RE.match(str(leg['Ticker']).strip())
                if m:
                    root, mc, yd = m.groups()
                    parsed.append((_expiry(mc, yd, trade_year), root, leg))
            if len(parsed) != 2:
                others.append(grp)
                continue
            parsed.sort(key=lambda p: p[0])
            (near_exp, root, near), (far_exp, _, far) = parsed
            if near_exp == far_exp:
                continue
            level = (far['Price'] / near['Price'] - 1) * 100
            notional = (near['Notional'] + far['Notional']) / 2
            key = (root, _code(*near_exp), _code(*far_exp))
            add(key, level, notional)
        else:
            single = grp[grp['Structure'].isin(['Roll', 'Roll Client'])]
            for _, row in single.iterrows():
                m = L0_RE.match(str(row['Ticker']).strip())
                if not m:
                    others.append(grp)
                    break
                root, m1, y1, m2, y2 = m.groups()
                e1 = _expiry(m1, y1, trade_year)
                e2 = _expiry(m2, y2, trade_year)
                (near_exp, far_exp) = sorted([e1, e2])
                price = row['Price']
                level = price if pd.notna(price) and 0 < abs(price) < 20 else None
                key = (root, _code(*near_exp), _code(*far_exp))
                add(key, level, row['Notional'])

    lines = [f"Recap MSCI Rolls {trade_date.strftime('%d/%m/%Y')}", '']
    ordered = sorted(
        groups.items(),
        key=lambda kv: (TICKER_RANK.get(kv[0][0], UNRANKED), -kv[1]['notional']),
    )
    current_group = None
    for (root, near, far), entry in ordered:
        group = TICKER_RANK.get(root, UNRANKED)[0]
        # Ligne vide entre deux groupes presents : un groupe sans trade du
        # jour ne laisse pas de trou dans le recap
        if current_group is not None and group != current_group:
            lines.append('')
        current_group = group
        levels = ' + '.join(entry['levels']) if entry['levels'] else '-'
        lines.append(f"{root}{near}{far} @ {levels} {_fmt_notional(entry['notional'])}")

    n_others = sum(len(g) for g in others)
    if n_others:
        lines.append('')
        lines.append(f'({n_others} lignes de roll non classées — tickers non reconnus)')
    return '\n'.join(lines)
