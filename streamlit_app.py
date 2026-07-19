from datetime import date

import pandas as pd
import streamlit as st
from app.processing import load_dataframes, process_data, build_excel
from app.recap import dedupe_dataframes, build_recap
from app.status import (
    load_processed_days, mark_day_processed, unmark_day_processed, last_days,
    load_day_legs, save_day_legs, filter_cross_day_duplicates,
)

CHANGELOG = [
    ("20/07/2026", "Mémoire des lignes traitées : relancer un jour cumule sans doublon"),
    ("20/07/2026", "Détection des lignes déjà traitées un autre jour"),
    ("20/07/2026", "Cases du calendrier décochables (mois affiché dans la case)"),
    ("19/07/2026", "Recap MSCI Rolls généré automatiquement, prêt à copier-coller"),
    ("19/07/2026", "Déduplication des trades entre fichiers qui se chevauchent"),
    ("19/07/2026", "Suivi des 45 derniers jours traités en haut de page"),
    ("19/07/2026", "Le résultat reste affiché après le téléchargement"),
]

WEEKDAY_LABELS = ['L', 'M', 'M', 'J', 'V', 'S', 'D']
MONTH_FR = ['janv', 'févr', 'mars', 'avr', 'mai', 'juin',
            'juil', 'août', 'sept', 'oct', 'nov', 'déc']


def render_changelog():
    if st.session_state.get('hide_changelog'):
        return
    with st.sidebar:
        head, close = st.columns([5, 1])
        head.subheader("Nouveautés")
        if close.button("✕", help="Fermer"):
            st.session_state['hide_changelog'] = True
            st.rerun()
        for day, text in CHANGELOG:
            st.markdown(f"**{day}** — {text}")


def render_status_grid():
    processed = load_processed_days()
    cells = []
    for day in last_days(45):
        done = day.isoformat() in processed
        weekend = day.weekday() >= 5
        color = '#21b558' if done else ('#f0f0f0' if weekend else '#dcdcdc')
        text_color = 'white' if done else '#888'
        cells.append(
            f'<div title="{WEEKDAY_LABELS[day.weekday()]} {day.strftime("%d/%m/%Y")}" '
            f'style="width:34px;height:34px;border-radius:4px;background:{color};'
            f'color:{text_color};font-size:12px;line-height:1.1;display:flex;'
            f'flex-direction:column;align-items:center;justify-content:center;">'
            f'{day.day}<span style="font-size:8px;">{MONTH_FR[day.month - 1]}</span></div>'
        )
    st.markdown(
        '<div style="display:flex;flex-wrap:wrap;gap:3px;margin-bottom:0.5rem;">'
        + ''.join(cells) + '</div>',
        unsafe_allow_html=True,
    )

    greens = sorted(d for d in processed)
    if greens:
        with st.popover("✏️ Décocher un jour"):
            st.caption(
                "Décocher un jour retire le vert et efface la mémoire "
                "des lignes traitées de ce jour."
            )
            to_remove = st.multiselect(
                "Jours traités",
                greens,
                format_func=lambda d: date.fromisoformat(d).strftime('%d/%m/%Y'),
            )
            if to_remove and st.button("Décocher"):
                for d in to_remove:
                    unmark_day_processed(d)
                st.rerun()


def main():
    st.set_page_config(page_title="Hugo Data", layout="wide")
    render_changelog()

    st.title("Application de traitement des fichiers Excel")
    st.caption("Jours traités sur les 45 derniers jours :")
    render_status_grid()

    uploaded_files = st.file_uploader(
        "Sélectionnez un ou plusieurs fichiers Excel",
        type=["xlsx"],
        accept_multiple_files=True,
    )
    trade_date = st.date_input("Sélectionnez la date")

    if uploaded_files and trade_date:
        if st.button("Traiter les fichiers"):
            with st.spinner("Traitement en cours..."):
                errors = []
                dataframes = load_dataframes(
                    uploaded_files,
                    on_error=lambda msg: errors.append(msg),
                )
                for err in errors:
                    st.error(err)

                if not dataframes:
                    st.error("Aucun fichier valide n'a été chargé.")
                    return

                previous = load_day_legs(trade_date)
                n_previous = len(previous) if previous is not None else 0
                if previous is not None:
                    dataframes = [previous] + dataframes

                merged, n_dupes = dedupe_dataframes(dataframes)
                merged, cross_days = filter_cross_day_duplicates(merged, trade_date)
                if merged.empty:
                    st.error(
                        "Toutes les lignes ont déjà été traitées un autre jour."
                    )
                    return

                processed_df, error = process_data([merged], pd.to_datetime(trade_date))
                if error:
                    st.error(error)
                    return

                save_day_legs(trade_date, merged)
                st.session_state['result'] = {
                    'date': trade_date,
                    'excel': build_excel(processed_df).getvalue(),
                    'recap': build_recap(processed_df, trade_date),
                    'n_dupes': n_dupes,
                    'n_previous': n_previous,
                    'cross_days': cross_days,
                    'df': processed_df,
                }
                mark_day_processed(trade_date)
            st.rerun()

    result = st.session_state.get('result')
    if result:
        filename_date = pd.to_datetime(result['date']).strftime('%Y%m%d')
        st.success(f"Traitement du {result['date'].strftime('%d/%m/%Y')} terminé !")
        if result['n_previous']:
            st.info(
                f"📚 {result['n_previous']} ligne(s) reprise(s) de la mémoire du jour "
                "et fusionnée(s) avec les nouveaux fichiers."
            )
        if result['n_dupes']:
            st.info(f"🧹 {result['n_dupes']} doublon(s) supprimé(s).")
        for other_day, count in result['cross_days']:
            st.warning(
                f"⚠️ {count} ligne(s) écartée(s) : déjà traitée(s) le "
                f"{date.fromisoformat(other_day).strftime('%d/%m/%Y')}."
            )
        st.download_button(
            label="Télécharger le fichier traité",
            data=result['excel'],
            file_name=f"{filename_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.subheader("Recap")
        st.code(result['recap'], language=None)
        st.dataframe(result['df'])


if __name__ == "__main__":
    main()
