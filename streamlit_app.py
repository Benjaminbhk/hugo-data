import pandas as pd
import streamlit as st
from app.processing import load_dataframes, process_data, build_excel
from app.recap import dedupe_dataframes, build_recap
from app.status import load_processed_days, mark_day_processed, last_days

CHANGELOG = [
    ("19/07/2026", "Recap MSCI Rolls généré automatiquement, prêt à copier-coller"),
    ("19/07/2026", "Déduplication des trades entre fichiers qui se chevauchent"),
    ("19/07/2026", "Suivi des 45 derniers jours traités en haut de page"),
    ("19/07/2026", "Le résultat reste affiché après le téléchargement"),
]

WEEKDAY_LABELS = ['L', 'M', 'M', 'J', 'V', 'S', 'D']


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
            f'style="width:26px;height:26px;border-radius:4px;background:{color};'
            f'color:{text_color};font-size:11px;display:flex;align-items:center;'
            f'justify-content:center;">{day.day}</div>'
        )
    st.markdown(
        '<div style="display:flex;flex-wrap:wrap;gap:3px;margin-bottom:1rem;">'
        + ''.join(cells) + '</div>',
        unsafe_allow_html=True,
    )


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

                merged, n_dupes = dedupe_dataframes(dataframes)
                processed_df, error = process_data([merged], pd.to_datetime(trade_date))
                if error:
                    st.error(error)
                    return

                st.session_state['result'] = {
                    'date': trade_date,
                    'excel': build_excel(processed_df).getvalue(),
                    'recap': build_recap(processed_df, trade_date),
                    'n_dupes': n_dupes,
                    'df': processed_df,
                }
                mark_day_processed(trade_date)
            st.rerun()

    result = st.session_state.get('result')
    if result:
        filename_date = pd.to_datetime(result['date']).strftime('%Y%m%d')
        st.success(f"Traitement du {result['date'].strftime('%d/%m/%Y')} terminé !")
        if result['n_dupes']:
            st.info(
                f"🧹 {result['n_dupes']} doublon(s) supprimé(s) "
                "(trades présents dans plusieurs fichiers)."
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
