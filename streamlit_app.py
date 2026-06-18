import streamlit as st
import pandas as pd
from app.processing import load_dataframes, process_data, build_excel


def main():
    st.title("Application de traitement des fichiers Excel")

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

                processed_df, error = process_data(dataframes, pd.to_datetime(trade_date))
                if error:
                    st.error(error)
                    return

                download_buffer = build_excel(processed_df)
                filename_date = pd.to_datetime(trade_date).strftime('%Y%m%d')

                st.success("Traitement terminé!")
                st.download_button(
                    label="Télécharger le fichier traité",
                    data=download_buffer,
                    file_name=f"{filename_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.info("💡 Le fichier sera automatiquement téléchargé dans le dossier 'Téléchargements'.")
                st.dataframe(processed_df)


if __name__ == "__main__":
    main()
