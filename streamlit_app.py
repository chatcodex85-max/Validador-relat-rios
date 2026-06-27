from __future__ import annotations

import csv
import tempfile
from io import BytesIO, StringIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook

from app import ValidationErrorRow, build_export_row, columns_for_export, validate_workbook


st.set_page_config(page_title="Validador de Relatorios XLSX", layout="wide")


def validate_uploads(uploaded_files) -> list[ValidationErrorRow]:
    results: list[ValidationErrorRow] = []

    for uploaded_file in uploaded_files:
        suffix = Path(uploaded_file.name).suffix or ".xlsx"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
            temp_file.write(uploaded_file.getbuffer())
            temp_path = Path(temp_file.name)

        try:
            results.extend(validate_workbook(temp_path))
        finally:
            temp_path.unlink(missing_ok=True)

    return results


def rows_to_dataframe(errors: list[ValidationErrorRow]) -> pd.DataFrame:
    columns = columns_for_export(errors)
    rows = [build_export_row(error, columns) for error in errors]
    return pd.DataFrame(rows, columns=columns)


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    buffer = StringIO()
    writer = csv.writer(buffer, delimiter=";", lineterminator="\n")
    writer.writerow(df.columns)
    writer.writerows(df.values.tolist())
    return buffer.getvalue().encode("utf-8-sig")


def dataframe_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Resultado"
    sheet.append(list(df.columns))

    for row in df.itertuples(index=False, name=None):
        sheet.append(list(row))

    for column in sheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))
        sheet.column_dimensions[column_letter].width = min(max_length + 3, 60)

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    filtered = df.copy()

    search = st.session_state.get("search", "").strip().lower()
    if search:
        mask = filtered.astype(str).apply(lambda row: search in " ".join(row).lower(), axis=1)
        filtered = filtered[mask]

    filter_columns = [column for column in ("UF", "Cidade", "Erro", "Tag", "Campo") if column in filtered.columns]
    filter_layout = st.columns(max(len(filter_columns), 1))

    for index, column in enumerate(filter_columns):
        values = ["Todos"] + sorted(value for value in filtered[column].dropna().astype(str).unique() if value)
        selected = filter_layout[index].selectbox(column, values, key=f"filter_{column}")
        if selected != "Todos":
            filtered = filtered[filtered[column].astype(str) == selected]

    return filtered


st.title("Validador de Relatorios XLSX")

uploaded_files = st.file_uploader(
    "Selecione um ou mais relatorios .xlsx",
    type=["xlsx"],
    accept_multiple_files=True,
)

st.text_input("Buscar", key="search", placeholder="Buscar por ID, cidade, nome, erro ou tag")

if not uploaded_files:
    st.info("Envie um arquivo .xlsx para iniciar a validacao.")
    st.stop()

with st.spinner("Validando relatorios..."):
    errors = validate_uploads(uploaded_files)

if not errors:
    st.success("Nenhum erro encontrado.")
    st.stop()

df = rows_to_dataframe(errors)
filtered_df = apply_filters(df)

metric_cols = st.columns(4)
metric_cols[0].metric("Itens encontrados", len(df))
metric_cols[1].metric("Itens exibidos", len(filtered_df))
if "Erro" in df.columns:
    metric_cols[2].metric("Tipos de erro", df["Erro"].nunique())
if "Tag" in df.columns:
    metric_cols[3].metric("Tags", df["Tag"].replace("", pd.NA).dropna().nunique())

st.dataframe(filtered_df, use_container_width=True, hide_index=True)

download_cols = st.columns(2)
download_cols[0].download_button(
    "Baixar CSV",
    dataframe_to_csv_bytes(filtered_df),
    file_name="resultado_validacao.csv",
    mime="text/csv",
)
download_cols[1].download_button(
    "Baixar XLSX",
    dataframe_to_xlsx_bytes(filtered_df),
    file_name="resultado_validacao.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
