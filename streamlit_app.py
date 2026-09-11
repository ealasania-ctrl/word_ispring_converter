from __future__ import annotations

import json
from io import BytesIO
from pathlib import Path
import tempfile
import zipfile

import streamlit as st

from docx_to_ispring_excel import convert

st.set_page_config(page_title="DOCX -> iSpring Excel", page_icon="📘", layout="centered")

st.title("Конвертер Word тестов в iSpring Excel")
st.write(
    "Загрузите файл .docx с тестами. Приложение создаст Excel-файлы для импорта в iSpring "
    "и отдаст их одним архивом."
)

uploaded_docx = st.file_uploader("Файл Word (.docx)", type=["docx"])

split_mode = st.selectbox(
    "Режим разбиения",
    options=["task", "type", "section-type"],
    index=0,
    help=(
        "task: 1 Task = 1 Excel файл, строго один тип в каждом Task. "
        "type: группировка только по типу TF/MC/TI. "
        "section-type: группировка по разделу из JSON и типу."
    ),
)

uploaded_sections = None
if split_mode == "section-type":
    uploaded_sections = st.file_uploader(
        "JSON конфиг разделов",
        type=["json"],
        help='Формат: {"Раздел 1": [1,2], "Раздел 2": [3,4]}',
    )

run_clicked = st.button("Сконвертировать", type="primary", use_container_width=True)

if run_clicked:
    if uploaded_docx is None:
        st.error("Сначала загрузите .docx файл.")
    elif split_mode == "section-type" and uploaded_sections is None:
        st.error("Для режима section-type загрузите JSON конфиг разделов.")
    else:
        try:
            with tempfile.TemporaryDirectory() as tmp_root:
                tmp_root_path = Path(tmp_root)
                input_docx_path = tmp_root_path / uploaded_docx.name
                input_docx_path.write_bytes(uploaded_docx.getbuffer())

                sections_path = None
                if uploaded_sections is not None:
                    sections_path = tmp_root_path / uploaded_sections.name
                    # Validate JSON early for clearer errors.
                    json.loads(uploaded_sections.getvalue().decode("utf-8"))
                    sections_path.write_bytes(uploaded_sections.getbuffer())

                output_dir = tmp_root_path / "output"
                created_files = convert(
                    docx_path=input_docx_path,
                    output_dir=output_dir,
                    split_mode=split_mode,
                    sections_config=sections_path,
                )

                if not created_files:
                    st.warning("Задания не найдены. Проверьте структуру таблиц в Word.")
                else:
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                        for file_path in created_files:
                            zf.write(file_path, arcname=file_path.name)

                    zip_buffer.seek(0)
                    archive_name = f"{input_docx_path.stem}_ispring.zip"

                    st.success(f"Готово. Создано файлов: {len(created_files)}")
                    st.write("Состав архива:")
                    for path in created_files:
                        st.write(f"- {path.name}")

                    st.download_button(
                        label="Скачать архив с Excel",
                        data=zip_buffer,
                        file_name=archive_name,
                        mime="application/zip",
                        use_container_width=True,
                    )
        except Exception as exc:
            st.error(f"Ошибка конвертации: {exc}")
