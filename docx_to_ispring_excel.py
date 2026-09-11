from __future__ import annotations

import argparse
from collections import defaultdict
import json
from pathlib import Path
import re
from typing import Dict, Iterable, List, Optional, Tuple

from docx import Document
from docx.document import Document as DocxDocument
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl import Workbook

HEADER = [
    "Тип вопроса",
    "Текст вопроса",
    "Рисунок",
    "Видео",
    "Аудио",
    "Ответ 1",
    "Ответ 2",
    "Ответ 3",
    "Ответ 4",
    "Ответ 5",
    "Ответ 6",
    "Ответ 7",
    "Ответ 8",
    "Ответ 9",
    "Ответ 10",
    "Сообщение, если верно",
    "Сообщение, если неверно",
    "Баллы",
]

SUCCESS_MSG = "Верно"
FAIL_MSG = "Неверно"
DEFAULT_POINTS = {
    "TF": 1,
    "MC": 1,
    "TI": 1,
}


def slugify(text: str) -> str:
    value = text.strip().replace(" ", "_")
    value = re.sub(r"[^0-9A-Za-zА-Яа-яЁё_.-]+", "_", value)
    value = re.sub(r"_+", "_", value).strip("_")
    return value or "section"


def extract_task_number(task_title: str) -> Optional[int]:
    match = re.search(r"^task\s+(\d+)\b", task_title, flags=re.IGNORECASE)
    if match:
        return int(match.group(1))
    return None


def load_sections_config(config_path: Path) -> List[Tuple[str, set[int]]]:
    data = json.loads(config_path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("sections_config должен быть JSON-объектом вида {\"Раздел\": [1,2,...]}.")

    sections: List[Tuple[str, set[int]]] = []
    for section_name, task_list in data.items():
        if not isinstance(section_name, str) or not isinstance(task_list, list):
            raise ValueError("Неверный формат sections_config: ожидаются строковые имена и массивы номеров Task.")
        task_numbers = set()
        for value in task_list:
            if not isinstance(value, int):
                raise ValueError("Номера Task в sections_config должны быть целыми числами.")
            task_numbers.add(value)
        sections.append((section_name, task_numbers))
    return sections


def iter_blocks(document: DocxDocument) -> Iterable[Paragraph | Table]:
    """Yield paragraphs and tables in their original order."""
    for child in document.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, document)
        elif isinstance(child, CT_Tbl):
            yield Table(child, document)


def clean_text(value: str) -> str:
    return " ".join(value.replace("\n", " ").split())


def classify_and_map_row(cells: List[str]) -> Tuple[str, List[str]]:
    """Map one table row from DOCX to iSpring import row."""
    question_text = clean_text(cells[1]) if len(cells) > 1 else ""
    answers = [clean_text(c) for c in cells[2:] if clean_text(c)]

    # TF table rows contain question + Right/Wrong options.
    if len(answers) == 2 and {
        a.lstrip("*").lower() for a in answers
    } == {"right", "wrong"}:
        q_type = "TF"
    # TI table rows usually contain one expected answer.
    elif len(answers) == 1:
        q_type = "TI"
        answers = [answers[0].lstrip("*").strip()]
    # Default to MC for rows with multiple options.
    else:
        q_type = "MC"

    row = [q_type, question_text, "", "", ""]

    if len(answers) > 10:
        answers = answers[:10]

    row.extend(answers)
    row.extend([""] * (10 - len(answers)))
    row.extend([SUCCESS_MSG, FAIL_MSG, DEFAULT_POINTS.get(q_type, 1)])

    return q_type, row


def map_row_with_forced_type(cells: List[str], forced_type: str) -> List[str]:
    """Map row while forcing the question type for the whole task/table."""
    question_text = clean_text(cells[1]) if len(cells) > 1 else ""
    answers = [clean_text(c) for c in cells[2:] if clean_text(c)]

    if forced_type == "TI" and answers:
        answers = [answers[0].lstrip("*").strip()]

    if len(answers) > 10:
        answers = answers[:10]

    row = [forced_type, question_text, "", "", ""]
    row.extend(answers)
    row.extend([""] * (10 - len(answers)))
    row.extend([SUCCESS_MSG, FAIL_MSG, DEFAULT_POINTS.get(forced_type, 1)])
    return row


def infer_table_type(cells: List[str]) -> str:
    """Infer one question type for the entire table from a representative row."""
    answers = [clean_text(c) for c in cells[2:] if clean_text(c)]

    if len(answers) == 2 and {a.lstrip("*").lower() for a in answers} == {"right", "wrong"}:
        return "TF"
    if len(answers) <= 1:
        return "TI"
    return "MC"


def build_group_key(
    q_type: str,
    task_num: Optional[int],
    split_mode: str,
    sections: List[Tuple[str, set[int]]],
) -> str:
    if split_mode == "type":
        return q_type

    if split_mode == "section-type":
        section_name = "БезРаздела"
        if task_num is not None:
            for candidate_name, numbers in sections:
                if task_num in numbers:
                    section_name = candidate_name
                    break
        return f"{slugify(section_name)}_{q_type}"

    if split_mode == "task":
        task_label = f"Task{task_num:02d}" if task_num is not None else "Task00"
        return f"{task_label}_{q_type}"

    raise ValueError(f"Неизвестный split_mode: {split_mode}")


def parse_docx_to_rows(
    docx_path: Path,
    split_mode: str,
    sections: List[Tuple[str, set[int]]],
) -> Dict[str, List[List[str]]]:
    doc = Document(str(docx_path))
    rows_by_group: Dict[str, List[List[str]]] = defaultdict(list)
    current_task_title = ""
    auto_task_counter = 0

    for block in iter_blocks(doc):
        if isinstance(block, Paragraph):
            text = clean_text(block.text)
            lower = text.lower()
            if lower.startswith("task ") or lower.startswith("задание "):
                current_task_title = text
            continue

        # block is a table
        table_data_rows: List[List[str]] = []

        for table_row in block.rows:
            cells = [c.text.strip() for c in table_row.cells]
            if len(cells) < 3:
                continue
            if not clean_text(cells[1]):
                continue

            table_data_rows.append(cells)

        if not table_data_rows:
            continue

        explicit_task_num = extract_task_number(current_task_title)
        if explicit_task_num is None:
            auto_task_counter += 1
            task_num = auto_task_counter
        else:
            task_num = explicit_task_num

        if split_mode == "task":
            table_type = infer_table_type(table_data_rows[0])
            group_key = build_group_key(table_type, task_num, split_mode, sections)
            for cells in table_data_rows:
                mapped = map_row_with_forced_type(cells, table_type)
                rows_by_group[group_key].append(mapped)
            continue

        for cells in table_data_rows:
            q_type, mapped = classify_and_map_row(cells)
            group_key = build_group_key(q_type, task_num, split_mode, sections)
            rows_by_group[group_key].append(mapped)

    return rows_by_group


def write_type_file(
    output_dir: Path,
    base_name: str,
    group_key: str,
    rows: List[List[str]],
) -> Path:
    output_path = output_dir / f"{base_name}_{group_key}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Импорт"
    ws.append(HEADER)

    for row in rows:
        ws.append(row)

    wb.save(output_path)
    return output_path


def convert(
    docx_path: Path,
    output_dir: Path,
    split_mode: str,
    sections_config: Optional[Path],
) -> List[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    sections: List[Tuple[str, set[int]]] = []
    if split_mode == "section-type":
        if not sections_config:
            raise ValueError("Для split_mode=section-type нужен --sections-config <json>")
        sections = load_sections_config(sections_config)

    rows_by_group = parse_docx_to_rows(docx_path, split_mode, sections)
    base_name = docx_path.stem

    created_files: List[Path] = []
    for group_key in sorted(rows_by_group.keys()):
        created_files.append(
            write_type_file(output_dir, base_name, group_key, rows_by_group[group_key])
        )
    return created_files


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Конвертирует тесты из DOCX в отдельные Excel-файлы для импорта в iSpring "
            "(каждый тип вопросов в отдельный файл)."
        )
    )
    parser.add_argument("docx", type=Path, help="Путь к входному .docx")
    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        default=Path("output_ispring"),
        help="Папка для выходных .xlsx файлов",
    )
    parser.add_argument(
        "--split-mode",
        choices=["type", "section-type", "task"],
        default="task",
        help=(
            "Режим разбиения: task (строго 1 Task = 1 файл), "
            "type (по типу), section-type (по разделу и типу)"
        ),
    )
    parser.add_argument(
        "--sections-config",
        type=Path,
        default=None,
        help="JSON-файл вида {\"Раздел\": [номера Task]} для режима section-type",
    )

    args = parser.parse_args()
    created = convert(args.docx, args.output_dir, args.split_mode, args.sections_config)

    if not created:
        print("Входной файл прочитан, но задания не найдены.")
        return

    print("Созданы файлы:")
    for file_path in created:
        print(f"- {file_path}")


if __name__ == "__main__":
    main()
