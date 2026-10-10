"""Combine and clean small contact CSV exports without third-party packages."""

from __future__ import annotations

import argparse
import csv
import json
import re
from pathlib import Path


ALIASES = {
    "name": {"name", "имя", "фио"},
    "email": {"email", "e-mail", "почта", "электронная почта"},
    "phone": {"phone", "телефон"},
    "company": {"company", "компания", "организация"},
}
EMAIL_PATTERN = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")
FIELDS = ["name", "email", "phone", "company", "source"]
REJECTED_FIELDS = ["source", "row", "reason", "name", "email", "phone", "company"]


def normalize_header(value: str) -> str:
    return " ".join(value.strip().lstrip("\ufeff").casefold().split())


def column_map(fieldnames: list[str]) -> dict[str, str]:
    result: dict[str, str] = {}
    for original in fieldnames:
        header = normalize_header(original)
        for canonical, aliases in ALIASES.items():
            if header in aliases and canonical not in result:
                result[canonical] = original
    if "email" not in result:
        raise ValueError("нет обязательной колонки email/почта")
    return result


def read_contacts(path: Path):
    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        sample = stream.read(4096)
        stream.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;")
        except csv.Error:
            dialect = csv.excel
        reader = csv.DictReader(stream, dialect=dialect)
        if not reader.fieldnames:
            raise ValueError(f"{path}: файл пустой или без заголовков")
        mapping = column_map(reader.fieldnames)
        for row_number, row in enumerate(reader, start=2):
            if None in row:
                raise ValueError(f"{path}:{row_number}: колонок больше, чем в заголовке")
            contact = {
                key: " ".join((row.get(mapping.get(key, "")) or "").split())
                for key in ("name", "email", "phone", "company")
            }
            contact["email"] = contact["email"].casefold()
            contact["source"] = path.name
            yield row_number, contact


def process(paths: list[Path]) -> tuple[list[dict], list[dict], dict]:
    cleaned: list[dict] = []
    rejected: list[dict] = []
    seen: set[str] = set()
    read_count = 0

    for path in paths:
        for row_number, contact in read_contacts(path):
            read_count += 1
            email = contact["email"]
            if not EMAIL_PATTERN.fullmatch(email):
                reason = "некорректный или пустой email"
            elif email in seen:
                reason = "дубликат email"
            else:
                seen.add(email)
                cleaned.append(contact)
                continue
            rejected.append({"source": path.name, "row": row_number, "reason": reason, **contact})

    summary = {"input_files": len(paths), "rows_read": read_count,
               "rows_saved": len(cleaned), "rows_rejected": len(rejected)}
    return cleaned, rejected, summary


def write_csv(path: Path, fields: list[str], rows: list[dict]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    parser = argparse.ArgumentParser(description="Объединить и очистить CSV с контактами")
    parser.add_argument("files", nargs="+", type=Path, help="исходные CSV-файлы")
    parser.add_argument("--out", type=Path, default=Path("output"), help="папка результата")
    args = parser.parse_args()
    try:
        cleaned, rejected, summary = process(args.files)
        args.out.mkdir(parents=True, exist_ok=True)
        write_csv(args.out / "cleaned.csv", FIELDS, cleaned)
        write_csv(args.out / "rejected.csv", REJECTED_FIELDS, rejected)
        (args.out / "summary.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
    except (OSError, UnicodeError, ValueError, csv.Error) as exc:
        parser.exit(1, f"Ошибка: {exc}\n")
    print(f"Готово: {summary['rows_saved']} записей, {summary['rows_rejected']} отклонено. Папка: {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
