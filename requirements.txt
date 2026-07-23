from __future__ import annotations

import argparse
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd


REQUIRED_COLUMNS = {
    "Артикул поставщика",
    "doc_type_name",
    "quantity",
    "supplier_oper_name",
    "order_dt",
    "sale_dt",
    "rr_dt",
    "delivery_rub",
    "penalty",
    "bonus_type_name",
    "techSize",
}

RUSSIAN_MONTHS = {
    "января": "01",
    "февраля": "02",
    "марта": "03",
    "апреля": "04",
    "мая": "05",
    "июня": "06",
    "июля": "07",
    "августа": "08",
    "сентября": "09",
    "октября": "10",
    "ноября": "11",
    "декабря": "12",
}


def parse_mixed_date_value(value: object) -> pd.Timestamp:
    """Преобразует дату из Excel, datetime или русской текстовой записи."""
    if pd.isna(value):
        return pd.NaT

    if isinstance(value, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(value)

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return pd.Timestamp("1899-12-30") + pd.to_timedelta(float(value), unit="D")

    text = str(value).strip().lower()
    text = re.sub(r"\s*г\.?$", "", text)

    for month_name, month_number in RUSSIAN_MONTHS.items():
        text = re.sub(rf"\b{month_name}\b", month_number, text)

    return pd.to_datetime(text, dayfirst=True, errors="coerce")


def parse_date_column(series: pd.Series) -> pd.Series:
    return series.map(parse_mixed_date_value)


def normalized_text(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip()


def normalize_size(value: object) -> str:
    if pd.isna(value):
        return ""

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = float(value)
        return str(int(number)) if number.is_integer() else str(number)

    text = str(value).strip()
    try:
        number = float(text.replace(",", "."))
        return str(int(number)) if number.is_integer() else str(number)
    except ValueError:
        return text


def clean_number(value: float) -> int | float:
    value = float(value)
    return int(value) if value.is_integer() else round(value, 2)


def load_report(input_path: Path) -> pd.DataFrame:
    if not input_path.exists():
        raise FileNotFoundError(f"Файл не найден: {input_path}")

    df = pd.read_excel(input_path)

    missing = sorted(REQUIRED_COLUMNS.difference(df.columns))
    if missing:
        raise ValueError(
            "В отчёте отсутствуют обязательные столбцы: " + ", ".join(missing)
        )

    for column in ("quantity", "delivery_rub", "penalty"):
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)

    df["order_date"] = parse_date_column(df["order_dt"])
    df["sale_date"] = parse_date_column(df["sale_dt"])
    df["report_date"] = parse_date_column(df["rr_dt"])

    df["operation"] = normalized_text(df["supplier_oper_name"])
    df["document_type"] = normalized_text(df["doc_type_name"])
    df["supplier_article"] = normalized_text(df["Артикул поставщика"])
    df["penalty_article"] = normalized_text(df["bonus_type_name"]).replace(
        "", "Статья не указана"
    )
    df["size"] = df["techSize"].map(normalize_size)

    return df


def analyze_report(
    df: pd.DataFrame,
    year: int,
    target_article: str,
) -> tuple[pd.DataFrame, dict[str, pd.DataFrame]]:
    july_sales = df[
        (df["order_date"].dt.year == year)
        & (df["order_date"].dt.month == 7)
        & (df["document_type"] == "Продажа")
    ]
    july_sales_quantity = clean_number(july_sales["quantity"].sum())

    october_logistics = df[
        (df["report_date"].dt.year == year)
        & (df["report_date"].dt.month == 10)
        & (df["operation"] == "Логистика")
    ]
    october_logistics_sum = clean_number(october_logistics["delivery_rub"].sum())

    september_penalties = df[
        (df["report_date"].dt.year == year)
        & (df["report_date"].dt.month == 9)
        & (df["penalty"] > 0)
    ]
    penalty_totals = (
        september_penalties.groupby("penalty_article", as_index=False)["penalty"]
        .sum()
        .sort_values("penalty", ascending=True)
    )

    if penalty_totals.empty:
        min_penalty_article = "Нет начислений"
        min_penalty_sum = 0
    else:
        min_penalty_article = str(penalty_totals.iloc[0]["penalty_article"])
        min_penalty_sum = clean_number(penalty_totals.iloc[0]["penalty"])

    august_article_sales = df[
        (df["sale_date"].dt.year == year)
        & (df["sale_date"].dt.month == 8)
        & (df["operation"] == "Продажа")
        & (df["supplier_article"].str.casefold() == target_article.casefold())
    ]
    size_sales = (
        august_article_sales.groupby("size", as_index=False)["quantity"]
        .sum()
        .sort_values("quantity", ascending=False)
    )

    if size_sales.empty:
        leading_size = "Нет продаж"
        leading_size_sales = 0
    else:
        leading_size = str(size_sales.iloc[0]["size"])
        leading_size_sales = clean_number(size_sales.iloc[0]["quantity"])

    november_returns = df[
        (df["sale_date"].dt.year == year)
        & (df["sale_date"].dt.month == 11)
        & (df["operation"] == "Возврат")
    ]
    return_totals = (
        november_returns.groupby("supplier_article", as_index=False)["quantity"]
        .sum()
        .sort_values("quantity", ascending=False)
    )

    if return_totals.empty:
        most_returned_article = "Нет возвратов"
        most_returned_quantity = 0
    else:
        most_returned_article = str(return_totals.iloc[0]["supplier_article"])
        most_returned_quantity = clean_number(return_totals.iloc[0]["quantity"])

    summary = pd.DataFrame(
        [
            {
                "№": 1,
                "Показатель": "Продажи товаров, заказанных в июле",
                "Значение": july_sales_quantity,
                "Единица": "шт.",
                "Комментарий": (
                    'SUM(quantity), где месяц order_dt = 7 '
                    'и doc_type_name = "Продажа"'
                ),
            },
            {
                "№": 2,
                "Показатель": "Сумма логистики за октябрь",
                "Значение": october_logistics_sum,
                "Единица": "руб.",
                "Комментарий": (
                    'SUM(delivery_rub), где месяц rr_dt = 10 '
                    'и supplier_oper_name = "Логистика"'
                ),
            },
            {
                "№": 3,
                "Показатель": "Минимальная сумма штрафа за сентябрь",
                "Значение": min_penalty_sum,
                "Единица": "руб.",
                "Комментарий": min_penalty_article,
            },
            {
                "№": 4,
                "Показатель": f"Лидер продаж по размеру артикула {target_article} в августе",
                "Значение": leading_size_sales,
                "Единица": "шт.",
                "Комментарий": f"Размер: {leading_size}",
            },
            {
                "№": 5,
                "Показатель": "Самый возвращаемый артикул в ноябре",
                "Значение": most_returned_quantity,
                "Единица": "шт.",
                "Комментарий": f"Артикул: {most_returned_article}",
            },
        ]
    )

    details = {
        "Штрафы сентябрь": penalty_totals.rename(
            columns={"penalty_article": "Статья штрафа", "penalty": "Сумма штрафа"}
        ),
        f"Продажи {target_article} август": size_sales.rename(
            columns={"size": "Размер", "quantity": "Количество продаж"}
        ),
        "Возвраты ноябрь": return_totals.rename(
            columns={
                "supplier_article": "Артикул поставщика",
                "quantity": "Количество возвратов",
            }
        ),
    }

    return summary, details


def save_result(
    output_path: Path,
    summary: pd.DataFrame,
    details: dict[str, pd.DataFrame],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Итоги", index=False)

        for sheet_name, table in details.items():
            table.to_excel(writer, sheet_name=sheet_name[:31], index=False)

        for worksheet in writer.book.worksheets:
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions

            for column_cells in worksheet.columns:
                max_length = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in column_cells
                )
                width = min(max(max_length + 2, 12), 55)
                worksheet.column_dimensions[column_cells[0].column_letter].width = width


def print_summary(summary: pd.DataFrame) -> None:
    print("\nРезультаты анализа:")
    for _, row in summary.iterrows():
        print(
            f"{row['№']}. {row['Показатель']}: "
            f"{row['Значение']} {row['Единица']} — {row['Комментарий']}"
        )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Анализ отчёта Wildberries по реализации."
    )
    parser.add_argument("input_file", type=Path, help="Путь к исходному XLSX-файлу")
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("output/analysis_result.xlsx"),
        help="Путь к итоговому XLSX-файлу",
    )
    parser.add_argument(
        "--year",
        type=int,
        default=2023,
        help="Год анализа. По умолчанию: 2023",
    )
    parser.add_argument(
        "--article",
        default="ack",
        help="Артикул для поиска лидирующего размера. По умолчанию: ack",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()

    try:
        report = load_report(args.input_file)
        summary, details = analyze_report(report, args.year, args.article)
        save_result(args.output, summary, details)
        print_summary(summary)
        print(f"\nИтоговый файл создан: {args.output.resolve()}")
    except (FileNotFoundError, ValueError, PermissionError) as error:
        raise SystemExit(f"Ошибка: {error}") from error


if __name__ == "__main__":
    main()
