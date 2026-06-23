import argparse
import re
import sys
import warnings
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    pd = None

warnings.filterwarnings(
    "ignore",
    message="Data Validation extension is not supported.*",
    category=UserWarning,
    module="openpyxl.*",
)


OUTPUT_COLUMNS = [
    "technical_code",
    "channel",
    "origin",
    "id_language",
    "technical_exception",
    "user_message",
    "message_description",
    "type",
    "message_type",
    "transaction_code",
    "itc_homologated_code",
    "response_type",
    "status",
    "reason",
]
EXTENDED_OUTPUT_COLUMNS = [
    "technical_code",
    "channel",
    "origin",
    "id_language",
    "technical_exception",
    "user_message",
    "message_description",
    "type",
    "message_type",
    "transaction_code",
    "itc_homologated_code",
    "response_type",
    "stage_type",
    "status",
    "reason",
]

INPUT_TO_OUTPUT = {
    "technicalCode": "technical_code",
    "channel": "channel",
    "origin": "origin",
    "idLanguage": "id_language",
    "technicalException": "technical_exception",
    "userMessage": "user_message",
    "messageDescription": "message_description",
    "type": "type",
    "messageType": "message_type",
    "transactionCode": "transaction_code",
    "itcHomologatedCode": "itc_homologated_code",
    "responseType": "response_type",
}


def normalize_column_name(name):
    return re.sub(r"[^a-z0-9]", "", str(name).strip().lower())


def clean_cell(value):
    if value is None or pd.isna(value):
        return ""

    text = str(value)
    return re.sub(r"[\r\n]+", " ", text)


def normalize_itc_homologated_code(value):
    text = value.strip()
    if not text:
        return ""

    if re.fullmatch(r"\d+(\.0+)?", text):
        number_text = text.split(".", 1)[0]
        return number_text.zfill(3)

    return value


def find_itc_fallback_column(df, normalized_columns):
    if normalize_column_name("itcHomologatedCode") in normalized_columns:
        return None

    column_names = list(df.columns)
    normalized_names = [normalize_column_name(column) for column in column_names]

    try:
        transaction_index = normalized_names.index(normalize_column_name("transactionCode"))
        response_index = normalized_names.index(normalize_column_name("responseType"))
    except ValueError:
        return None

    if response_index <= transaction_index:
        return None

    for candidate_index in range(transaction_index + 1, response_index):
        candidate = column_names[candidate_index]
        if non_empty_column_count(df[candidate]) > 0:
            return candidate

    return None


def non_empty_column_count(series):
    return int(series.astype(str).str.strip().ne("").sum())


def format_csv_value(value):
    text = str(value)
    text = re.sub(r'(?<!\\)"', r'\\"', text)
    return f'"{text}"'


def write_csv(output_path, rows, columns):
    with output_path.open("w", encoding="utf-8-sig", newline="") as csv_file:
        csv_file.write(",".join(format_csv_value(column) for column in columns))
        csv_file.write("\n")

        for row in rows:
            csv_file.write(",".join(format_csv_value(row[column]) for column in columns))
            csv_file.write("\n")


def expected_column_count(columns):
    normalized_columns = {normalize_column_name(column) for column in columns}
    expected_columns = {normalize_column_name(column) for column in INPUT_TO_OUTPUT}
    return len(normalized_columns & expected_columns)


def non_empty_row_count(df):
    if df.empty:
        return 0
    return int((df.astype(str).apply(lambda row: "".join(row).strip(), axis=1) != "").sum())


def read_best_sheet(input_path, sheet_name=None):
    if sheet_name:
        return pd.read_excel(
            input_path,
            sheet_name=sheet_name,
            dtype=str,
            keep_default_na=False,
        )

    try:
        workbook = pd.ExcelFile(input_path)
    except Exception as exc:
        raise RuntimeError(f"No se pudo abrir el Excel '{input_path}': {exc}") from exc

    sheet_candidates = []

    for sheet_index, current_sheet in enumerate(workbook.sheet_names):
        try:
            current_df = pd.read_excel(
                workbook,
                sheet_name=current_sheet,
                dtype=str,
                keep_default_na=False,
            )
        except Exception:
            continue

        sheet_candidates.append(
            {
                "df": current_df,
                "expected_columns": expected_column_count(current_df.columns),
                "rows": non_empty_row_count(current_df),
                "sheet_index": sheet_index,
            }
        )

    if not sheet_candidates:
        raise RuntimeError("No se pudo leer ninguna hoja del Excel.")

    max_expected_columns = max(candidate["expected_columns"] for candidate in sheet_candidates)
    if max_expected_columns == 0:
        raise RuntimeError("No se encontraron hojas con columnas esperadas en el Excel.")

    minimum_expected_columns = min(6, max_expected_columns)
    valid_candidates = [
        candidate
        for candidate in sheet_candidates
        if candidate["expected_columns"] >= minimum_expected_columns
    ]

    best_candidate = max(
        valid_candidates,
        key=lambda candidate: (
            candidate["rows"],
            candidate["expected_columns"],
            -candidate["sheet_index"],
        ),
    )

    best_df = best_candidate["df"]
    return best_df


def convert_excel_to_csv(input_path, output_path=None, sheet_name=None, extended=False):
    if pd is None:
        raise RuntimeError("No se encontro pandas. Instalalo con: pip install pandas openpyxl")

    input_path = Path(input_path)
    if output_path is None:
        suffix = "_extendido.csv" if extended else ".csv"
        output_path = input_path.with_name(f"{input_path.stem}{suffix}")
    else:
        output_path = Path(output_path)

    if not input_path.exists():
        raise FileNotFoundError(f"No existe el archivo de entrada: {input_path}")

    try:
        df = read_best_sheet(input_path, sheet_name=sheet_name)
    except Exception as exc:
        raise RuntimeError(f"No se pudo leer el Excel '{input_path}': {exc}") from exc

    normalized_columns = {
        normalize_column_name(column): column
        for column in df.columns
    }
    itc_fallback_column = find_itc_fallback_column(df, normalized_columns)

    output_columns = EXTENDED_OUTPUT_COLUMNS if extended else OUTPUT_COLUMNS
    output_rows = []
    for _, row in df.iterrows():
        output_row = {column: "" for column in output_columns}
        if extended:
            output_row["stage_type"] = "PILOTO"

        for input_column, output_column in INPUT_TO_OUTPUT.items():
            source_column = normalized_columns.get(normalize_column_name(input_column))
            if source_column is not None:
                value = clean_cell(row[source_column])
                if output_column == "itc_homologated_code":
                    value = normalize_itc_homologated_code(value)
                output_row[output_column] = value

        if not output_row["itc_homologated_code"] and itc_fallback_column is not None:
            output_row["itc_homologated_code"] = normalize_itc_homologated_code(
                clean_cell(row[itc_fallback_column])
            )

        if extended and output_row["response_type"].strip().upper() == "ALERT":
            output_row["response_type"] = "WARNING"

        output_rows.append(output_row)

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        write_csv(output_path, output_rows, output_columns)
    except Exception as exc:
        raise RuntimeError(f"No se pudo escribir el CSV '{output_path}': {exc}") from exc

    return output_path, len(output_rows)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Convierte un catalogo de errores bancarios desde Excel a CSV."
    )
    parser.add_argument("entrada", help="Ruta del archivo Excel de entrada (.xlsx).")
    parser.add_argument(
        "salida",
        nargs="?",
        help="Ruta del CSV de salida. Si se omite, usa el mismo nombre con extension .csv.",
    )
    parser.add_argument(
        "--hoja",
        help="Nombre de la hoja a convertir. Si se omite, se detecta automaticamente.",
    )
    parser.add_argument(
        "--extendido",
        action="store_true",
        help="Genera la version extendida con la columna stage_type=PILOTO.",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    try:
        output_path, row_count = convert_excel_to_csv(
            args.entrada,
            args.salida,
            sheet_name=args.hoja,
            extended=args.extendido,
        )
    except (FileNotFoundError, RuntimeError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)

    print(f"Filas generadas: {row_count}")
    print(f"CSV generado: {output_path}")


if __name__ == "__main__":
    main()
