"""Conversor de estados de cuenta bancarios mexicanos de PDF a Excel.

Ejecucion local:
    streamlit run app.py

Dependencias:
    streamlit pdfplumber pandas openpyxl
"""

from __future__ import annotations

import io
import re
import smtplib
import ssl
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from email.message import EmailMessage
from typing import Callable, Sequence

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


BANKS = [
    "BBVA",
    "Banorte",
    "Santander",
    "Citibanamex",
    "Scotiabank",
    "HSBC",
    "Banca Afirme",
    "Banregio",
    "Banco Inbursa",
    "Banco Azteca",
    "Banco del Bajío",
    "BanCoppel",
    "Banco Bancrea",
    "Banco Mifel",
    "Banco Actinver",
]
COLUMNS = ["Fecha", "Concepto / Descripción", "Depósito", "Retiro", "Saldo"]
EMAIL_RE = re.compile(r"^[A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Z0-9-]+(?:\.[A-Z0-9-]+)+$", re.I)

DATE_TOKEN = r"(?:\d{1,2}[/-]\d{1,2}(?:[/-]\d{2,4})?|\d{1,2}\s+(?:ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)[A-Z]*\.?(?:\s+\d{2,4})?)"
MONEY_TOKEN = r"(?:\$\s*)?-?\s*\d[\d,]*(?:\.\d{2})"
DATE_RE = re.compile(rf"^\s*(?P<date>{DATE_TOKEN})\b", re.I)
MONEY_RE = re.compile(rf"(?<!\w)({MONEY_TOKEN})(?!\w)", re.I)

MONTHS = {
    "ENE": 1, "ENERO": 1, "FEB": 2, "FEBRERO": 2, "MAR": 3, "MARZO": 3,
    "ABR": 4, "ABRIL": 4, "MAY": 5, "MAYO": 5, "JUN": 6, "JUNIO": 6,
    "JUL": 7, "JULIO": 7, "AGO": 8, "AGOSTO": 8, "SEP": 9,
    "SEPT": 9, "SEPTIEMBRE": 9, "OCT": 10, "OCTUBRE": 10,
    "NOV": 11, "NOVIEMBRE": 11, "DIC": 12, "DICIEMBRE": 12,
}


@dataclass(frozen=True)
class BankConfig:
    """Pistas de formato para interpretar los movimientos de cada banco."""

    headers: tuple[str, ...]
    ignore: tuple[str, ...]
    amount_order: tuple[str, ...]
    deposit_words: tuple[str, ...]
    withdrawal_words: tuple[str, ...]


COMMON_IGNORE = (
    "saldo anterior", "saldo inicial", "total depositos", "total depósitos",
    "total retiros", "resumen", "pagina ", "página ", "fecha concepto",
    "fecha descripcion", "fecha descripción", "estado de cuenta", "periodo",
)

CONFIGS = {
    "BBVA": BankConfig(
        ("FECHA", "OPERACIÓN", "LIQUIDACIÓN", "CONCEPTO", "CARGO", "ABONO", "SALDO"),
        COMMON_IGNORE, ("Retiro", "Depósito", "Saldo"),
        ("abono", "deposito", "depósito", "spei recibido", "traspaso recibido"),
        ("cargo", "retiro", "compra", "comision", "comisión", "pago", "spei enviado", "transferencia enviada"),
    ),
    "Banorte": BankConfig(
        ("FECHA", "DESCRIPCIÓN", "DEPÓSITOS", "RETIROS", "SALDO"),
        COMMON_IGNORE, ("Depósito", "Retiro", "Saldo"),
        ("deposito", "depósito", "abono", "transferencia recibida", "spei recibido"),
        ("retiro", "cargo", "compra", "comision", "comisión", "pago", "spei enviado", "transferencia enviada"),
    ),
    "Santander": BankConfig(
        ("FECHA", "CONCEPTO", "RETIRO", "DEPÓSITO", "SALDO"),
        COMMON_IGNORE, ("Retiro", "Depósito", "Saldo"),
        ("deposito", "depósito", "abono", "transferencia recibida"),
        ("retiro", "cargo", "compra", "comision", "comisión", "pago", "spei enviado", "transferencia enviada"),
    ),
    "Citibanamex": BankConfig(
        ("FECHA", "CONCEPTO", "RETIROS", "DEPÓSITOS", "SALDO"),
        COMMON_IGNORE, ("Retiro", "Depósito", "Saldo"),
        ("deposito", "depósito", "abono", "transferencia recibida"),
        ("retiro", "cargo", "compra", "comision", "comisión", "pago", "spei enviado", "transferencia enviada"),
    ),
    "Scotiabank": BankConfig(
        ("FECHA", "CONCEPTO", "CARGOS", "ABONOS", "SALDO"),
        COMMON_IGNORE, ("Retiro", "Depósito", "Saldo"),
        ("abono", "deposito", "depósito", "transferencia recibida"),
        ("cargo", "retiro", "compra", "comision", "comisión", "pago", "spei enviado", "transferencia enviada"),
    ),
    "HSBC": BankConfig(
        ("FECHA", "DESCRIPCIÓN", "DEPÓSITO/ABONO", "RETIRO/CARGO", "SALDO"),
        COMMON_IGNORE, ("Depósito", "Retiro", "Saldo"),
        ("abono", "deposito", "depósito", "transferencia recibida"),
        ("cargo", "retiro", "compra", "comision", "comisión", "pago", "spei enviado", "transferencia enviada"),
    ),
    "Banca Afirme": BankConfig(
        ("FECHA", "REFERENCIA", "CONCEPTO", "CARGOS", "ABONOS", "SALDO"),
        COMMON_IGNORE, ("Retiro", "Depósito", "Saldo"),
        ("abono", "deposito", "depósito", "credito", "crédito", "spei recibido", "traspaso recibido"),
        ("cargo", "retiro", "compra", "comision", "comisión", "pago", "spei enviado", "traspaso enviado"),
    ),
    "Banregio": BankConfig(
        ("FECHA", "DESCRIPCIÓN", "DEPÓSITOS", "RETIROS", "SALDO"),
        COMMON_IGNORE, ("Depósito", "Retiro", "Saldo"),
        ("deposito", "depósito", "abono", "credito", "crédito", "spei recibido", "transferencia recibida"),
        ("retiro", "cargo", "debito", "débito", "compra", "comision", "comisión", "pago", "spei enviado"),
    ),
    "Banco Inbursa": BankConfig(
        ("FECHA", "CONCEPTO", "DESCRIPCIÓN", "CARGO", "ABONO", "SALDO"),
        COMMON_IGNORE, ("Retiro", "Depósito", "Saldo"),
        ("abono", "deposito", "depósito", "credito", "crédito", "spei recibido", "transferencia recibida"),
        ("cargo", "retiro", "debito", "débito", "compra", "comision", "comisión", "pago", "spei enviado"),
    ),
    "Banco Azteca": BankConfig(
        ("FECHA", "MOVIMIENTO", "DESCRIPCIÓN", "CARGO", "ABONO", "SALDO"),
        COMMON_IGNORE, ("Retiro", "Depósito", "Saldo"),
        ("abono", "deposito", "depósito", "credito", "crédito", "spei recibido", "pago recibido"),
        ("cargo", "retiro", "debito", "débito", "compra", "comision", "comisión", "pago", "spei enviado"),
    ),
    "Banco del Bajío": BankConfig(
        ("FECHA", "REFERENCIA", "DESCRIPCIÓN", "CARGOS", "ABONOS", "SALDO"),
        COMMON_IGNORE, ("Retiro", "Depósito", "Saldo"),
        ("abono", "deposito", "depósito", "credito", "crédito", "spei recibido", "transferencia recibida"),
        ("cargo", "retiro", "debito", "débito", "compra", "comision", "comisión", "pago", "spei enviado"),
    ),
    "BanCoppel": BankConfig(
        ("FECHA", "DESCRIPCIÓN", "CONCEPTO", "CARGO", "ABONO", "SALDO"),
        COMMON_IGNORE, ("Retiro", "Depósito", "Saldo"),
        ("abono", "deposito", "depósito", "credito", "crédito", "spei recibido", "transferencia recibida"),
        ("cargo", "retiro", "debito", "débito", "compra", "comision", "comisión", "pago", "spei enviado"),
    ),
    "Banco Bancrea": BankConfig(
        ("FECHA", "OPERACIÓN", "CONCEPTO", "CARGO", "ABONO", "SALDO"),
        COMMON_IGNORE, ("Retiro", "Depósito", "Saldo"),
        ("abono", "deposito", "depósito", "credito", "crédito", "spei recibido", "transferencia recibida"),
        ("cargo", "retiro", "debito", "débito", "compra", "comision", "comisión", "pago", "spei enviado"),
    ),
    "Banco Mifel": BankConfig(
        ("FECHA", "REFERENCIA", "DESCRIPCIÓN", "CARGOS", "ABONOS", "SALDO"),
        COMMON_IGNORE, ("Retiro", "Depósito", "Saldo"),
        ("abono", "deposito", "depósito", "credito", "crédito", "spei recibido", "transferencia recibida"),
        ("cargo", "retiro", "debito", "débito", "compra", "comision", "comisión", "pago", "spei enviado"),
    ),
    "Banco Actinver": BankConfig(
        ("FECHA", "MOVIMIENTO", "CONCEPTO", "RETIROS", "DEPÓSITOS", "SALDO"),
        COMMON_IGNORE, ("Retiro", "Depósito", "Saldo"),
        ("abono", "deposito", "depósito", "credito", "crédito", "spei recibido", "venta", "liquidación a favor"),
        ("cargo", "retiro", "debito", "débito", "compra", "comision", "comisión", "pago", "spei enviado", "compra de títulos"),
    ),
}


def clean_text(value: object) -> str:
    """Normaliza espacios y saltos de linea sin eliminar acentos."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\u00a0", " ")).strip()


def comparable(value: str) -> str:
    value = unicodedata.normalize("NFKD", clean_text(value)).encode("ascii", "ignore").decode()
    return value.casefold()


def money_to_float(value: object) -> float:
    """Convierte $ 1,234.56, (1,234.56) o -1,234.56 a float."""
    if value is None or clean_text(value) in {"", "-", "—"}:
        return 0.0
    raw = clean_text(value)
    negative = raw.startswith("(") and raw.endswith(")")
    raw = re.sub(r"[^\d.\-]", "", raw.replace(",", ""))
    try:
        amount = float(raw)
    except ValueError:
        return 0.0
    return round(-abs(amount) if negative else amount, 2)


def infer_statement_year(lines: Sequence[str]) -> int:
    """Busca el anio del periodo; usa el actual solo como ultimo recurso."""
    period_re = re.compile(r"(?:PERIODO|CORTE|AL)\D{0,25}(20\d{2})", re.I)
    for line in lines:
        match = period_re.search(line)
        if match:
            return int(match.group(1))
    years = [int(y) for line in lines[:80] for y in re.findall(r"\b(20\d{2})\b", line)]
    return max(set(years), key=years.count) if years else datetime.now().year


def normalize_date(value: str, default_year: int) -> str:
    """Devuelve una fecha bancaria en formato DD/MM/YYYY."""
    text = comparable(value).upper().replace(".", "")
    numeric = re.fullmatch(r"(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?", text)
    if numeric:
        day, month = int(numeric.group(1)), int(numeric.group(2))
        year = int(numeric.group(3)) if numeric.group(3) else default_year
        year += 2000 if year < 100 else 0
        return datetime(year, month, day).strftime("%d/%m/%Y")
    named = re.fullmatch(r"(\d{1,2})\s+([A-Z]+)(?:\s+(\d{2,4}))?", text)
    if not named or named.group(2) not in MONTHS:
        raise ValueError(f"Fecha no reconocida: {value}")
    year = int(named.group(3)) if named.group(3) else default_year
    year += 2000 if year < 100 else 0
    return datetime(year, MONTHS[named.group(2)], int(named.group(1))).strftime("%d/%m/%Y")


def words_to_lines(words: Sequence[dict], tolerance: float = 3.0) -> list[str]:
    """Reconstruye lineas por coordenadas (top/x0) conservando columnas visuales."""
    if not words:
        return []
    rows: list[list[dict]] = []
    for word in sorted(words, key=lambda w: (round(float(w["top"]) / tolerance), float(w["x0"]))):
        if not rows or abs(float(word["top"]) - sum(float(w["top"]) for w in rows[-1]) / len(rows[-1])) > tolerance:
            rows.append([word])
        else:
            rows[-1].append(word)
    return [clean_text(" ".join(w["text"] for w in sorted(row, key=lambda item: float(item["x0"])))) for row in rows]


def extract_pdf_rows(pdf_bytes: bytes) -> tuple[list[str], list[list[str]]]:
    """Extrae texto posicional y tablas delineadas de todas las paginas."""
    lines: list[str] = []
    table_rows: list[list[str]] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            page_words = page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=False)
            lines.extend(words_to_lines(page_words))
            settings_candidates = (
                {"vertical_strategy": "lines", "horizontal_strategy": "lines", "snap_tolerance": 3},
                {"vertical_strategy": "text", "horizontal_strategy": "text", "intersection_tolerance": 5},
            )
            page_tables: list[list[list[str | None]]] = []
            for settings in settings_candidates:
                try:
                    page_tables = page.extract_tables(table_settings=settings) or []
                except Exception:  # Algunos PDFs tienen trazos malformados.
                    page_tables = []
                if page_tables:
                    break
            for table in page_tables:
                for row in table:
                    cleaned = [clean_text(cell) for cell in (row or [])]
                    if any(cleaned):
                        table_rows.append(cleaned)
    return [line for line in lines if line], table_rows


def is_noise(line: str, config: BankConfig) -> bool:
    normalized = comparable(line)
    if any(token in normalized for token in map(comparable, config.ignore)):
        return True
    header_hits = sum(comparable(header) in normalized for header in config.headers)
    return header_hits >= 3


def table_rows_to_lines(rows: Sequence[Sequence[str]]) -> list[str]:
    return [clean_text(" ".join(cell for cell in row if clean_text(cell))) for row in rows]


def classify_amounts(description: str, amounts: list[float], config: BankConfig) -> tuple[float, float, float | None]:
    """Mapea importes a deposito/retiro/saldo usando columnas y pistas semanticas."""
    if not amounts:
        return 0.0, 0.0, None
    values: dict[str, float] = {}
    for label, amount in zip(reversed(config.amount_order), reversed(amounts)):
        values[label] = abs(amount)
    deposit = values.get("Depósito", 0.0)
    withdrawal = values.get("Retiro", 0.0)
    balance = values.get("Saldo")

    normalized = comparable(description)
    is_deposit = any(comparable(word) in normalized for word in config.deposit_words)
    is_withdrawal = any(comparable(word) in normalized for word in config.withdrawal_words)

    # Cuando solo existe un importe, las palabras del concepto desempatan.
    if len(amounts) == 1 and balance is not None:
        if is_deposit and not is_withdrawal:
            deposit, balance = abs(amounts[0]), None
        elif is_withdrawal and not is_deposit:
            withdrawal, balance = abs(amounts[0]), None
    # Si una columna vacia desaparecio del texto, corrige el desplazamiento
    # usando la semantica de la descripcion. El ultimo importe sigue siendo saldo.
    elif len(amounts) == 2:
        if is_withdrawal and not is_deposit and deposit and not withdrawal:
            withdrawal, deposit = deposit, 0.0
        elif is_deposit and not is_withdrawal and withdrawal and not deposit:
            deposit, withdrawal = withdrawal, 0.0
    return round(deposit, 2), round(withdrawal, 2), None if balance is None else round(balance, 2)


def parse_bank(lines: Sequence[str], table_rows: Sequence[Sequence[str]], config: BankConfig) -> pd.DataFrame:
    """Parser base: detecta movimientos y une sus continuaciones multilinea."""
    candidates = list(lines)
    # Las tablas aportan una segunda ruta cuando el texto posicional no es usable.
    table_lines = table_rows_to_lines(table_rows)
    if sum(bool(DATE_RE.match(line)) for line in table_lines) > sum(bool(DATE_RE.match(line)) for line in candidates):
        candidates = table_lines

    year = infer_statement_year(list(lines) + table_lines)
    records: list[dict] = []
    current: dict | None = None

    for raw_line in candidates:
        line = clean_text(raw_line)
        match = DATE_RE.match(line)
        if match and not is_noise(line, config):
            if current:
                records.append(current)
            date_text = match.group("date")
            remainder = clean_text(line[match.end():])
            money_matches = list(MONEY_RE.finditer(remainder))
            amounts = [money_to_float(item.group(1)) for item in money_matches]
            description = remainder[: money_matches[0].start()] if money_matches else remainder
            deposit, withdrawal, balance = classify_amounts(description, amounts, config)
            try:
                normalized_date = normalize_date(date_text, year)
            except (ValueError, OverflowError):
                current = None
                continue
            current = {
                "Fecha": normalized_date,
                "Concepto / Descripción": clean_text(description),
                "Depósito": deposit,
                "Retiro": withdrawal,
                "Saldo": balance,
            }
        elif current and line and not is_noise(line, config):
            # Continuaciones sin fecha pertenecen al concepto anterior; importes sueltos
            # se ignoran para no confundir totales o referencias numericas con dinero.
            continuation = MONEY_RE.sub("", line)
            continuation = clean_text(continuation)
            if continuation and not DATE_RE.match(continuation):
                current["Concepto / Descripción"] = clean_text(
                    f"{current['Concepto / Descripción']} {continuation}"
                )
    if current:
        records.append(current)

    frame = pd.DataFrame(records, columns=COLUMNS)
    if frame.empty:
        return frame
    frame["Concepto / Descripción"] = frame["Concepto / Descripción"].map(clean_text)
    for column in ("Depósito", "Retiro"):
        frame[column] = pd.to_numeric(frame[column], errors="coerce").fillna(0.0).round(2)
    frame["Saldo"] = pd.to_numeric(frame["Saldo"], errors="coerce").round(2)
    frame = frame[frame["Concepto / Descripción"].ne("")].drop_duplicates().reset_index(drop=True)
    return frame


def parse_bbva(lines: Sequence[str], table_rows: Sequence[Sequence[str]]) -> pd.DataFrame:
    return parse_bank(lines, table_rows, CONFIGS["BBVA"])


def parse_banorte(lines: Sequence[str], table_rows: Sequence[Sequence[str]]) -> pd.DataFrame:
    return parse_bank(lines, table_rows, CONFIGS["Banorte"])


def parse_santander(lines: Sequence[str], table_rows: Sequence[Sequence[str]]) -> pd.DataFrame:
    return parse_bank(lines, table_rows, CONFIGS["Santander"])


def parse_citibanamex(lines: Sequence[str], table_rows: Sequence[Sequence[str]]) -> pd.DataFrame:
    return parse_bank(lines, table_rows, CONFIGS["Citibanamex"])


def parse_scotiabank(lines: Sequence[str], table_rows: Sequence[Sequence[str]]) -> pd.DataFrame:
    return parse_bank(lines, table_rows, CONFIGS["Scotiabank"])


def parse_hsbc(lines: Sequence[str], table_rows: Sequence[Sequence[str]]) -> pd.DataFrame:
    return parse_bank(lines, table_rows, CONFIGS["HSBC"])


def parse_afirme(lines: Sequence[str], table_rows: Sequence[Sequence[str]]) -> pd.DataFrame:
    return parse_bank(lines, table_rows, CONFIGS["Banca Afirme"])


def parse_banregio(lines: Sequence[str], table_rows: Sequence[Sequence[str]]) -> pd.DataFrame:
    return parse_bank(lines, table_rows, CONFIGS["Banregio"])


def parse_inbursa(lines: Sequence[str], table_rows: Sequence[Sequence[str]]) -> pd.DataFrame:
    return parse_bank(lines, table_rows, CONFIGS["Banco Inbursa"])


def parse_banco_azteca(lines: Sequence[str], table_rows: Sequence[Sequence[str]]) -> pd.DataFrame:
    return parse_bank(lines, table_rows, CONFIGS["Banco Azteca"])


def parse_banbajio(lines: Sequence[str], table_rows: Sequence[Sequence[str]]) -> pd.DataFrame:
    return parse_bank(lines, table_rows, CONFIGS["Banco del Bajío"])


def parse_bancoppel(lines: Sequence[str], table_rows: Sequence[Sequence[str]]) -> pd.DataFrame:
    return parse_bank(lines, table_rows, CONFIGS["BanCoppel"])


def parse_bancrea(lines: Sequence[str], table_rows: Sequence[Sequence[str]]) -> pd.DataFrame:
    return parse_bank(lines, table_rows, CONFIGS["Banco Bancrea"])


def parse_mifel(lines: Sequence[str], table_rows: Sequence[Sequence[str]]) -> pd.DataFrame:
    return parse_bank(lines, table_rows, CONFIGS["Banco Mifel"])


def parse_actinver(lines: Sequence[str], table_rows: Sequence[Sequence[str]]) -> pd.DataFrame:
    return parse_bank(lines, table_rows, CONFIGS["Banco Actinver"])


PARSERS: dict[str, Callable[[Sequence[str], Sequence[Sequence[str]]], pd.DataFrame]] = {
    "BBVA": parse_bbva,
    "Banorte": parse_banorte,
    "Santander": parse_santander,
    "Citibanamex": parse_citibanamex,
    "Scotiabank": parse_scotiabank,
    "HSBC": parse_hsbc,
    "Banca Afirme": parse_afirme,
    "Banregio": parse_banregio,
    "Banco Inbursa": parse_inbursa,
    "Banco Azteca": parse_banco_azteca,
    "Banco del Bajío": parse_banbajio,
    "BanCoppel": parse_bancoppel,
    "Banco Bancrea": parse_bancrea,
    "Banco Mifel": parse_mifel,
    "Banco Actinver": parse_actinver,
}


def dataframe_to_excel(frame: pd.DataFrame) -> bytes:
    """Genera un XLSX con filtros, encabezados, formatos y anchos adecuados."""
    output = io.BytesIO()
    export = frame.copy()
    export["Fecha"] = pd.to_datetime(export["Fecha"], format="%d/%m/%Y", errors="coerce")
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export.to_excel(writer, sheet_name="Movimientos", index=False)
        sheet = writer.book["Movimientos"]
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        header_fill = PatternFill("solid", fgColor="1F4E78")
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
        for cell in sheet["A"][1:]:
            cell.number_format = "DD/MM/YYYY"
        for col in ("C", "D", "E"):
            for cell in sheet[col][1:]:
                cell.number_format = '$#,##0.00;[Red]-$#,##0.00'
        for index, column_cells in enumerate(sheet.columns, start=1):
            values = [len(str(cell.value or "")) for cell in column_cells]
            width = min(max(max(values, default=10) + 2, 12), 60)
            sheet.column_dimensions[get_column_letter(index)].width = width
        sheet.column_dimensions["B"].width = min(max(sheet.column_dimensions["B"].width, 35), 80)
    return output.getvalue()


def secret_value(*keys: str, default: object = None) -> object:
    """Lee claves anidadas de st.secrets sin asumir que la seccion existe."""
    value: object = st.secrets
    try:
        for key in keys:
            value = value[key]  # type: ignore[index]
        return value
    except (KeyError, TypeError):
        return default


def send_excel_email(recipient: str, excel_bytes: bytes, bank: str) -> None:
    """Envia el resultado por Gmail o cualquier servidor SMTP configurado."""
    host = str(secret_value("smtp", "host", default="smtp.gmail.com"))
    port = int(secret_value("smtp", "port", default=587))
    username = str(secret_value("smtp", "username", default=""))
    password = str(secret_value("smtp", "password", default=""))
    sender = str(secret_value("smtp", "sender", default=username))
    use_ssl = bool(secret_value("smtp", "use_ssl", default=False))
    if not username or not password or not sender:
        raise RuntimeError("Faltan smtp.username, smtp.password o smtp.sender en .streamlit/secrets.toml")

    message = EmailMessage()
    message["Subject"] = f"Estado de cuenta {bank} convertido a Excel"
    message["From"] = sender
    message["To"] = recipient
    message.set_content("Adjuntamos los movimientos extraidos de tu estado de cuenta.")
    message.add_attachment(
        excel_bytes,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"movimientos_{bank.lower()}.xlsx",
    )
    context = ssl.create_default_context()
    if use_ssl:
        with smtplib.SMTP_SSL(host, port, context=context, timeout=30) as server:
            server.login(username, password)
            server.send_message(message)
    else:
        with smtplib.SMTP(host, port, timeout=30) as server:
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            server.login(username, password)
            server.send_message(message)


def merge_preview_edits(original: pd.DataFrame, edited_preview: pd.DataFrame) -> pd.DataFrame:
    result = original.copy()
    edited = edited_preview.reindex(columns=COLUMNS).copy()
    for column in ("Depósito", "Retiro", "Saldo"):
        edited[column] = pd.to_numeric(edited[column], errors="coerce")
    result.iloc[: len(edited)] = edited.to_numpy()
    return result


def inject_styles() -> None:
    """Aplica una identidad visual limpia sin dependencias de frontend."""
    st.markdown(
        """
        <style>
        :root {
            --ink: #102a43;
            --muted: #627d98;
            --brand: #0f766e;
            --brand-dark: #115e59;
            --surface: rgba(255, 255, 255, 0.94);
            --line: #d9e2ec;
        }

        .stApp {
            background:
                radial-gradient(circle at 8% 0%, rgba(45, 212, 191, .14), transparent 28rem),
                radial-gradient(circle at 96% 12%, rgba(59, 130, 246, .10), transparent 26rem),
                #f5f8fb;
            color: var(--ink);
        }

        [data-testid="stHeader"] { background: transparent; }
        [data-testid="stToolbar"] { right: 1rem; }
        [data-testid="stAppViewContainer"] > .main .block-container {
            max-width: 1120px;
            padding-top: 2rem;
            padding-bottom: 4rem;
        }

        .hero {
            position: relative;
            overflow: hidden;
            padding: 2.7rem 3rem;
            margin-bottom: 1.4rem;
            border: 1px solid rgba(255,255,255,.14);
            border-radius: 28px;
            background: linear-gradient(125deg, #102a43 0%, #123f56 55%, #0f766e 120%);
            box-shadow: 0 24px 60px rgba(16, 42, 67, .18);
            color: white;
        }

        .hero::after {
            content: "";
            position: absolute;
            width: 240px;
            height: 240px;
            right: -55px;
            top: -75px;
            border-radius: 50%;
            background: rgba(45, 212, 191, .16);
            box-shadow: 0 0 0 48px rgba(45, 212, 191, .055);
        }

        .hero-badge {
            display: inline-flex;
            align-items: center;
            gap: .45rem;
            padding: .38rem .72rem;
            border: 1px solid rgba(255,255,255,.22);
            border-radius: 999px;
            background: rgba(255,255,255,.1);
            color: #ccfbf1;
            font-size: .76rem;
            font-weight: 700;
            letter-spacing: .08em;
            text-transform: uppercase;
        }

        .hero h1 {
            max-width: 760px;
            margin: 1rem 0 .55rem;
            color: white;
            font-size: clamp(2rem, 4vw, 3.35rem);
            line-height: 1.04;
            letter-spacing: -.045em;
        }

        .hero p {
            max-width: 680px;
            margin: 0;
            color: #d9eaf2;
            font-size: 1.05rem;
            line-height: 1.65;
        }

        .section-title {
            margin: .2rem 0 -.35rem;
            color: var(--ink);
            font-size: 1.25rem;
            font-weight: 750;
            letter-spacing: -.02em;
        }

        .section-copy {
            margin-bottom: .75rem;
            color: var(--muted);
            font-size: .92rem;
        }

        [data-testid="stForm"] {
            padding: 1.7rem 1.8rem 1.35rem;
            border: 1px solid rgba(188, 204, 220, .75);
            border-radius: 22px;
            background: var(--surface);
            box-shadow: 0 14px 38px rgba(50, 73, 94, .08);
        }

        [data-testid="stFileUploaderDropzone"] {
            min-height: 120px;
            border: 1.5px dashed #9fb3c8;
            border-radius: 15px;
            background: #f8fbfd;
        }

        [data-testid="stTextInput"] input,
        [data-baseweb="select"] > div {
            border-color: #bcccdc;
            border-radius: 11px;
            background: #fbfdff;
        }

        .stButton > button,
        .stDownloadButton > button,
        [data-testid="stFormSubmitButton"] > button {
            min-height: 3rem;
            border: 0;
            border-radius: 12px;
            background: linear-gradient(135deg, var(--brand), #0d9488);
            box-shadow: 0 8px 20px rgba(15, 118, 110, .2);
            color: white;
            font-weight: 750;
            transition: transform .16s ease, box-shadow .16s ease;
        }

        .stButton > button:hover,
        .stDownloadButton > button:hover,
        [data-testid="stFormSubmitButton"] > button:hover {
            transform: translateY(-1px);
            box-shadow: 0 11px 24px rgba(15, 118, 110, .28);
            color: white;
        }

        [data-testid="stMetric"] {
            padding: 1rem 1.15rem;
            border: 1px solid #d9e2ec;
            border-radius: 16px;
            background: rgba(255,255,255,.88);
            box-shadow: 0 8px 24px rgba(50, 73, 94, .055);
        }

        [data-testid="stDataFrame"] {
            overflow: hidden;
            border: 1px solid #d9e2ec;
            border-radius: 16px;
            box-shadow: 0 8px 24px rgba(50, 73, 94, .055);
        }

        [data-testid="stAlert"] { border-radius: 13px; }
        details { border-radius: 14px !important; }

        .privacy-note {
            margin: .85rem 0 0;
            color: #829ab1;
            font-size: .78rem;
            text-align: center;
        }

        .app-footer {
            padding-top: 2rem;
            color: #829ab1;
            font-size: .78rem;
            text-align: center;
        }

        @media (max-width: 720px) {
            [data-testid="stAppViewContainer"] > .main .block-container { padding: .8rem .8rem 2rem; }
            .hero { padding: 2rem 1.35rem; border-radius: 20px; }
            .hero h1 { font-size: 2.15rem; }
            [data-testid="stForm"] { padding: 1.2rem 1rem 1rem; border-radius: 17px; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_app() -> None:
    st.set_page_config(
        page_title="Estado de cuenta a Excel",
        page_icon="↗",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    inject_styles()
    st.markdown(
        """
        <section class="hero">
            <span class="hero-badge">✦ Conversión inteligente</span>
            <h1>Tu estado de cuenta,<br>listo para trabajar.</h1>
            <p>Convierte movimientos bancarios de PDF a Excel en segundos. Revisa, corrige y descarga un archivo limpio y ordenado.</p>
        </section>
        """,
        unsafe_allow_html=True,
    )

    with st.form("conversion_form"):
        st.markdown('<div class="section-title">Comienza la conversión</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="section-copy">Completa los datos y selecciona tu estado de cuenta en PDF.</div>',
            unsafe_allow_html=True,
        )
        contact_col, bank_col = st.columns(2, gap="large")
        with contact_col:
            email = st.text_input("Correo electrónico", placeholder="nombre@ejemplo.com")
        with bank_col:
            bank = st.selectbox("Banco", BANKS)
        pdf_file = st.file_uploader("Archivo PDF", type=["pdf"], accept_multiple_files=False)
        send_automatically = st.checkbox("Enviar también el Excel por email")
        submitted = st.form_submit_button("Convertir a Excel  →", type="primary", use_container_width=True)
        st.markdown(
            '<div class="privacy-note">🔒 Tu documento se procesa únicamente para realizar la conversión.</div>',
            unsafe_allow_html=True,
        )

    if submitted:
        st.session_state.pop("conversion", None)
        if not EMAIL_RE.fullmatch(email.strip()):
            st.error("Ingresa un email válido.")
        elif pdf_file is None:
            st.error("Selecciona un archivo PDF.")
        elif pdf_file.type not in {"application/pdf", "application/x-pdf"} and not pdf_file.name.lower().endswith(".pdf"):
            st.error("El archivo debe ser un PDF.")
        else:
            try:
                pdf_bytes = pdf_file.getvalue()
                if not pdf_bytes.startswith(b"%PDF"):
                    raise ValueError("El archivo no tiene una firma PDF válida.")
                with st.spinner("Extrayendo y organizando movimientos…"):
                    lines, tables = extract_pdf_rows(pdf_bytes)
                    frame = PARSERS[bank](lines, tables)
                if frame.empty:
                    st.warning(
                        "No se detectaron movimientos. El PDF puede ser una imagen escaneada, "
                        "estar protegido o usar un formato distinto al esperado."
                    )
                    st.session_state.pop("conversion", None)
                else:
                    st.session_state["conversion"] = {
                        "frame": frame,
                        "bank": bank,
                        "email": email.strip(),
                        "auto_send": send_automatically,
                        "sent": False,
                    }
                    st.success(f"Se detectaron {len(frame):,} movimientos.")
            except Exception as exc:
                st.session_state.pop("conversion", None)
                st.error(f"No fue posible procesar el PDF: {exc}")

    conversion = st.session_state.get("conversion")
    if not conversion:
        st.markdown(
            '<div class="app-footer">Compatible con 15 bancos de México · Conversión segura a Excel</div>',
            unsafe_allow_html=True,
        )
        return

    st.markdown('<div class="section-title">Vista previa editable</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-copy">Corrige cualquier dato de las primeras 10 filas antes de generar el archivo.</div>',
        unsafe_allow_html=True,
    )
    metric_a, metric_b, metric_c = st.columns(3, gap="medium")
    with metric_a:
        st.metric("Movimientos", f"{len(conversion['frame']):,}")
    with metric_b:
        st.metric("Depósitos", f"${conversion['frame']['Depósito'].sum():,.2f}")
    with metric_c:
        st.metric("Retiros", f"${conversion['frame']['Retiro'].sum():,.2f}")
    preview = conversion["frame"].head(10).copy()
    edited_preview = st.data_editor(
        preview,
        key="preview_editor",
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        column_config={
            "Fecha": st.column_config.TextColumn("Fecha", help="DD/MM/YYYY"),
            "Concepto / Descripción": st.column_config.TextColumn("Concepto / Descripción", width="large"),
            "Depósito": st.column_config.NumberColumn("Depósito", format="$ %.2f", min_value=0.0),
            "Retiro": st.column_config.NumberColumn("Retiro", format="$ %.2f", min_value=0.0),
            "Saldo": st.column_config.NumberColumn("Saldo", format="$ %.2f"),
        },
    )
    final_frame = merge_preview_edits(conversion["frame"], edited_preview)
    try:
        excel_bytes = dataframe_to_excel(final_frame)
    except Exception as exc:
        st.error(f"Revisa las fechas editadas; no se pudo crear el Excel: {exc}")
        return

    st.download_button(
        "Descargar Excel",
        data=excel_bytes,
        file_name=f"movimientos_{conversion['bank'].lower()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    if conversion["auto_send"] and not conversion["sent"]:
        try:
            send_excel_email(conversion["email"], excel_bytes, conversion["bank"])
            conversion["sent"] = True
            st.success(f"Excel enviado a {conversion['email']}.")
        except Exception as exc:
            st.warning(f"El Excel está listo, pero no se pudo enviar por email: {exc}")

    with st.expander("Configuración del envío por email"):
        st.code(
            '[smtp]\n'
            'host = "smtp.gmail.com"\n'
            'port = 587\n'
            'username = "tu_cuenta@gmail.com"\n'
            'password = "tu_contraseña_de_aplicación"\n'
            'sender = "tu_cuenta@gmail.com"\n'
            'use_ssl = false',
            language="toml",
        )
        st.caption("Guarda estos valores en .streamlit/secrets.toml. No subas ese archivo al repositorio.")

    st.markdown(
        '<div class="app-footer">Conversión terminada · Revisa siempre los datos antes de utilizarlos.</div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    render_app()
