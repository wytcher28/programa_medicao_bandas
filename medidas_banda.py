"""Programa de cálculo de perfil de banda a partir de logs brutos em Excel.

Este módulo lê a planilha ``dados_log_bruto.xlsx``, aba ``Dados``, calcula:

1. Transação = Bytes / Sessions
2. Perfil:
   - ``4k``  quando Transação <= 8192
   - ``16k`` quando Transação <= 32768
   - ``64k`` quando Transação > 32768

Como saída, gera uma nova planilha com:
- aba ``Dados`` preservada;
- aba ``Contas`` contendo os dados originais acrescidos de ``Transação`` e ``Perfil``;
- aba ``Resultado`` contendo um resumo por perfil, no mesmo padrão da planilha modelo.

O código foi escrito para ser empacotado com PyInstaller e executado como programa.
"""

from __future__ import annotations

import logging
import math
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


@dataclass(frozen=True)
class AppConfig:
    """Configurações principais do processamento."""

    input_file: str = "dados_log_bruto.xlsx"
    output_file: str = "dados_log_bruto_resultado.xlsx"
    source_sheet: str = "Dados"
    calc_sheet: str = "Contas"
    result_sheet: str = "Resultado"
    default_bandwidth_mbps: float = 1700.0
    reference_bandwidth_mbps: float = 3000.0


@dataclass(frozen=True)
class RequiredColumns:
    """Nomes das colunas obrigatórias na aba de origem."""

    application: str = "Application"
    bytes_col: str = "Bytes"
    sessions: str = "Sessions"


@dataclass
class ProfileSummary:
    """Resumo calculado para cada perfil de transação."""

    profile: str
    sessions: float
    traffic_percentage: float
    bandwidth_mbps: float


PROFILE_ORDER = tuple(sorted(("4k", "16k", "64k")))
HEADER_FILL = PatternFill(fill_type="solid", fgColor="D9EAF7")
TOTAL_FILL = PatternFill(fill_type="solid", fgColor="E2F0D9")


def configure_logging() -> None:
    """Configura logs simples para console e arquivo local."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler("medidas_banda.log", encoding="utf-8"),
        ],
    )


def get_runtime_dir() -> Path:
    """Retorna o diretório onde o executável/script está sendo executado."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def read_headers(sheet: Worksheet) -> dict[str, int]:
    """Lê o cabeçalho da primeira linha e retorna um mapa nome -> índice da coluna."""
    headers: dict[str, int] = {}
    for cell in sheet[1]:
        if cell.value is not None:
            headers[str(cell.value).strip()] = cell.column
    return headers


def validate_required_columns(headers: dict[str, int], required: RequiredColumns) -> None:
    """Valida se as colunas obrigatórias existem na planilha."""
    missing = [
        column_name
        for column_name in (required.application, required.bytes_col, required.sessions)
        if column_name not in headers
    ]
    if missing:
        raise ValueError(f"Colunas obrigatórias ausentes na aba Dados: {', '.join(missing)}")


def to_float(value: Any) -> float:
    """Converte valores numéricos de Excel para float de forma segura."""
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    value_as_text = str(value).strip().replace(".", "").replace(",", ".")
    return float(value_as_text)


def ceil_number(value: float | int | None) -> int:
    """Arredonda números para cima e retorna inteiro.

    Exemplo: 10.01 vira 11; 10.00 permanece 10.
    Valores nulos são tratados como zero.
    """
    if value is None:
        return 0
    return int(math.ceil(float(value)))


def calculate_transaction(bytes_value: Any, sessions_value: Any) -> int | None:
    """Calcula a transação média em bytes por sessão.

    O resultado é sempre arredondado para cima para evitar casas decimais
    na planilha final. Retorna ``None`` quando Sessions é zero ou inválido,
    evitando erro de divisão por zero.
    """
    bytes_number = to_float(bytes_value)
    sessions_number = to_float(sessions_value)
    if sessions_number <= 0:
        return None
    return ceil_number(bytes_number / sessions_number)


def classify_profile(transaction_value: float | None) -> str:
    """Classifica a transação conforme as faixas solicitadas."""
    if transaction_value is None:
        return "Sem Sessão"
    if transaction_value <= 8192:
        return "4k"
    if transaction_value <= 32768:
        return "16k"
    return "64k"


def remove_sheet_if_exists(workbook: Workbook, sheet_name: str) -> None:
    """Remove uma aba caso ela já exista no arquivo de saída."""
    if sheet_name in workbook.sheetnames:
        del workbook[sheet_name]


def copy_sheet_values(source: Worksheet, target: Worksheet) -> None:
    """Copia valores da aba de origem para uma aba de destino."""
    for row in source.iter_rows():
        for cell in row:
            target[cell.coordinate].value = cell.value


def sort_sheet_by_first_column(sheet: Worksheet, header_row: int = 1) -> None:
    """Ordena as linhas da planilha pela coluna A em ordem crescente.

    Apenas as linhas abaixo do cabeçalho são ordenadas. O cabeçalho permanece
    na posição original. A ordenação é feita pelos valores exibidos na coluna A.
    """
    if sheet.max_row <= header_row + 1:
        return

    rows = list(sheet.iter_rows(min_row=header_row + 1, values_only=True))
    rows.sort(key=lambda row: str(row[0] or "").strip().lower())

    for row_index, row_values in enumerate(rows, start=header_row + 1):
        for column_index, value in enumerate(row_values, start=1):
            sheet.cell(row=row_index, column=column_index, value=value)


def create_calc_sheet(workbook: Workbook, source_sheet: Worksheet, config: AppConfig) -> Worksheet:
    """Cria a aba Contas com as colunas Transação e Perfil."""
    remove_sheet_if_exists(workbook, config.calc_sheet)
    calc_sheet = workbook.create_sheet(config.calc_sheet)
    copy_sheet_values(source_sheet, calc_sheet)

    headers = read_headers(calc_sheet)
    required = RequiredColumns()
    validate_required_columns(headers, required)

    transaction_col = calc_sheet.max_column + 1
    profile_col = calc_sheet.max_column + 2
    calc_sheet.cell(row=1, column=transaction_col, value="Transação")
    calc_sheet.cell(row=1, column=profile_col, value="Perfil")

    for row_number in range(2, calc_sheet.max_row + 1):
        bytes_value = calc_sheet.cell(row=row_number, column=headers[required.bytes_col]).value
        sessions_value = calc_sheet.cell(row=row_number, column=headers[required.sessions]).value
        transaction = calculate_transaction(bytes_value, sessions_value)
        profile = classify_profile(transaction)
        calc_sheet.cell(row=row_number, column=transaction_col, value=transaction)
        calc_sheet.cell(row=row_number, column=profile_col, value=profile)

    sort_sheet_by_first_column(calc_sheet, header_row=1)
    format_calc_sheet(calc_sheet)
    return calc_sheet


def build_summary(calc_sheet: Worksheet, config: AppConfig) -> list[ProfileSummary]:
    """Agrupa sessões por perfil e calcula percentual e banda estimada."""
    headers = read_headers(calc_sheet)
    required = RequiredColumns()
    total_sessions = 0.0
    sessions_by_profile: dict[str, float] = {profile: 0.0 for profile in PROFILE_ORDER}

    for row_number in range(2, calc_sheet.max_row + 1):
        profile = str(calc_sheet.cell(row=row_number, column=headers["Perfil"]).value or "").strip()
        sessions = to_float(calc_sheet.cell(row=row_number, column=headers[required.sessions]).value)
        if profile in sessions_by_profile:
            sessions_by_profile[profile] += sessions
            total_sessions += sessions

    if total_sessions <= 0:
        return [ProfileSummary(profile, 0.0, 0.0, 0.0) for profile in PROFILE_ORDER]

    return sorted(
        [
            ProfileSummary(
                profile=profile,
                sessions=ceil_number(sessions_by_profile[profile]),
                traffic_percentage=sessions_by_profile[profile] / total_sessions,
                bandwidth_mbps=ceil_number(
                    (sessions_by_profile[profile] / total_sessions) * config.default_bandwidth_mbps
                ),
            )
            for profile in PROFILE_ORDER
        ],
        key=lambda item: item.profile.lower(),
    )


def create_result_sheet(workbook: Workbook, calc_sheet: Worksheet, config: AppConfig) -> Worksheet:
    """Cria a aba Resultado no padrão da planilha modelo."""
    remove_sheet_if_exists(workbook, config.result_sheet)
    result_sheet = workbook.create_sheet(config.result_sheet, 0)

    summary = build_summary(calc_sheet, config)
    total_sessions = sum(item.sessions for item in summary)

    result_sheet.append([])
    result_sheet.append([])
    result_sheet.append(["Row Labels", "% do Trafego", "Total Conexoes", "Cálculo de Banda em Mbps"])

    for item in summary:
        result_sheet.append([
            item.profile,
            item.traffic_percentage,
            item.sessions,
            item.bandwidth_mbps,
        ])

    result_sheet.append([
        "Total Geral",
        1 if total_sessions else 0,
        ceil_number(total_sessions),
        ceil_number(config.default_bandwidth_mbps),
    ])

    format_result_sheet(result_sheet)
    return result_sheet


def format_header(sheet: Worksheet, row_number: int = 1) -> None:
    """Aplica formatação padrão ao cabeçalho."""
    for cell in sheet[row_number]:
        if cell.value is not None:
            cell.font = Font(bold=True)
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(horizontal="center")


def autofit_columns(sheet: Worksheet) -> None:
    """Ajusta largura das colunas com limite para manter boa legibilidade."""
    for column_cells in sheet.columns:
        column_letter = get_column_letter(column_cells[0].column)
        max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
        sheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 35)


def format_calc_sheet(sheet: Worksheet) -> None:
    """Formata a aba Contas."""
    format_header(sheet, 1)
    headers = read_headers(sheet)
    if "Transação" in headers:
        for cell in sheet.iter_cols(min_col=headers["Transação"], max_col=headers["Transação"], min_row=2):
            for item in cell:
                item.number_format = "#,##0"
    sheet.freeze_panes = "A2"
    autofit_columns(sheet)


def format_result_sheet(sheet: Worksheet) -> None:
    """Formata a aba Resultado."""
    format_header(sheet, 3)
    for row in range(4, sheet.max_row + 1):
        sheet.cell(row=row, column=2).number_format = "0%"
        sheet.cell(row=row, column=3).number_format = "#,##0"
        sheet.cell(row=row, column=4).number_format = "#,##0"

    for cell in sheet[sheet.max_row]:
        cell.font = Font(bold=True)
        cell.fill = TOTAL_FILL

    sheet.freeze_panes = "A4"
    autofit_columns(sheet)


def process_workbook(config: AppConfig) -> Path:
    """Executa o processamento completo da planilha."""
    runtime_dir = get_runtime_dir()
    input_path = runtime_dir / config.input_file
    output_path = runtime_dir / config.output_file

    if not input_path.exists():
        raise FileNotFoundError(f"Arquivo de entrada não encontrado: {input_path}")

    logging.info("Lendo planilha: %s", input_path)
    workbook = load_workbook(input_path)

    if config.source_sheet not in workbook.sheetnames:
        raise ValueError(f"Aba obrigatória não encontrada: {config.source_sheet}")

    source_sheet = workbook[config.source_sheet]
    calc_sheet = create_calc_sheet(workbook, source_sheet, config)
    create_result_sheet(workbook, calc_sheet, config)

    workbook.save(output_path)
    logging.info("Planilha processada com sucesso: %s", output_path)
    return output_path


def main() -> int:
    """Ponto de entrada do programa."""
    configure_logging()
    config = AppConfig()

    try:
        output_path = process_workbook(config)
        print(f"Processamento concluído com sucesso: {output_path}")
        return 0
    except Exception as exc:  # noqa: BLE001 - captura final para execução amigável ao usuário.
        logging.exception("Erro ao processar a planilha")
        print(f"Erro ao processar a planilha: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
