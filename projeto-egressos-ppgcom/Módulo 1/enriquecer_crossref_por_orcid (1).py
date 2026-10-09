#!/usr/bin/env python3
"""Busca publicacoes na Crossref usando os ORCIDs de uma planilha.

Exemplo no Windows:
    python enriquecer_crossref_por_orcid.py egressos_ppgcom_com_orcid.xlsx --email voce@email.com

O resultado padrao e "publicacoes_crossref_por_orcid.xlsx" e contem:
  - Publicacoes: uma linha por egresso e publicacao;
  - Resumo: totais encontrados para cada egresso;
  - Sem_ORCID: registros que nao puderam ser consultados.
"""

from __future__ import annotations

import argparse
import html
import re
import time
import unicodedata
from pathlib import Path
from typing import Any, Iterator

try:
    import pandas as pd
    import requests
except ModuleNotFoundError as error:
    raise SystemExit(
        f"A biblioteca '{error.name}' nao esta instalada. Execute:\n"
        "python -m pip install pandas requests openpyxl"
    ) from error


API_URL = "https://api.crossref.org/works"
ORCID_PATTERN = re.compile(r"\b\d{4}-\d{4}-\d{4}-[\dX]{4}\b", re.I)
ORCID_COLUMNS = {"orcid", "orcid_id", "id_orcid", "orcid_do_egresso"}
NAME_COLUMNS = {"nome", "nome_egresso", "egresso", "name", "author_name"}
LATTES_COLUMNS = {"id_lattes", "lattes", "lattes_id"}


def normalize_column(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"[^a-z0-9]+", "_", text.lower()).strip("_")


def find_column(columns: list[Any], candidates: set[str]) -> Any | None:
    normalized = {normalize_column(column): column for column in columns}
    for candidate in candidates:
        if candidate in normalized:
            return normalized[candidate]
    return None


def clean_orcid(value: Any) -> str | None:
    if pd.isna(value):
        return None
    match = ORCID_PATTERN.search(str(value).strip())
    return match.group(0).upper() if match else None


def first(value: Any, default: str = "") -> str:
    return str(value[0]) if isinstance(value, list) and value else default


def join_values(value: Any, separator: str = "; ") -> str:
    return separator.join(str(item) for item in value if item) if isinstance(value, list) else ""


def publication_date(item: dict[str, Any]) -> tuple[str, str]:
    for field in ("published-print", "published-online", "published", "issued"):
        parts = item.get(field, {}).get("date-parts", [])
        if parts and parts[0]:
            values = parts[0]
            year = str(values[0]) if values else ""
            date = "-".join(
                [year]
                + ([f"{values[1]:02d}"] if len(values) > 1 else [])
                + ([f"{values[2]:02d}"] if len(values) > 2 else [])
            )
            return year, date
    return "", ""


def strip_markup(value: str) -> str:
    if not value:
        return ""
    value = re.sub(r"<[^>]+>", " ", value)
    return re.sub(r"\s+", " ", html.unescape(value)).strip()


def author_name(author: dict[str, Any]) -> str:
    return " ".join(
        part for part in (author.get("given", ""), author.get("family", "")) if part
    ).strip()


def author_orcid(author: dict[str, Any]) -> str | None:
    return clean_orcid(author.get("ORCID", ""))


def work_to_row(
    person_name: str,
    person_orcid: str,
    lattes_id: str,
    item: dict[str, Any],
) -> dict[str, Any]:
    authors = item.get("author", []) or []
    names = [author_name(author) for author in authors]
    target_positions = [
        str(index + 1)
        for index, author in enumerate(authors)
        if author_orcid(author) == person_orcid
    ]
    affiliations = sorted(
        {
            affiliation.get("name", "").strip()
            for author in authors
            for affiliation in (author.get("affiliation", []) or [])
            if affiliation.get("name", "").strip()
        }
    )
    funders = sorted(
        {funder.get("name", "").strip() for funder in item.get("funder", []) if funder.get("name")}
    )
    year, date = publication_date(item)
    licenses = [entry.get("URL", "") for entry in item.get("license", []) if entry.get("URL")]
    links = [entry.get("URL", "") for entry in item.get("link", []) if entry.get("URL")]

    return {
        "Nome_egresso": person_name,
        "ORCID_egresso": person_orcid,
        "ID_Lattes": lattes_id,
        "DOI": item.get("DOI", ""),
        "Titulo": first(item.get("title")),
        "Subtitulo": first(item.get("subtitle")),
        "Tipo": item.get("type", ""),
        "Ano_publicacao": year,
        "Data_publicacao": date,
        "Periodico_ou_veiculo": first(item.get("container-title")),
        "Autores": "; ".join(name for name in names if name),
        "Quantidade_autores": len(authors),
        "Posicao_egresso_autoria": "; ".join(target_positions),
        "Colaboradores": "; ".join(
            name
            for author, name in zip(authors, names)
            if name and author_orcid(author) != person_orcid
        ),
        "Afiliacoes": "; ".join(affiliations),
        "Editora": item.get("publisher", ""),
        "ISSN": join_values(item.get("ISSN", [])),
        "ISBN": join_values(item.get("ISBN", [])),
        "Volume": item.get("volume", ""),
        "Numero": item.get("issue", ""),
        "Paginas": item.get("page", ""),
        "Assuntos": join_values(item.get("subject", [])),
        "Resumo": strip_markup(item.get("abstract", "")),
        "Citacoes_Crossref": item.get("is-referenced-by-count", 0),
        "Quantidade_referencias": item.get("reference-count", 0),
        "Financiadores": "; ".join(funders),
        "Licencas": "; ".join(licenses),
        "URL_Crossref": item.get("URL", ""),
        "Links_texto_completo": "; ".join(links),
        "Fonte": "Crossref REST API",
    }


class CrossrefClient:
    def __init__(self, email: str, interval: float = 0.25, timeout: int = 45):
        self.email = email
        self.interval = interval
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update(
            {"User-Agent": f"PPGCOM-Egressos/1.0 (mailto:{self.email})"}
        )

    def get(self, params: dict[str, Any]) -> requests.Response:
        params = {**params, "mailto": self.email}
        for attempt in range(5):
            response = self.session.get(API_URL, params=params, timeout=self.timeout)
            if response.status_code == 429 or response.status_code >= 500:
                if attempt == 4:
                    response.raise_for_status()
                retry_after = response.headers.get("Retry-After", "")
                wait = float(retry_after) if retry_after.isdigit() else min(30, 2**attempt)
                time.sleep(wait)
                continue
            response.raise_for_status()
            time.sleep(self.interval)
            return response
        raise RuntimeError("A Crossref nao respondeu apos varias tentativas.")

    def works_by_orcid(self, orcid: str) -> Iterator[dict[str, Any]]:
        cursor = "*"
        while cursor:
            response = self.get(
                {
                    "filter": f"orcid:{orcid}",
                    "rows": 1000,
                    "cursor": cursor,
                }
            )
            message = response.json().get("message", {})
            items = message.get("items", [])
            yield from items
            next_cursor = message.get("next-cursor")
            if not items or not next_cursor or next_cursor == cursor:
                break
            cursor = next_cursor


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Busca na Crossref as publicacoes dos egressos usando ORCID."
    )
    parser.add_argument("arquivo", type=Path, help="Planilha .xlsx com nome e ORCID")
    parser.add_argument("--email", required=True, help="E-mail para o acesso educado da Crossref")
    parser.add_argument(
        "--saida", type=Path, default=Path("publicacoes_crossref_por_orcid.xlsx")
    )
    parser.add_argument("--coluna-orcid", help="Nome exato da coluna de ORCID")
    parser.add_argument("--coluna-nome", help="Nome exato da coluna do egresso")
    parser.add_argument("--coluna-lattes", help="Nome exato da coluna de ID Lattes")
    parser.add_argument("--aba", default=0, help="Nome ou numero da aba de entrada")
    parser.add_argument("--intervalo", type=float, default=0.25, help="Pausa entre consultas")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.arquivo.exists():
        raise SystemExit(f"Arquivo nao encontrado: {args.arquivo}")
    if args.arquivo.suffix.lower() not in {".xlsx", ".xls"}:
        raise SystemExit("Use uma planilha Excel .xlsx ou .xls.")

    sheet: str | int = int(args.aba) if str(args.aba).isdigit() else args.aba
    people = pd.read_excel(args.arquivo, sheet_name=sheet, dtype=str).fillna("")
    orcid_column = args.coluna_orcid or find_column(list(people.columns), ORCID_COLUMNS)
    name_column = args.coluna_nome or find_column(list(people.columns), NAME_COLUMNS)
    lattes_column = args.coluna_lattes or find_column(list(people.columns), LATTES_COLUMNS)

    if not orcid_column:
        raise SystemExit(
            "Nao encontrei a coluna ORCID. Use --coluna-orcid \"Nome da coluna\"."
        )
    if args.coluna_orcid and args.coluna_orcid not in people.columns:
        raise SystemExit(f"A coluna '{args.coluna_orcid}' nao existe.")

    client = CrossrefClient(args.email, args.intervalo)
    publication_rows: list[dict[str, Any]] = []
    summary_rows: list[dict[str, Any]] = []
    no_orcid_rows: list[dict[str, Any]] = []
    cache: dict[str, list[dict[str, Any]]] = {}

    for index, row in people.iterrows():
        person_name = str(row.get(name_column, "")).strip() if name_column else ""
        lattes_id = str(row.get(lattes_column, "")).strip() if lattes_column else ""
        orcid = clean_orcid(row.get(orcid_column, ""))

        if not orcid:
            no_orcid_rows.append(
                {"Nome_egresso": person_name, "ID_Lattes": lattes_id, "Motivo": "ORCID ausente ou invalido"}
            )
            print(f"[{index + 1}/{len(people)}] {person_name}: sem ORCID")
            continue

        try:
            if orcid not in cache:
                cache[orcid] = list(client.works_by_orcid(orcid))
            works = cache[orcid]
            publication_rows.extend(
                work_to_row(person_name, orcid, lattes_id, item) for item in works
            )
            summary_rows.append(
                {
                    "Nome_egresso": person_name,
                    "ORCID": orcid,
                    "ID_Lattes": lattes_id,
                    "Publicacoes_encontradas_Crossref": len(works),
                    "Soma_citacoes_Crossref": sum(
                        int(item.get("is-referenced-by-count", 0) or 0) for item in works
                    ),
                    "Status": "Encontrado" if works else "Nenhuma publicacao encontrada",
                }
            )
            print(f"[{index + 1}/{len(people)}] {person_name}: {len(works)} publicacao(oes)")
        except requests.RequestException as error:
            summary_rows.append(
                {
                    "Nome_egresso": person_name,
                    "ORCID": orcid,
                    "ID_Lattes": lattes_id,
                    "Publicacoes_encontradas_Crossref": 0,
                    "Soma_citacoes_Crossref": 0,
                    "Status": f"Erro na consulta: {error}",
                }
            )
            print(f"[{index + 1}/{len(people)}] {person_name}: erro na consulta")

    publications = pd.DataFrame(publication_rows)
    summary = pd.DataFrame(summary_rows)
    no_orcid = pd.DataFrame(no_orcid_rows)

    if not publications.empty:
        publications = publications.drop_duplicates(
            subset=["ORCID_egresso", "DOI", "Titulo"], keep="first"
        ).sort_values(["Nome_egresso", "Ano_publicacao", "Titulo"], ascending=[True, False, True])

    args.saida.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(args.saida, engine="openpyxl") as writer:
        publications.to_excel(writer, sheet_name="Publicacoes", index=False)
        summary.to_excel(writer, sheet_name="Resumo", index=False)
        no_orcid.to_excel(writer, sheet_name="Sem_ORCID", index=False)

        for sheet_name in writer.book.sheetnames:
            worksheet = writer.book[sheet_name]
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            for cell in worksheet[1]:
                cell.font = cell.font.copy(bold=True, color="FFFFFF")
                cell.fill = cell.fill.copy(fill_type="solid", fgColor="1F4E78")
            for column_cells in worksheet.columns:
                width = min(55, max(12, max(len(str(cell.value or "")) for cell in column_cells) + 2))
                worksheet.column_dimensions[column_cells[0].column_letter].width = width

    print("\nConsulta concluida.")
    print(f"Egressos consultados: {len(summary)}")
    print(f"Publicacoes salvas: {len(publications)}")
    print(f"Egressos sem ORCID: {len(no_orcid)}")
    print(f"Resultado: {args.saida.resolve()}")


if __name__ == "__main__":
    main()
