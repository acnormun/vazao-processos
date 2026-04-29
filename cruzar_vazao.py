#!/usr/bin/env python3
import argparse
import csv
import datetime as dt
import statistics
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path


NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

OUTPUT_COLUMNS = [
    "Numero do processo",
    "Tarefa",
    "Assessor",
    "Data entrada",
    "Data saida",
    "Vazao (dias)",
]


def normalize_header(value):
    text = str(value or "").strip().lower()
    replacements = str.maketrans(
        {
            "á": "a",
            "à": "a",
            "â": "a",
            "ã": "a",
            "é": "e",
            "ê": "e",
            "í": "i",
            "ó": "o",
            "ô": "o",
            "õ": "o",
            "ú": "u",
            "ç": "c",
            "º": "o",
        }
    )
    return re.sub(r"[^a-z0-9]+", " ", text.translate(replacements)).strip()


def normalize_spaces(value):
    return re.sub(r"\s+", " ", str(value or "").strip())


def normalize_assessor(value):
    return normalize_spaces(value).upper()


def find_column(headers, candidates):
    normalized = {normalize_header(header): header for header in headers}
    for candidate in candidates:
        key = normalize_header(candidate)
        if key in normalized:
            return normalized[key]
    raise ValueError(
        "Nao encontrei nenhuma destas colunas: "
        + ", ".join(candidates)
        + ". Colunas disponiveis: "
        + ", ".join(map(str, headers))
    )


def parse_date(value):
    if value is None:
        return None

    if isinstance(value, dt.datetime):
        return value.date()

    if isinstance(value, dt.date):
        return value

    text = str(value).strip()
    if not text:
        return None

    if re.fullmatch(r"\d+(\.\d+)?", text):
        return dt.date(1899, 12, 30) + dt.timedelta(days=float(text))

    for fmt in ("%d/%m/%Y", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            parsed = dt.datetime.strptime(text, fmt)
            return parsed.date()
        except ValueError:
            pass

    raise ValueError(f"Data em formato nao reconhecido: {value!r}")


def format_date(value):
    if value is None:
        return ""
    return value.strftime("%d/%m/%Y")


def read_distribuicoes(path):
    with open(path, "r", encoding="utf-8-sig", newline="") as file:
        sample = file.read(4096)
        file.seek(0)
        dialect = csv.Sniffer().sniff(sample, delimiters=";,")
        rows = list(csv.DictReader(file, dialect=dialect))

    if not rows:
        return {}

    headers = rows[0].keys()
    processo_col = find_column(headers, ["Nº Processo", "Numero do Processo", "Número do Processo"])
    entrada_col = find_column(headers, ["Data de Chegada", "Data Entrada", "Data"])
    assessor_col = find_column(headers, ["Responsável", "Responsavel", "Assessor"])

    distribuicoes = {}
    for row in rows:
        processo = str(row.get(processo_col, "")).strip()
        if not processo:
            continue
        distribuicoes.setdefault(processo, []).append(
            {
                "entrada": parse_date(row.get(entrada_col)),
                "assessor": normalize_assessor(row.get(assessor_col)),
            }
        )

    for entradas in distribuicoes.values():
        entradas.sort(key=lambda item: item["entrada"] or dt.date.min)

    return distribuicoes


def column_index(cell_reference):
    letters = "".join(ch for ch in cell_reference if ch.isalpha())
    index = 0
    for letter in letters:
        index = index * 26 + ord(letter.upper()) - ord("A") + 1
    return index - 1


def read_shared_strings(archive):
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []

    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings = []
    for item in root.findall("main:si", NS):
        strings.append("".join(text.text or "" for text in item.findall(".//main:t", NS)))
    return strings


def first_sheet_path(archive):
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    first_sheet = workbook.find("main:sheets/main:sheet", NS)
    if first_sheet is None:
        raise ValueError("O arquivo XLSX nao tem planilhas.")

    relationship_id = first_sheet.attrib[f"{{{NS['rel']}}}id"]
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    for rel in rels.findall("pkgrel:Relationship", NS):
        if rel.attrib["Id"] == relationship_id:
            target = rel.attrib["Target"]
            if target.startswith("/"):
                return target.lstrip("/")
            return posixpath.normpath(posixpath.join("xl", target))

    raise ValueError("Nao consegui localizar a primeira planilha no XLSX.")


def read_xlsx_rows(path):
    with zipfile.ZipFile(path) as archive:
        shared_strings = read_shared_strings(archive)
        sheet_path = first_sheet_path(archive)
        root = ET.fromstring(archive.read(sheet_path))

    rows = []
    for row in root.findall(".//main:sheetData/main:row", NS):
        cells = {}
        for cell in row.findall("main:c", NS):
            reference = cell.attrib.get("r", "A")
            value_node = cell.find("main:v", NS)
            inline_node = cell.find("main:is/main:t", NS)
            value = ""
            if inline_node is not None:
                value = inline_node.text or ""
            elif value_node is not None:
                value = value_node.text or ""
                if cell.attrib.get("t") == "s" and value:
                    value = shared_strings[int(value)]
            cells[column_index(reference)] = value
        if cells:
            rows.append([cells.get(index, "") for index in range(max(cells) + 1)])
    return rows


def read_saidas(path):
    raw_rows = read_xlsx_rows(path)
    if not raw_rows:
        return []

    headers = [str(value).strip() for value in raw_rows[0]]
    processo_col = find_column(headers, ["Número do Processo", "Numero do Processo", "Nº Processo"])
    saida_col = find_column(headers, ["Data", "Data Saída", "Data Saida"])
    tarefa_col = find_column(headers, ["Tarefa"])
    assessor_col = find_column(headers, ["Assessor"])

    saidas = []
    for raw in raw_rows[1:]:
        row = dict(zip(headers, raw + [""] * (len(headers) - len(raw))))
        processo = str(row.get(processo_col, "")).strip()
        if not processo or processo.lower() == "totais":
            continue
        saidas.append(
            {
                "processo": processo,
                "saida": parse_date(row.get(saida_col)),
                "tarefa": str(row.get(tarefa_col, "")).strip(),
                "assessor": normalize_assessor(row.get(assessor_col)),
            }
        )
    return saidas


def cruzar(distribuicoes, saidas):
    resultado = []
    saidas_indexadas = list(enumerate(saidas))
    saidas_usadas = set()
    saidas_por_processo = {}
    for indice, saida in saidas_indexadas:
        saidas_por_processo.setdefault(saida["processo"], []).append((indice, saida))

    for saidas_do_processo in saidas_por_processo.values():
        saidas_do_processo.sort(key=lambda item: item[1]["saida"] or dt.date.max)

    for processo, entradas in distribuicoes.items():
        saidas_do_processo = saidas_por_processo.get(processo, [])
        for index, entrada in enumerate(entradas):
            data_entrada = entrada["entrada"]
            proxima_entrada = None
            if index + 1 < len(entradas):
                proxima_entrada = entradas[index + 1]["entrada"]

            saidas_do_ciclo = [
                (indice_saida, saida)
                for indice_saida, saida in saidas_do_processo
                if data_entrada
                and saida["saida"]
                and data_entrada <= saida["saida"]
                and (proxima_entrada is None or saida["saida"] < proxima_entrada)
            ]

            if not saidas_do_ciclo:
                resultado.append(
                    {
                        "Numero do processo": processo,
                        "Tarefa": "",
                        "Assessor": entrada["assessor"],
                        "Data entrada": format_date(data_entrada),
                        "Data saida": "",
                        "Vazao (dias)": "",
                    }
                )
                continue

            for indice_saida, saida in saidas_do_ciclo:
                saidas_usadas.add(indice_saida)
                data_saida = saida["saida"]
                resultado.append(
                    {
                        "Numero do processo": processo,
                        "Tarefa": saida["tarefa"],
                        "Assessor": saida["assessor"] or entrada["assessor"],
                        "Data entrada": format_date(data_entrada),
                        "Data saida": format_date(data_saida),
                        "Vazao (dias)": (data_saida - data_entrada).days,
                    }
                )

    for indice_saida, saida in saidas_indexadas:
        if indice_saida in saidas_usadas:
            continue
        resultado.append(
            {
                "Numero do processo": saida["processo"],
                "Tarefa": saida["tarefa"],
                "Assessor": saida["assessor"],
                "Data entrada": "",
                "Data saida": format_date(saida["saida"]),
                "Vazao (dias)": "",
            }
        )
    return resultado


def count_distribuicoes(distribuicoes):
    return sum(len(entradas) for entradas in distribuicoes.values())


def count_pendentes(rows):
    return sum(1 for row in rows if not row["Data saida"])


def count_saidas_sem_entrada(rows):
    return sum(1 for row in rows if not row["Data entrada"])


def tarefas_disponiveis(rows):
    return sorted({row["Tarefa"] for row in rows if row.get("Tarefa")})


def periodo_datas(periodo, hoje=None, inicio=None, fim=None):
    hoje = hoje or dt.date.today()
    if periodo == "Este mes":
        return hoje.replace(day=1), hoje
    if periodo == "Ultimos 30 dias":
        return hoje - dt.timedelta(days=30), hoje
    if periodo == "Este ano":
        return dt.date(hoje.year, 1, 1), hoje
    if periodo == "Ultimos 12 meses":
        return hoje - dt.timedelta(days=365), hoje
    if periodo == "Personalizado":
        return inicio, fim
    return None, None


def row_date(row, column):
    try:
        return parse_date(row.get(column))
    except ValueError:
        return None


def in_period(value, inicio, fim):
    if value is None:
        return False
    if inicio and value < inicio:
        return False
    if fim and value > fim:
        return False
    return True


def resumo_por_assessor(rows, tarefa=None, periodo="Todos", inicio=None, fim=None, hoje=None):
    inicio, fim = periodo_datas(periodo, hoje=hoje, inicio=inicio, fim=fim)
    resumo = {}

    for row in rows:
        if tarefa and row.get("Tarefa") != tarefa:
            continue

        assessor = row.get("Assessor") or "Sem assessor"
        item = resumo.setdefault(
            assessor,
            {
                "Assessor": assessor,
                "Entradas": 0,
                "Saidas": 0,
                "Saldo": 0,
                "Vazao media": "",
                "Vazao mediana": "",
                "Menor vazao": "",
                "Maior vazao": "",
                "Entradas sem saida": 0,
                "Saidas sem entrada": 0,
                "_vazoes": [],
            },
        )

        data_saida = row_date(row, "Data saida")
        data_entrada = row_date(row, "Data entrada")
        tem_saida = bool(row.get("Data saida"))
        tem_entrada = bool(row.get("Data entrada"))
        dentro_saida = in_period(data_saida, inicio, fim) if inicio or fim else tem_saida
        dentro_entrada = in_period(data_entrada, inicio, fim) if inicio or fim else tem_entrada

        if tem_saida and dentro_saida:
            item["Saidas"] += 1
            if row.get("Vazao (dias)") != "":
                item["_vazoes"].append(int(row["Vazao (dias)"]))

        if tem_entrada and dentro_entrada:
            item["Entradas"] += 1

        if tem_entrada and not tem_saida and dentro_entrada:
            item["Entradas sem saida"] += 1

        if tem_saida and not tem_entrada and dentro_saida:
            item["Saidas sem entrada"] += 1

    final = []
    for item in resumo.values():
        vazoes = item.pop("_vazoes")
        if vazoes:
            item["Vazao media"] = round(sum(vazoes) / len(vazoes), 1)
            item["Vazao mediana"] = round(statistics.median(vazoes), 1)
            item["Menor vazao"] = min(vazoes)
            item["Maior vazao"] = max(vazoes)
        item["Saldo"] = item["Entradas"] - item["Saidas"]
        if item["Entradas"] or item["Saidas"] or item["Entradas sem saida"] or item["Saidas sem entrada"]:
            final.append(item)

    return sorted(final, key=lambda item: (item["Saidas"], item["Entradas"]), reverse=True)


def write_csv(path, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=OUTPUT_COLUMNS, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)


def excel_column_name(index):
    name = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(ord("A") + remainder) + name
    return name


def xml_escape(value):
    return (
        str(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def worksheet_xml(rows):
    all_rows = [OUTPUT_COLUMNS] + [[row[column] for column in OUTPUT_COLUMNS] for row in rows]
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        "<sheetData>",
    ]
    for row_number, row in enumerate(all_rows, start=1):
        parts.append(f'<row r="{row_number}">')
        for column_number, value in enumerate(row):
            cell = f"{excel_column_name(column_number)}{row_number}"
            if isinstance(value, int):
                parts.append(f'<c r="{cell}"><v>{value}</v></c>')
            else:
                parts.append(
                    f'<c r="{cell}" t="inlineStr"><is><t>{xml_escape(value)}</t></is></c>'
                )
        parts.append("</row>")
    parts.extend(["</sheetData>", "</worksheet>"])
    return "".join(parts)


def write_xlsx(path, rows):
    files = {
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>""",
        "_rels/.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""",
        "xl/workbook.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Vazao" sheetId="1" r:id="rId1"/></sheets></workbook>""",
        "xl/_rels/workbook.xml.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""",
        "xl/worksheets/sheet1.xml": worksheet_xml(rows),
    }
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, content in files.items():
            archive.writestr(name, content)


def write_output(path, rows):
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        write_xlsx(path, rows)
    else:
        write_csv(path, rows)


def pdf_escape(value):
    text = str(value)
    return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def pdf_text(value):
    return pdf_escape(str(value).encode("latin-1", "replace").decode("latin-1"))


def write_pdf_report(path, resumo, titulo="Relatorio de vazao por assessor", filtros=""):
    width, height = 842, 595
    margin = 36
    row_height = 18
    columns = [
        ("Assessor", 205),
        ("Entradas", 55),
        ("Saidas", 50),
        ("Saldo", 45),
        ("Media", 55),
        ("Mediana", 60),
        ("Min", 40),
        ("Max", 40),
        ("Sem saida", 65),
        ("Sem entrada", 75),
    ]
    content_width = sum(size for _label, size in columns)
    pages = []

    def draw_header(lines):
        y = height - margin
        lines.append(f"BT /F1 16 Tf {margin} {y} Td ({pdf_text(titulo)}) Tj ET")
        y -= 20
        if filtros:
            lines.append(f"BT /F1 9 Tf {margin} {y} Td ({pdf_text(filtros)}) Tj ET")
            y -= 18
        lines.append(f"BT /F1 9 Tf {margin} {y} Td ({pdf_text('Gerado em ' + dt.datetime.now().strftime('%d/%m/%Y %H:%M'))}) Tj ET")
        y -= 18
        return y

    def draw_table_header(lines, y):
        x = margin
        lines.append(f"{margin} {y - 4} {content_width} 16 re S")
        for label, size in columns:
            lines.append(f"BT /F1 8 Tf {x + 3} {y} Td ({pdf_text(label)}) Tj ET")
            x += size
        return y - row_height

    current = []
    y = draw_header(current)
    y = draw_table_header(current, y)

    total_entradas = 0
    total_saidas = 0
    total_sem_saida = 0
    total_sem_entrada = 0
    for item in resumo:
        if y < margin + row_height:
            pages.append(current)
            current = []
            y = draw_header(current)
            y = draw_table_header(current, y)

        values = [
            item["Assessor"][:45],
            item["Entradas"],
            item["Saidas"],
            item["Saldo"],
            item["Vazao media"],
            item["Vazao mediana"],
            item["Menor vazao"],
            item["Maior vazao"],
            item["Entradas sem saida"],
            item["Saidas sem entrada"],
        ]
        x = margin
        current.append(f"{margin} {y - 4} {content_width} 16 re S")
        for value, (_label, size) in zip(values, columns):
            current.append(f"BT /F1 7 Tf {x + 3} {y} Td ({pdf_text(value)}) Tj ET")
            x += size
        total_entradas += item["Entradas"]
        total_saidas += item["Saidas"]
        total_sem_saida += item["Entradas sem saida"]
        total_sem_entrada += item["Saidas sem entrada"]
        y -= row_height

    if y < margin + row_height:
        pages.append(current)
        current = []
        y = draw_header(current)

    total_text = (
        f"Totais: entradas={total_entradas} | saidas={total_saidas} | "
        f"saldo={total_entradas - total_saidas} | entradas sem saida={total_sem_saida} | "
        f"saidas sem entrada={total_sem_entrada}"
    )
    current.append(f"BT /F1 10 Tf {margin} {y - 8} Td ({pdf_text(total_text)}) Tj ET")
    pages.append(current)

    objects = [
        "<< /Type /Catalog /Pages 2 0 R >>",
        "",
    ]
    page_object_numbers = []
    content_object_numbers = []
    next_object = 3
    for _page in pages:
        page_object_numbers.append(next_object)
        content_object_numbers.append(next_object + 1)
        next_object += 2

    kids = " ".join(f"{number} 0 R" for number in page_object_numbers)
    objects[1] = f"<< /Type /Pages /Kids [{kids}] /Count {len(pages)} >>"

    for page_number, lines in enumerate(pages):
        content_number = content_object_numbers[page_number]
        page = (
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {width} {height}] "
            f"/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> "
            f"/Contents {content_number} 0 R >>"
        )
        stream = "\n".join(lines)
        stream_bytes = stream.encode("latin-1", "replace")
        content = (
            f"<< /Length {len(stream_bytes)} >>\nstream\n"
            + stream_bytes.decode("latin-1")
            + "\nendstream"
        )
        objects.extend([page, content])

    output = bytearray()
    output.extend(b"%PDF-1.4\n")
    offsets = [0]
    for index, obj in enumerate(objects, start=1):
        offsets.append(len(output))
        output.extend(f"{index} 0 obj\n{obj}\nendobj\n".encode("latin-1", "replace"))

    xref = len(output)
    output.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
    output.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        output.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
    output.extend(
        (
            f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\n"
            f"startxref\n{xref}\n%%EOF\n"
        ).encode("ascii")
    )
    Path(path).write_bytes(output)


def main():
    parser = argparse.ArgumentParser(
        description="Cruza distribuicoes e saidas para calcular a vazao dos processos."
    )
    parser.add_argument("distribuicoes", help="CSV com as distribuicoes/entradas.")
    parser.add_argument("saidas", help="XLSX com as saidas.")
    parser.add_argument(
        "-o",
        "--output",
        default="vazao-processos.xlsx",
        help="Arquivo de saida .xlsx ou .csv. Padrao: vazao-processos.xlsx",
    )
    args = parser.parse_args()

    distribuicoes = read_distribuicoes(Path(args.distribuicoes))
    saidas = read_saidas(Path(args.saidas))
    resultado = cruzar(distribuicoes, saidas)

    output_path = Path(args.output)
    write_output(output_path, resultado)

    processos_entrada = set(distribuicoes)
    processos_saida = {saida["processo"] for saida in saidas}
    print(f"Entradas lidas: {count_distribuicoes(distribuicoes)}")
    print(f"Processos com entrada: {len(distribuicoes)}")
    print(f"Saidas lidas: {len(saidas)}")
    print(f"Processos com cruzamento: {len(processos_entrada & processos_saida)}")
    print(f"Linhas geradas: {len(resultado)}")
    print(f"Entradas ainda sem saida: {count_pendentes(resultado)}")
    print(f"Saidas sem entrada: {count_saidas_sem_entrada(resultado)}")
    print(f"Arquivo gerado: {output_path.resolve()}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"Erro: {exc}", file=sys.stderr)
        sys.exit(1)
