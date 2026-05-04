from __future__ import annotations

import logging
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from itertools import islice
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from downloader.nfse_files import build_nfse_file_info

logger = logging.getLogger(__name__)

NAVY = colors.HexColor("#17313B")
TEAL = colors.HexColor("#0F766E")
BLUE = colors.HexColor("#2563EB")
AMBER = colors.HexColor("#B45309")
RED = colors.HexColor("#B91C1C")
INK = colors.HexColor("#23313A")
MUTED = colors.HexColor("#667085")
BG = colors.HexColor("#F7F3EC")
CARD = colors.HexColor("#FFFDF9")
LINE = colors.HexColor("#E8DFD1")
ROW_ALT = colors.HexColor("#FBF7F1")


@dataclass(slots=True)
class NFSeReportItem:
    section: str
    numero: str
    chave: str
    emissao: str
    competencia: str
    prestador_nome: str
    tomador_nome: str
    valor_servico: Decimal
    valor_retido: Decimal
    valor_liquido: Decimal
    cancelada: bool
    retida: bool
    xml_path: Path
    pdf_path: Path | None

    @property
    def situacao(self) -> str:
        base = "Prestada" if self.section == "PRESTADOS" else "Tomada"
        suffixes: list[str] = []
        if self.retida:
            suffixes.append("Retido")
        if self.cancelada:
            suffixes.append("Cancelada")
        if suffixes:
            return f"{base} - {' - '.join(suffixes)}"
        return base

    @property
    def counterparty(self) -> str:
        if self.section == "PRESTADOS":
            return self.tomador_nome or "Tomador nao informado"
        return self.prestador_nome or "Prestador nao informado"

    @property
    def effective_servico(self) -> Decimal:
        if self.cancelada:
            return Decimal("0")
        return self.valor_servico

    @property
    def effective_retido(self) -> Decimal:
        if self.cancelada:
            return Decimal("0")
        return self.valor_retido

    @property
    def effective_liquido(self) -> Decimal:
        if self.cancelada:
            return Decimal("0")
        return self.valor_liquido

    @property
    def sort_key(self) -> tuple:
        return (_parse_date_sort(self.emissao), _parse_int(self.numero), self.chave)


@dataclass(slots=True)
class NFSeReportData:
    company_name: str
    company_cnpj: str
    year: str
    month: str
    prestados: list[NFSeReportItem]
    tomados: list[NFSeReportItem]

    @property
    def competencia_label(self) -> str:
        return f"{int(self.month):02d}-{int(self.year):04d}"

    @property
    def total_documents(self) -> int:
        return len(self.prestados) + len(self.tomados)

    @property
    def canceladas(self) -> int:
        return sum(1 for item in self.all_items if item.cancelada)

    @property
    def retidas(self) -> int:
        return sum(1 for item in self.all_items if item.retida)

    @property
    def all_items(self) -> list[NFSeReportItem]:
        return [*self.prestados, *self.tomados]


def generate_report_pdf(
    period_dir: str | Path,
    cod_company: int | str,
    company_name: str,
    company_cnpj: str,
    year: str,
    month: str,
) -> Path:
    period_path = Path(period_dir)
    report_data = collect_report_data(period_path, company_name, company_cnpj, year, month)
    styles = _build_styles()
    story = _build_story(report_data, styles)

    def draw_page(canvas, doc_obj):
        canvas.saveState()
        canvas.setFillColor(BG)
        canvas.rect(0, 0, A4[0], A4[1], stroke=0, fill=1)
        canvas.setFillColor(NAVY)
        canvas.rect(doc_obj.leftMargin, A4[1] - 18 * mm, doc_obj.width, 4, stroke=0, fill=1)

        canvas.setFont("Helvetica", 8.5)
        canvas.setFillColor(MUTED)
        canvas.drawString(doc_obj.leftMargin, 10 * mm, f"{report_data.company_name} | {report_data.competencia_label}")
        canvas.drawRightString(A4[0] - doc_obj.rightMargin, 10 * mm, f"Pagina {doc_obj.page}")
        canvas.restoreState()

    targets = [
        period_path / f"Relatorio NFSe - {int(month):02d}-{int(year):04d}.pdf",
        period_path / f"Relatorio NFSe - {int(month):02d}-{int(year):04d} - {datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
    ]

    last_error: PermissionError | None = None
    for output_path in targets:
        try:
            doc = SimpleDocTemplate(
                str(output_path),
                pagesize=A4,
                leftMargin=16 * mm,
                rightMargin=16 * mm,
                topMargin=20 * mm,
                bottomMargin=16 * mm,
                title=f"Relatorio NFSe {report_data.competencia_label}",
                author="Portal NFSe",
            )
            doc.build(story, onFirstPage=draw_page, onLaterPages=draw_page)
            logger.info("Relatorio PDF gerado com sucesso em %s.", output_path)
            return output_path
        except PermissionError as exc:
            last_error = exc

    if last_error is not None:
        raise last_error
    raise RuntimeError("Nao foi possivel gerar o relatorio PDF.")


def collect_report_data(
    period_dir: str | Path,
    company_name: str,
    company_cnpj: str,
    year: str,
    month: str,
) -> NFSeReportData:
    period_path = Path(period_dir)
    prestados = _load_section_items(period_path / "PRESTADOS", "PRESTADOS", year, month)
    tomados = _load_section_items(period_path / "TOMADOS", "TOMADOS", year, month)

    prestados.sort(key=lambda item: item.sort_key)
    tomados.sort(key=lambda item: item.sort_key)

    return NFSeReportData(
        company_name=company_name.strip() or f"Empresa {company_cnpj}",
        company_cnpj=company_cnpj.strip(),
        year=str(year),
        month=f"{int(month):02d}",
        prestados=prestados,
        tomados=tomados,
    )


def _load_section_items(section_dir: Path, section: str, year: str, month: str) -> list[NFSeReportItem]:
    items: list[NFSeReportItem] = []
    if not section_dir.exists():
        return items

    for xml_path in sorted(section_dir.rglob("*.xml")):
        try:
            items.append(_parse_item(xml_path, section, year, month))
        except Exception as exc:
            logger.warning("Falha ao processar XML para relatorio (%s): %s", xml_path, exc)

    return items


def _parse_item(xml_path: Path, section: str, year: str, month: str) -> NFSeReportItem:
    xml_bytes = xml_path.read_bytes()
    root = ET.fromstring(xml_bytes)

    cancelada = any(part.lower() == "canceladas" for part in xml_path.parts)
    chave = _extract_access_key(root, xml_path)
    file_info = build_nfse_file_info(xml_bytes, chave, xml_path.stem)

    emissao = _format_date(
        _first_text_for_paths(
            root,
            [
                ("dhProc",),
                ("dhEmi",),
                ("dhEvento",),
                ("DataEmissao",),
            ],
        )
    )
    competencia = _format_competencia(
        _first_text_for_paths(
            root,
            [
                ("DPS", "infDPS", "dCompet"),
                ("dCompet",),
                ("Competencia",),
            ],
        ),
        year,
        month,
    )

    prestador_nome = _first_text_for_paths(
        root,
        [
            ("DPS", "infDPS", "prest", "xNome"),
            ("emit", "xNome"),
        ],
    )
    tomador_nome = _first_text_for_paths(
        root,
        [
            ("DPS", "infDPS", "toma", "xNome"),
            ("toma", "xNome"),
        ],
    )

    valor_servico = _parse_decimal(
        _first_text_for_paths(
            root,
            [
                ("DPS", "infDPS", "valores", "vServPrest", "vServ"),
                ("vServ",),
            ],
        )
    )
    valor_retido = _parse_decimal(_first_text_for_paths(root, [("vTotalRet",)]))
    valor_liquido = _parse_decimal(_first_text_for_paths(root, [("vLiq",)]))
    if valor_liquido == Decimal("0") and valor_servico > Decimal("0"):
        valor_liquido = max(Decimal("0"), valor_servico - valor_retido)

    pdf_path = xml_path.with_suffix(".pdf")
    if not pdf_path.exists():
        pdf_path = None

    return NFSeReportItem(
        section=section,
        numero=file_info.numero,
        chave=file_info.chave,
        emissao=emissao,
        competencia=competencia,
        prestador_nome=prestador_nome or "Prestador nao informado",
        tomador_nome=tomador_nome or "Tomador nao informado",
        valor_servico=valor_servico,
        valor_retido=valor_retido,
        valor_liquido=valor_liquido,
        cancelada=cancelada,
        retida=file_info.retido,
        xml_path=xml_path,
        pdf_path=pdf_path,
    )


def _build_story(report_data: NFSeReportData, styles) -> list:
    story: list = []

    story.append(Paragraph("Relatorio Gerencial de NFSe", styles["PortalTitle"]))
    story.append(Paragraph("Visao consolidada da competencia para conferencia rapida.", styles["PortalSubTitle"]))
    story.append(Spacer(1, 5))
    story.append(_build_info_table(report_data, styles))
    story.append(Spacer(1, 12))

    metrics = Table(
        [
            [
                _make_metric_card("Prestados", str(len(report_data.prestados)), _format_money(_sum_servico(report_data.prestados)), TEAL, styles),
                _make_metric_card("Tomados", str(len(report_data.tomados)), _format_money(_sum_servico(report_data.tomados)), BLUE, styles),
                _make_metric_card("Canceladas", str(report_data.canceladas), "Documentos cancelados", RED, styles),
                _make_metric_card("Retidos", str(report_data.retidas), _format_money(_sum_retido(report_data.all_items)), AMBER, styles),
            ]
        ],
        colWidths=[43 * mm, 43 * mm, 43 * mm, 43 * mm],
    )
    metrics.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(metrics)
    story.append(Spacer(1, 16))
    story.append(Paragraph("Resumo financeiro", styles["SectionTitle"]))
    story.append(_build_summary_table(report_data))

    story.extend(_build_section_pages("Relatorio Prestados", report_data.prestados, report_data.competencia_label, styles))
    story.extend(_build_section_pages("Relatorio Tomados", report_data.tomados, report_data.competencia_label, styles))
    return story


def _build_info_table(report_data: NFSeReportData, styles) -> Table:
    rows = [
        [
            Paragraph("<b>Empresa</b><br/>" + report_data.company_name, styles["Body"]),
            Paragraph("<b>CNPJ</b><br/>" + (report_data.company_cnpj or "Nao informado"), styles["Body"]),
        ],
        [
            Paragraph("<b>Competencia</b><br/>" + report_data.competencia_label, styles["Body"]),
            Paragraph(f"<b>Gerado em</b><br/>{datetime.now().strftime('%d/%m/%Y')}", styles["Body"]),
        ],
    ]
    table = Table(rows, colWidths=[88 * mm, 78 * mm], rowHeights=[18 * mm, 18 * mm])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), CARD),
                ("BOX", (0, 0), (-1, -1), 0.8, LINE),
                ("INNERGRID", (0, 0), (-1, -1), 0.6, LINE),
                ("LEFTPADDING", (0, 0), (-1, -1), 12),
                ("RIGHTPADDING", (0, 0), (-1, -1), 12),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    return table


def _build_summary_table(report_data: NFSeReportData) -> Table:
    data = [
        ["Categoria", "Qtde", "Base de Calculo", "ISS", "Liquido"],
        [
            "Prestados",
            str(len(report_data.prestados)),
            _format_money(_sum_servico(report_data.prestados)),
            _format_money(_sum_retido(report_data.prestados)),
            _format_money(_sum_liquido(report_data.prestados)),
        ],
        [
            "Tomados",
            str(len(report_data.tomados)),
            _format_money(_sum_servico(report_data.tomados)),
            _format_money(_sum_retido(report_data.tomados)),
            _format_money(_sum_liquido(report_data.tomados)),
        ],
        [
            "Canceladas",
            str(report_data.canceladas),
            _format_money(Decimal("0")),
            _format_money(Decimal("0")),
            _format_money(Decimal("0")),
        ],
        [
            "Retidos",
            str(report_data.retidas),
            _format_money(_sum_servico([item for item in report_data.all_items if item.retida and not item.cancelada])),
            _format_money(_sum_retido(report_data.all_items)),
            _format_money(_sum_liquido([item for item in report_data.all_items if item.retida and not item.cancelada])),
        ],
    ]
    table = Table(
        data,
        colWidths=[40 * mm, 18 * mm, 42 * mm, 30 * mm, 40 * mm],
        rowHeights=[11 * mm, 10.5 * mm, 10.5 * mm, 10.5 * mm, 10.5 * mm],
    )
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9.5),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [CARD, ROW_ALT]),
                ("GRID", (0, 0), (-1, -1), 0.6, LINE),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    return table


def _build_section_pages(title: str, rows: list[NFSeReportItem], competencia_label: str, styles) -> list:
    story: list = []
    rows_per_page = 16
    canceled = sum(1 for item in rows if item.cancelada)
    retained = sum(1 for item in rows if item.retida)
    total_label = _format_money(_sum_servico(rows))

    if not rows:
        story.append(PageBreak())
        story.append(Paragraph(title, styles["PortalTitle"]))
        story.append(Paragraph("Nenhum documento encontrado nesta secao.", styles["PortalSubTitle"]))
        story.append(_build_empty_section_card(competencia_label, styles))
        return story

    for index, chunk in enumerate(_chunked(rows, rows_per_page), start=1):
        story.append(PageBreak())
        heading = title if index == 1 else f"{title} - continuacao"
        story.append(Paragraph(heading, styles["PortalTitle"]))
        story.append(
            Paragraph(
                "Notas canceladas e retidas permanecem nesta secao conforme a origem do documento.",
                styles["PortalSubTitle"],
            )
        )
        story.append(
            _build_section_metrics(
                section_name=title.replace("Relatorio ", ""),
                section_count=len(rows),
                total_label=total_label,
                canceled=canceled,
                retained=retained,
                competencia_label=competencia_label,
                styles=styles,
            )
        )
        story.append(Spacer(1, 10))
        story.append(_build_section_table(chunk, styles))

    return story


def _build_empty_section_card(competencia_label: str, styles) -> Table:
    table = Table(
        [[Paragraph("Sem NFSe para a competencia " + competencia_label, styles["Body"])]],
        colWidths=[174 * mm],
        rowHeights=[24 * mm],
    )
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), CARD),
                ("BOX", (0, 0), (-1, -1), 0.8, LINE),
                ("LEFTPADDING", (0, 0), (-1, -1), 12),
                ("RIGHTPADDING", (0, 0), (-1, -1), 12),
                ("TOPPADDING", (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
            ]
        )
    )
    return table


def _build_section_metrics(
    section_name: str,
    section_count: int,
    total_label: str,
    canceled: int,
    retained: int,
    competencia_label: str,
    styles,
) -> Table:
    accent = TEAL if "Prestad" in section_name else BLUE
    cards = Table(
        [
            [
                _make_metric_card(section_name, str(section_count), total_label, accent, styles),
                _make_metric_card("Canceladas", str(canceled), "Mantidas nesta secao", RED, styles),
                _make_metric_card("Retidos", str(retained), "Mantidos nesta secao", AMBER, styles),
                _make_metric_card("Competencia", competencia_label, "Resumo da secao", NAVY, styles),
            ]
        ],
        colWidths=[43 * mm, 43 * mm, 43 * mm, 43 * mm],
    )
    cards.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    return cards


def _build_section_table(rows: list[NFSeReportItem], styles) -> Table:
    data = [["NFSe", "Emissao", "Chave de acesso", "Tomador/Prestador", "Valor", "Situacao"]]
    for item in rows:

        nome = item.counterparty
        item_display = nome[:27] + "..." if len(nome) > 30 else nome
        data.append(
            [
                item.numero,
                item.emissao,
                _display_access_key(item.chave),
                item_display,
                _format_money(item.valor_servico),
                _make_status(item.situacao, styles),
            ]
        )

    table = Table(
        data,
        colWidths=[14 * mm, 21 * mm, 27 * mm, 58 * mm, 24 * mm, 34 * mm],
        rowHeights=[10.2 * mm] + [8.1 * mm] * (len(data) - 1),
    )
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (4, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, 0), 8.3),
                ("FONTSIZE", (0, 1), (4, -1), 7.6),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [CARD, ROW_ALT]),
                ("GRID", (0, 0), (-1, -1), 0.55, LINE),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (1, -1), "CENTER"),
                ("ALIGN", (2, 1), (2, -1), "CENTER"),
                ("ALIGN", (4, 1), (4, -1), "RIGHT"),
                ("ALIGN", (5, 1), (5, -1), "CENTER"),
            ]
        )
    )
    return table


def _make_metric_card(title: str, number: str, detail: str, accent: colors.Color, styles) -> Table:
    rows = [
        [Paragraph(f"<font color='{accent.hexval()}'>{title.upper()}</font>", styles["MetricTitle"])],
        [Paragraph(number, styles["MetricNumber"])],
        [Paragraph(detail, styles["MetricDetail"])],
    ]
    table = Table(rows, colWidths=[42.5 * mm], rowHeights=[9 * mm, 13 * mm, 12 * mm])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), CARD),
                ("BOX", (0, 0), (-1, -1), 0.75, LINE),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    return table


def _make_status(text: str, styles) -> Paragraph:
    lowered = text.lower()
    color = INK
    if "cancel" in lowered:
        color = RED
    elif "retid" in lowered:
        color = AMBER
    elif "tomad" in lowered:
        color = BLUE
    elif "prestad" in lowered:
        color = TEAL
    return Paragraph(f"<font color='{color.hexval()}'>{text}</font>", styles["StatusCenter"])


def _build_styles():
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="PortalTitle",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=22,
            leading=27,
            textColor=NAVY,
            spaceAfter=6,
            alignment=TA_CENTER,
        )
    )
    styles.add(
        ParagraphStyle(
            name="PortalSubTitle",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10,
            leading=13,
            textColor=MUTED,
            spaceAfter=8,
            alignment=TA_CENTER,
        )
    )
    styles.add(
        ParagraphStyle(
            name="SectionTitle",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=13,
            leading=16,
            textColor=INK,
            spaceAfter=6,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Body",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9.6,
            leading=13,
            textColor=INK,
        )
    )
    styles.add(
        ParagraphStyle(
            name="MetricTitle",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=9,
            leading=11,
            textColor=NAVY,
        )
    )
    styles.add(
        ParagraphStyle(
            name="MetricNumber",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=22,
            textColor=INK,
        )
    )
    styles.add(
        ParagraphStyle(
            name="MetricDetail",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=8.4,
            leading=11,
            textColor=MUTED,
        )
    )
    styles.add(
        ParagraphStyle(
            name="StatusCenter",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.4,
            leading=9,
            textColor=INK,
            alignment=TA_CENTER,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Right",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9.4,
            leading=12,
            textColor=INK,
            alignment=TA_RIGHT,
        )
    )
    return styles


def _sum_servico(items: list[NFSeReportItem]) -> Decimal:
    total = Decimal("0")
    for item in items:
        total += item.effective_servico
    return total


def _sum_retido(items: list[NFSeReportItem]) -> Decimal:
    total = Decimal("0")
    for item in items:
        total += item.effective_retido
    return total


def _sum_liquido(items: list[NFSeReportItem]) -> Decimal:
    total = Decimal("0")
    for item in items:
        total += item.effective_liquido
    return total


def _chunked(items: list[NFSeReportItem], size: int):
    iterator = iter(items)
    while True:
        chunk = list(islice(iterator, size))
        if not chunk:
            break
        yield chunk


def _extract_access_key(root: ET.Element, xml_path: Path) -> str:
    inf_nfse = _find_first_by_local_name(root, "infNFSe")
    if inf_nfse is not None:
        raw_id = str(inf_nfse.attrib.get("Id", "")).strip()
        if raw_id:
            cleaned = re.sub(r"^NFS", "", raw_id, flags=re.IGNORECASE).strip()
            if cleaned:
                return cleaned

    stem_parts = [part.strip() for part in xml_path.stem.split(" - ") if part.strip()]
    if len(stem_parts) >= 2:
        return stem_parts[1]

    return xml_path.stem


def _display_access_key(value: str) -> str:
    raw = str(value or "").strip()
    if len(raw) <= 14:
        return raw
    return f"{raw[:8]}...{raw[-5:]}"


def _format_money(value: Decimal) -> str:
    quantized = value.quantize(Decimal("0.01"))
    text = f"{quantized:,.2f}"
    text = text.replace(",", "_").replace(".", ",").replace("_", ".")
    return f"R$ {text}"


def _format_date(raw: str) -> str:
    value = str(raw or "").strip()
    if not value:
        return ""

    normalized = value.replace("Z", "").strip()
    for candidate in (normalized, normalized[:19], normalized[:10]):
        for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(candidate, fmt).strftime("%d/%m/%Y")
            except ValueError:
                continue

    return value[:10]


def _format_competencia(raw: str, default_year: str, default_month: str) -> str:
    value = str(raw or "").strip()
    if not value:
        return f"{int(default_month):02d}-{int(default_year):04d}"

    normalized = value.replace("Z", "").strip()
    for candidate in (normalized, normalized[:10], normalized[:7]):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y-%m"):
            try:
                dt = datetime.strptime(candidate, fmt)
                return f"{dt.month:02d}-{dt.year:04d}"
            except ValueError:
                continue

    digits = "".join(char for char in normalized if char.isdigit())
    if len(digits) >= 6:
        return f"{int(digits[4:6]):02d}-{int(digits[:4]):04d}"
    return f"{int(default_month):02d}-{int(default_year):04d}"


def _parse_decimal(value: str) -> Decimal:
    raw = str(value or "").strip()
    if not raw:
        return Decimal("0")

    if "," in raw and "." in raw:
        normalized = raw.replace(".", "").replace(",", ".")
    elif "," in raw:
        normalized = raw.replace(",", ".")
    else:
        normalized = raw

    try:
        return Decimal(normalized)
    except (InvalidOperation, ValueError):
        return Decimal("0")


def _first_text_for_paths(root: ET.Element, paths: list[tuple[str, ...]]) -> str:
    for path in paths:
        node = _find_first_by_path(root, path)
        if node is not None and (node.text or "").strip():
            return (node.text or "").strip()
    return ""


def _find_first_by_path(root: ET.Element, path: tuple[str, ...]) -> ET.Element | None:
    if not path:
        return None

    for start in root.iter():
        if _local_name(start.tag) != path[0]:
            continue

        current = start
        found = True
        for part in path[1:]:
            current = _first_child(current, part)
            if current is None:
                found = False
                break
        if found:
            return current
    return None


def _first_child(element: ET.Element, local_name: str) -> ET.Element | None:
    for child in list(element):
        if _local_name(child.tag) == local_name:
            return child
    return None


def _find_first_by_local_name(root: ET.Element, name: str) -> ET.Element | None:
    for element in root.iter():
        if _local_name(element.tag) == name:
            return element
    return None


def _local_name(tag: str) -> str:
    if "}" in str(tag):
        return str(tag).rsplit("}", 1)[-1]
    return str(tag)


def _parse_date_sort(value: str) -> tuple[int, int, int]:
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(value, fmt)
            return dt.year, dt.month, dt.day
        except ValueError:
            continue
    return 0, 0, 0


def _parse_int(value: str) -> int:
    digits = "".join(char for char in str(value or "") if char.isdigit())
    if not digits:
        return 0
    try:
        return int(digits)
    except ValueError:
        return 0
