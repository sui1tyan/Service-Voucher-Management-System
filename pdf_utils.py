"""Service voucher PDF generation."""

from __future__ import annotations

import html
import os
import re
from pathlib import Path
from typing import Any, Mapping

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import KeepInFrame, Paragraph

from config import LOGO_PATH, PDF_DIR, SHOP_ADDR, SHOP_NAME, SHOP_TEL, logger


_styles = getSampleStyleSheet()
_body_style = ParagraphStyle(
    "VoucherBody",
    parent=_styles["Normal"],
    fontName="Helvetica",
    fontSize=9,
    leading=11,
)


def _safe_text(value: Any) -> str:
    return html.escape(str(value or "-")).replace("\n", "<br/>")


def _safe_filename(value: Any) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", str(value or "unknown"))
    return cleaned[:80] or "unknown"


def _draw_box(
    pdf: canvas.Canvas,
    *,
    x: float,
    y: float,
    width: float,
    height: float,
    title: str,
    value: Any,
) -> None:
    pdf.rect(x, y, width, height)
    pdf.setFont("Helvetica-Bold", 8)
    pdf.drawString(x + 2 * mm, y + height - 5 * mm, title)
    paragraph = Paragraph(_safe_text(value), _body_style)
    frame = KeepInFrame(
        width - 4 * mm,
        height - 9 * mm,
        [paragraph],
        mode="shrink",
        mergeSpace=True,
    )
    frame.wrapOn(pdf, width - 4 * mm, height - 9 * mm)
    frame.drawOn(pdf, x + 2 * mm, y + 2 * mm)


def _money(value: Any) -> str:
    try:
        return f"RM {float(value or 0):,.2f}"
    except (TypeError, ValueError):
        return "RM 0.00"


def generate_voucher_pdf(
    voucher: Mapping[str, Any], output_dir: str | Path | None = None
) -> str:
    """Generate an atomic, printable PDF for a complete voucher record."""

    directory = Path(output_dir or PDF_DIR)
    directory.mkdir(parents=True, exist_ok=True)
    voucher_id = _safe_filename(voucher.get("voucher_id"))
    final_path = directory / f"voucher_{voucher_id}.pdf"
    temp_path = directory / f".voucher_{voucher_id}.{os.getpid()}.part"

    try:
        pdf = canvas.Canvas(str(temp_path), pagesize=A4)
        page_width, page_height = A4
        left = 12 * mm
        right = page_width - 12 * mm
        usable_width = right - left
        top = page_height - 13 * mm

        pdf.setTitle(f"Service Voucher {voucher_id}")
        pdf.setAuthor(SHOP_NAME)
        pdf.setSubject("Service voucher")

        if Path(LOGO_PATH).is_file():
            try:
                pdf.drawImage(
                    str(LOGO_PATH),
                    right - 30 * mm,
                    top - 18 * mm,
                    30 * mm,
                    18 * mm,
                    preserveAspectRatio=True,
                    anchor="c",
                    mask="auto",
                )
            except Exception:
                logger.warning("Could not draw logo in voucher PDF", exc_info=True)

        pdf.setFont("Helvetica-Bold", 16)
        pdf.drawString(left, top, SHOP_NAME)
        pdf.setFont("Helvetica", 8.5)
        pdf.drawString(left, top - 5 * mm, SHOP_ADDR)
        pdf.drawString(left, top - 9 * mm, SHOP_TEL)

        pdf.setLineWidth(1)
        pdf.line(left, top - 21 * mm, right, top - 21 * mm)
        pdf.setFont("Helvetica-Bold", 14)
        pdf.drawCentredString(page_width / 2, top - 28 * mm, "SERVICE VOUCHER")
        pdf.setFont("Helvetica-Bold", 11)
        pdf.drawRightString(right, top - 28 * mm, f"No: {voucher_id}")

        y = top - 38 * mm
        pdf.setFont("Helvetica", 9.5)
        pdf.drawString(left, y, f"Customer: {voucher.get('customer_name') or '-'}")
        pdf.drawRightString(
            right, y, f"Date: {str(voucher.get('created_at') or '')[:10] or '-'}"
        )
        y -= 6 * mm
        pdf.drawString(left, y, f"Contact: {voucher.get('contact_number') or '-'}")
        pdf.drawRightString(right, y, f"Units: {voucher.get('units') or 1}")

        y -= 8 * mm
        half = usable_width / 2
        box_height = 42 * mm
        _draw_box(
            pdf,
            x=left,
            y=y - box_height,
            width=half,
            height=box_height,
            title="PARTICULARS",
            value=voucher.get("particulars"),
        )
        _draw_box(
            pdf,
            x=left + half,
            y=y - box_height,
            width=half,
            height=box_height,
            title="REPORTED PROBLEM",
            value=voucher.get("problem"),
        )

        y -= box_height + 7 * mm
        pdf.setFont("Helvetica", 9)
        pdf.drawString(left, y, f"Status: {voucher.get('status') or 'Pending'}")
        pdf.drawString(left + 55 * mm, y, f"Recipient: {voucher.get('recipient') or '-'}")
        y -= 6 * mm
        technician = voucher.get("technician_name") or "-"
        technician_id = voucher.get("technician_id") or "-"
        pdf.drawString(left, y, f"Technician: {technician} ({technician_id})")

        y -= 8 * mm
        solution_height = 32 * mm
        _draw_box(
            pdf,
            x=left,
            y=y - solution_height,
            width=usable_width,
            height=solution_height,
            title="SOLUTION / WORK PERFORMED",
            value=voucher.get("solution"),
        )

        y -= solution_height + 8 * mm
        pdf.setFont("Helvetica", 9)
        bill = voucher.get("ref_bill") or "-"
        bill_date = voucher.get("ref_bill_date") or "-"
        pdf.drawString(left, y, f"Reference bill: {bill}")
        pdf.drawString(left + 70 * mm, y, f"Bill date: {bill_date}")
        pdf.drawRightString(right, y, f"Amount: {_money(voucher.get('amount_rm'))}")
        y -= 6 * mm
        pdf.drawRightString(
            right,
            y,
            f"Technician commission: {_money(voucher.get('tech_commission'))}",
        )

        signature_y = 28 * mm
        pdf.line(left, signature_y, left + 62 * mm, signature_y)
        pdf.line(right - 62 * mm, signature_y, right, signature_y)
        pdf.setFont("Helvetica", 8)
        pdf.drawString(left, signature_y - 4 * mm, "Recipient signature")
        pdf.drawString(right - 62 * mm, signature_y - 4 * mm, "Customer signature")

        pdf.setFont("Helvetica-Oblique", 7)
        pdf.drawCentredString(
            page_width / 2,
            10 * mm,
            "Please retain this voucher for collection and service enquiries.",
        )

        pdf.showPage()
        pdf.save()
        os.replace(temp_path, final_path)
        return str(final_path.resolve())
    except Exception:
        if temp_path.exists():
            temp_path.unlink(missing_ok=True)
        logger.exception("PDF generation failed for voucher %s", voucher_id)
        raise


def generate_pdf(
    voucher_id: str,
    customer_name: str,
    contact_number: str,
    units: int,
    particulars: str,
    problem: str,
    staff_name: str,
    status: str,
    created_at: str,
    recipient: str,
) -> str:
    """Compatibility wrapper for callers from older modular revisions."""

    return generate_voucher_pdf(
        {
            "voucher_id": voucher_id,
            "customer_name": customer_name,
            "contact_number": contact_number,
            "units": units,
            "particulars": particulars,
            "problem": problem,
            "staff_name": staff_name,
            "status": status,
            "created_at": created_at,
            "recipient": recipient,
        }
    )
