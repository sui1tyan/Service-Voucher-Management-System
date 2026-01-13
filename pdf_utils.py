import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as rl_canvas 
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime
from config import PDF_DIR, SHOP_NAME, SHOP_ADDR, SHOP_TEL, LOGO_PATH, logger

_styles = getSampleStyleSheet()
_styleN = _styles["Normal"]

def draw_wrapped(c, text, x, y, w, h, fontsize=10, bold=False, leading=None):
    style = _styleN.clone('wrap')
    style.fontName = "Helvetica-Bold" if bold else "Helvetica"
    style.fontSize = fontsize
    style.leading = leading if leading else fontsize + 2
    para = Paragraph((text or "-").replace("\n", "<br/>"), style)
    _, h_used = para.wrap(w, h)
    para.drawOn(c, x, y + h - h_used)
    return h_used

def draw_wrapped_top(c, text, x, top_y, w, fontsize=10, bold=False, leading=None):
    style = _styleN.clone('wrapTop')
    style.fontName = "Helvetica-Bold" if bold else "Helvetica"
    style.fontSize = fontsize
    style.leading = leading if leading else fontsize + 2
    para = Paragraph((text or "-").replace("\n", "<br/>"), style)
    _, h_used = para.wrap(w, 1000 * mm)
    para.drawOn(c, x, top_y - h_used)
    return h_used

def generate_pdf(voucher_id, customer_name, contact_number, units,
                 particulars, problem, staff_name, status, created_at, recipient):
    os.makedirs(PDF_DIR, exist_ok=True)
    final_pdf = os.path.join(PDF_DIR, f"voucher_{voucher_id}.pdf")
    tmp_pdf = final_pdf + ".part"
    
    try:
        c = rl_canvas.Canvas(tmp_pdf, pagesize=A4)
        width, height = A4
        left, right, top_y = 12 * mm, width - 12 * mm, height - 15 * mm
        
        # Header
        y = top_y
        LOGO_W = 28 * mm; LOGO_H = 18 * mm
        if LOGO_PATH and os.path.exists(LOGO_PATH):
            try:
                c.drawImage(LOGO_PATH, right - LOGO_W, y - LOGO_H, LOGO_W, LOGO_H, mask='auto')
            except: pass
        c.setFont("Helvetica-Bold", 14)
        c.drawString(left, y, SHOP_NAME)
        c.setFont("Helvetica", 9.2)
        c.drawString(left, y - 5.0 * mm, SHOP_ADDR)
        c.drawString(left, y - 9.0 * mm, SHOP_TEL)
        c.setFont("Helvetica-Bold", 13)
        c.drawCentredString((left + right) / 2, y - 16.0 * mm, "SERVICE VOUCHER")
        c.drawRightString(right, y - 16.0 * mm, f"No : {voucher_id}")
        
        # Date
        base_y = y - 16.0 * mm
        try:
            dt = datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S")
            date_str, time_str = dt.strftime("%d-%m-%Y"), dt.strftime("%H:%M:%S")
        except:
            date_str, time_str = created_at[:10], created_at[11:19]
        c.setFont("Helvetica", 10)
        c.drawString(left, base_y - 8.0 * mm, "Date :")
        c.drawString(left + 18 * mm, base_y - 8.0 * mm, date_str)
        c.drawRightString(right - 27 * mm, base_y - 8.0 * mm, "Time In :")
        c.drawRightString(right, base_y - 8.0 * mm, time_str)

        # Main Table
        top_table = base_y - 12 * mm
        qty_col_w = 20 * mm
        left_col_w = 74 * mm
        name_col_x = left + left_col_w
        qty_col_x = right - qty_col_w
        row1_h = 20 * mm
        row2_h = 20 * mm
        mid_y = top_table - row1_h
        bottom_table = top_table - (row1_h + row2_h)
        pad = 3 * mm

        c.rect(left, bottom_table, right - left, (row1_h + row2_h), stroke=1, fill=0)
        c.line(name_col_x, top_table, name_col_x, bottom_table)
        c.line(qty_col_x, top_table, qty_col_x, mid_y)
        c.line(left, mid_y, right, mid_y)

        c.setFont("Helvetica-Bold", 10.4)
        c.drawString(left + pad, top_table - pad - 8, "CUSTOMER NAME")
        draw_wrapped(c, customer_name, left + pad, mid_y + pad, left_col_w - 2 * pad, row1_h - 2 * pad - 10)
        
        c.drawString(name_col_x + pad, top_table - pad - 8, "PARTICULARS")
        draw_wrapped(c, particulars, name_col_x + pad, mid_y + pad, (right - left) - left_col_w - qty_col_w - 2 * pad, row1_h - 2 * pad - 10)

        c.drawCentredString(qty_col_x + qty_col_w / 2, top_table - pad - 8, "QTY")
        c.setFont("Helvetica", 11)
        c.drawCentredString(qty_col_x + qty_col_w / 2, mid_y + (row1_h / 2) - 3, str(units))

        c.setFont("Helvetica-Bold", 10.4)
        c.drawString(left + pad, mid_y - pad - 8, "TEL")
        draw_wrapped(c, contact_number, left + pad, bottom_table + pad, left_col_w - 2 * pad, row2_h - 2 * pad - 10)

        c.drawString(name_col_x + pad, mid_y - pad - 8, "PROBLEM")
        draw_wrapped(c, problem, name_col_x + pad, bottom_table + pad, (right - left) - left_col_w - 2 * pad, row2_h - 2 * pad - 10)

        # Footer / Signatures
        y_rec = bottom_table - 9 * mm
        c.setFont("Helvetica-Bold", 10.4)
        c.drawString(left, y_rec, "RECIPIENT :")
        c.setFont("Helvetica", 9)
        c.drawString(left + 30 * mm, y_rec, recipient or "")
        
        # Terms
        policies_top = y_rec - 7 * mm
        draw_wrapped_top(c, "Kindly collect your goods within <font color='red'>60 days</font>.", left, policies_top, left_col_w, fontsize=7)
        
        # Signatures
        y_sig = max(10 * mm, (A4[1] / 2) - 20 * mm)
        c.line(right - 50*mm, y_sig, right, y_sig)
        c.drawString(right - 50*mm, y_sig - 4*mm, "CUSTOMER SIGNATURE")

        c.showPage()
        c.save()
        
        if os.path.exists(final_pdf): os.remove(final_pdf)
        os.rename(tmp_pdf, final_pdf)
        return final_pdf

    except Exception as e:
        logger.exception("PDF Gen failed")
        if os.path.exists(tmp_pdf): os.remove(tmp_pdf)
        raise
