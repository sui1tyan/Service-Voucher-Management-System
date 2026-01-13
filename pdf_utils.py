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

def draw_wrapped(c, text, x, y, w, h, fontsize=10, bold=False):
    style = _styleN.clone('wrap')
    style.fontName = "Helvetica-Bold" if bold else "Helvetica"
    style.fontSize = fontsize
    style.leading = fontsize + 2
    para = Paragraph((text or "-").replace("\n", "<br/>"), style)
    _, h_used = para.wrap(w, h)
    para.drawOn(c, x, y + h - h_used)
    return h_used

def generate_pdf(voucher_id, customer_name, contact_number, units,
                 particulars, problem, staff_name, status, created_at, recipient):
    os.makedirs(PDF_DIR, exist_ok=True)
    final_pdf = os.path.join(PDF_DIR, f"voucher_{voucher_id}.pdf")
    tmp_pdf = final_pdf + ".part"
    
    try:
        c = rl_canvas.Canvas(tmp_pdf, pagesize=A4)
        width, height = A4
        left, right = 12 * mm, width - 12 * mm
        top_y = height - 15 * mm
        
        # Header
        if LOGO_PATH and os.path.exists(LOGO_PATH):
            try:
                c.drawImage(LOGO_PATH, right - 28*mm, top_y - 18*mm, 28*mm, 18*mm, mask='auto')
            except: pass
            
        c.setFont("Helvetica-Bold", 14)
        c.drawString(left, top_y, SHOP_NAME)
        c.setFont("Helvetica", 9)
        c.drawString(left, top_y - 5*mm, SHOP_ADDR)
        c.drawString(left, top_y - 9*mm, SHOP_TEL)
        
        c.setFont("Helvetica-Bold", 13)
        c.drawCentredString(width/2, top_y - 16*mm, "SERVICE VOUCHER")
        c.drawRightString(right, top_y - 16*mm, f"No : {voucher_id}")
        
        # Details
        y = top_y - 30*mm
        c.setFont("Helvetica", 10)
        c.drawString(left, y, f"Customer: {customer_name}")
        c.drawRightString(right, y, f"Date: {created_at[:10]}")
        
        y -= 8*mm
        c.drawString(left, y, f"Contact: {contact_number}")
        
        y -= 15*mm
        c.rect(left, y - 40*mm, right-left, 45*mm)
        c.drawString(left + 2*mm, y, "Particulars:")
        draw_wrapped(c, particulars, left + 2*mm, y - 18*mm, (right-left)/2, 15*mm)
        
        c.drawString(left + (right-left)/2 + 2*mm, y, "Problem:")
        draw_wrapped(c, problem, left + (right-left)/2 + 2*mm, y - 18*mm, (right-left)/2 - 4*mm, 15*mm)

        # Footer
        y_sig = 30 * mm
        c.line(left, y_sig, left + 60*mm, y_sig)
        c.drawString(left, y_sig - 4*mm, "Recipient Signature")
        
        c.line(right - 60*mm, y_sig, right, y_sig)
        c.drawString(right - 60*mm, y_sig - 4*mm, "Customer Signature")

        c.showPage()
        c.save()
        
        if os.path.exists(final_pdf): os.remove(final_pdf)
        os.rename(tmp_pdf, final_pdf)
        return final_pdf
    except Exception:
        logger.exception("PDF generation failed")
        raise
