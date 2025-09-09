def _draw_voucher(c, width, height, voucher_id, customer_name, contact_number,
                  units, remark, particulars, problem, staff_name, created_at):
    """
    Draws a SINGLE voucher on A4 (no second copy). 
    Uses draw_wrapped() for fields to prevent overlapping.
    """
    left   = 12*mm
    right  = width - 12*mm
    top_y  = height - 15*mm
    y      = top_y
    lh     = 4.2*mm
    row_h  = 10.5*mm

    # --- Header ---
    if LOGO_PATH and os.path.exists(LOGO_PATH):
        try:
            c.drawImage(LOGO_PATH, left, y-12*mm, width=18*mm, height=18*mm,
                        preserveAspectRatio=True, mask='auto')
            text_x = left + 20*mm
        except Exception:
            text_x = left
    else:
        text_x = left

    c.setFont("Helvetica-Bold", 14)
    c.drawString(text_x, y, SHOP_NAME)
    c.setFont("Helvetica", 9.2)
    c.drawString(text_x, y - 5.0*mm, SHOP_ADDR)
    c.drawString(text_x, y - 9.0*mm, SHOP_TEL)

    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString((left+right)/2, y - 16.0*mm, "SERVICE VOUCHER")
    c.drawRightString(right, y - 16.0*mm, f"No : {voucher_id}")

    # Date / Time
    c.setFont("Helvetica", 10)
    c.drawString(left, y - 24.0*mm, "Date :")
    c.drawString(left + 18*mm, y - 24.0*mm, created_at[:10])
    c.drawRightString(right - 27*mm, y - 24.0*mm, "Time In :")
    c.drawRightString(right, y - 24.0*mm, created_at[11:19])

    # --- Table ---
    top_table = y - 28.0*mm
    c.rect(left, top_table - row_h*3, right-left, row_h*3, stroke=1, fill=0)

    name_col_x = left + 58*mm
    qty_col_x  = right - 20*mm
    c.line(name_col_x, top_table, name_col_x, top_table - row_h*3)
    c.line(qty_col_x,  top_table, qty_col_x,  top_table - row_h*3)
    c.line(left, top_table - row_h,  right, top_table - row_h)
    c.line(left, top_table - row_h*2,right, top_table - row_h*2)

    # headers
    c.setFont("Helvetica-Bold", 10.4)
    c.drawString(left+2*mm, top_table - row_h + lh, "CUSTOMER NAME")
    c.drawCentredString((name_col_x+qty_col_x)/2, top_table - row_h + lh, "PARTICULARS")
    c.drawCentredString((qty_col_x+right)/2, top_table - row_h + lh, "QTY")
    c.drawString(left+2*mm, top_table - row_h*2 + lh, "TEL:")

    # values (wrapped)
    draw_wrapped(c, customer_name,
                 left+42*mm, top_table - row_h + 2,
                 w=(name_col_x - (left+42*mm) - 4), h=row_h, fontsize=10)

    draw_wrapped(c, contact_number,
                 left+22*mm, top_table - row_h*2 + 2,
                 w=(name_col_x - (left+22*mm) - 4), h=row_h, fontsize=10)

    draw_wrapped(c, particulars or "-",
                 name_col_x + 3*mm, top_table - row_h*2 + 2,
                 w=(qty_col_x - (name_col_x+3*mm) - 4), h=row_h, fontsize=10)

    c.setFont("Helvetica", 10)
    c.drawCentredString((qty_col_x+right)/2, top_table - row_h*2 + lh, str(units or 1))

    # problem
    y2 = top_table - row_h*3 - 12
    c.setFont("Helvetica-Bold", 10.4); c.drawString(left, y2, "PROBLEM:")
    draw_wrapped(c, problem or "-", left+26*mm, y2,
                 w=(right - (left+26*mm) - 4), h=row_h, fontsize=10)

    # recipient / staff
    y3 = y2 - 10.0*mm
    c.setFont("Helvetica-Bold", 10.4); c.drawString(left, y3, "RECIPIENT :")
    c.setFont("Helvetica-Bold", 10.4); c.drawString(left+70*mm, y3, "NAME OF STAFF :")
    draw_wrapped(c, staff_name or "-", left+110*mm, y3,
                 w=(right - (left+110*mm) - 4), h=row_h, fontsize=10)

    # remark
    y4 = y3 - 10.0*mm
    c.setFont("Helvetica-Bold", 10.4); c.drawString(left, y4, "REMARK:")
    draw_wrapped(c, remark or "-", left+22*mm, y4,
                 w=(right - (left+22*mm) - 4), h=row_h, fontsize=10)

    # signatures
    y5 = y4 - 12*mm
    c.line(left, y5, left+60*mm, y5)
    c.setFont("Helvetica", 9); c.drawString(left, y5 - 4*mm, "CUSTOMER SIGNATURE")
    c.line(right-60*mm, y5, right, y5)
    c.drawString(right-60*mm, y5 - 4*mm, "DATE COLLECTED")

    # warnings
    y6 = y5 - 8.5*mm
    c.setFont("Helvetica", 8.5)
    c.drawString(left, y6, "* Kindly collect your goods within 60 days of sending for repair.")
    c.drawString(left, y6 - 4*mm, "A) We do not hold ourselves responsible for any loss or damage.")
    c.drawString(left, y6 - 8*mm, "B) We reserve our right to sell off the goods to cover our cost and loss.")
    c.drawString(left, y6 - 12*mm, "* MINIMUM RM45.00 WILL BE CHARGED ON TROUBLESHOOTING / INSPECTION / SERVICE.")

    # QR code
    try:
        qr_size = 18*mm
        qr_data = f"Voucher:{voucher_id}|Name:{customer_name}|Tel:{contact_number}|Date:{created_at[:10]}"
        qr_img  = qrcode.make(qr_data)
        qr_path = os.path.join(PDF_DIR, f"qr_{voucher_id}.png")
        qr_img.save(qr_path)
        c.drawImage(qr_path, right - qr_size, y6 - 16*mm, qr_size, qr_size)
        os.remove(qr_path)
    except Exception:
        pass
