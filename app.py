import streamlit as st
import datetime
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement # For setting table cell borders
import io # To handle file in memory

# --- Helper function to set table cell borders ---
def set_cell_border(cell, **kwargs):
    """
    Set cell border properties.
    Args:
        cell: The docx.table._Cell object.
        kwargs: Keyword arguments for borders (top, bottom, left, right, start, end, insideH, insideV).
                Each value should be a dict with 'sz' (size in eighths of a point) and 'color' (RGBColor).
    Example: set_cell_border(cell, top={'sz': 12, 'color': RGBColor(0, 0, 0)})
    """
    tcPr = cell._element.get_or_add_tcPr()

    # Define border properties in shorthand for convenience
    borders = {
        'top': {'tag': 'w:topBdr', 'val': 'single'},
        'bottom': {'tag': 'w:bottomBdr', 'val': 'single'},
        'left': {'tag': 'w:leftBdr', 'val': 'single'},
        'right': {'tag': 'w:rightBdr', 'val': 'single'},
        'insideH': {'tag': 'w:insideH'}, # Horizontal internal borders for rows
        'insideV': {'tag': 'w:insideV'}, # Vertical internal borders for columns
    }

    for border_name, default_props in borders.items():
        if border_name in kwargs:
            bdr = OxmlElement(default_props['tag'])
            bdr.set(qn('w:val'), kwargs[border_name].get('val', default_props.get('val', 'single')))
            bdr.set(qn('w:sz'), str(kwargs[border_name].get('sz', 12))) # Default 12 means 1.5pt (12/8)
            color_val = kwargs[border_name].get('color', RGBColor(0, 0, 0)) # Default black
            bdr.set(qn('w:color'), f'{color_val.rgb[0]:02X}{color_val.rgb[1]:02X}{color_val.rgb[2]:02X}')
            tcPr.append(bdr)
        # If no specific border is given, ensure default borders are removed if not explicitly set by table style
        else:
            # This part is tricky. If you want *no* border, you need to set val='nil'
            # But the default table style might already have borders.
            # For simplicity, we'll only *add* borders if specified.
            pass


# --- Function to generate DOCX ---
def generate_invoice_docx(invoice_data):
    document = Document()

    # Set default document margins (adjust if your template has different margins)
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(1.5) # Approx 1.5cm or 0.6 inches
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)

    # --- Header Section: "TAX INVOICE/DELIVERY NOTE" ---
    title = document.add_paragraph('TAX INVOICE/DELIVERY NOTE')
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title.runs[0]
    run_title.font.name = 'Arial' # Assuming Arial or similar sans-serif font
    run_title.bold = True
    run_title.font.size = Pt(18) # Larger font for title

    document.add_paragraph() # Add a blank line for spacing

    # --- Table for Contact and Invoice Details ---
    # This section appears to be a two-column layout visually, with some fields left blank.
    # We will create a 4-column table for flexibility and merge cells where necessary to match visual.
    # No visible borders are typical for this section.
    
    # Create table with 4 columns
    header_info_table = document.add_table(rows=6, cols=4)
    header_info_table.autofit = False
    header_info_table.allow_autofit = False # Disable autofit to control widths manually
    
    # Set preferred table style (e.g., no borders by default if you want to add them selectively)
    # This might require creating a custom style in the document template or manually clearing all borders.
    # For now, we'll ensure no explicit borders are set by python-docx on cells.

    # Approximate column widths for the header table (adjust these precisely)
    header_info_table.columns[0].width = Cm(2.5) # Label column (To:, Address:, etc.)
    header_info_table.columns[1].width = Cm(7.0) # Value column (Company Name, Address)
    header_info_table.columns[2].width = Cm(3.0) # Second Label column (Date:, SSS Invoice No:)
    header_info_table.columns[3].width = Cm(6.5) # Second Value column (Date, Invoice No)

    # Row 1: To: / Date:
    cell_row1_col0 = header_info_table.rows[0].cells[0]
    cell_row1_col0.text = "To:"
    cell_row1_col1 = header_info_table.rows[0].cells[1]
    cell_row1_col1.text = invoice_data['to_company']
    cell_row1_col2 = header_info_table.rows[0].cells[2]
    cell_row1_col2.text = "Date:"
    cell_row1_col3 = header_info_table.rows[0].cells[3]
    cell_row1_col3.text = invoice_data['invoice_date']

    # Row 2: Address: / SSS Invoice No:
    cell_row2_col0 = header_info_table.rows[1].cells[0]
    cell_row2_col0.text = "Address:"
    cell_row2_col1 = header_info_table.rows[1].cells[1]
    cell_row2_col1.text = invoice_data['customer_address']
    cell_row2_col2 = header_info_table.rows[1].cells[2]
    cell_row2_col2.text = "SSS Invoice No:"
    cell_row2_col3 = header_info_table.rows[1].cells[3]
    cell_row2_col3.text = invoice_data['sss_invoice_no']

    # Row 3: Tel: / Customer VAT No.:
    cell_row3_col0 = header_info_table.rows[2].cells[0]
    cell_row3_col0.text = "Tel:"
    cell_row3_col1 = header_info_table.rows[2].cells[1]
    cell_row3_col1.text = invoice_data['customer_tel']
    cell_row3_col2 = header_info_table.rows[2].cells[2]
    cell_row3_col2.text = "Customer VAT No.:"
    cell_row3_col3 = header_info_table.rows[2].cells[3]
    cell_row3_col3.text = invoice_data['customer_vat_no']

    # Row 4: ATTN: (Spans 2 columns on the right)
    cell_row4_col0 = header_info_table.rows[3].cells[0]
    cell_row4_col0.text = "ATTN:"
    cell_row4_col1 = header_info_table.rows[3].cells[1]
    cell_row4_col1.text = invoice_data['attn_person']
    # Merge cells for empty space on the right for ATTN line
    header_info_table.rows[3].cells[2].merge(header_info_table.rows[3].cells[3])

    # Row 5: Email: (Spans 2 columns on the right)
    cell_row5_col0 = header_info_table.rows[4].cells[0]
    cell_row5_col0.text = "Email:"
    cell_row5_col1 = header_info_table.rows[4].cells[1]
    cell_row5_col1.text = invoice_data['customer_email']
    # Merge cells for empty space on the right for Email line
    header_info_table.rows[4].cells[2].merge(header_info_table.rows[4].cells[3])

    # Row 6: Customer PO#: (Spans 2 columns on the right)
    cell_row6_col0 = header_info_table.rows[5].cells[0]
    cell_row6_col0.text = "Customer PO#:"
    cell_row6_col1 = header_info_table.rows[5].cells[1]
    cell_row6_col1.text = invoice_data['customer_po']
    # Merge cells for empty space on the right for Customer PO# line
    header_info_table.rows[5].cells[2].merge(header_info_table.rows[5].cells[3])

    # Apply font formatting to all cells in the header info table
    for row in header_info_table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Arial' # Consistent font
                    run.font.size = Pt(10) # Consistent font size
                    # No borders for this table
                    set_cell_border(cell, top={'sz': 0, 'val': 'nil'}, bottom={'sz': 0, 'val': 'nil'},
                                    left={'sz': 0, 'val': 'nil'}, right={'sz': 0, 'val': 'nil'})


    document.add_paragraph() # Spacing after header table

    # --- Line Items Table ---
    item_table = document.add_table(rows=1, cols=7)
    item_table.autofit = False
    item_table.allow_autofit = False

    # Set explicit column widths for line items table (crucial for exact layout)
    # These widths are estimations; fine-tune them against your original DOCX
    item_table.columns[0].width = Cm(1.0)  # No
    item_table.columns[1].width = Cm(8.0)  # Description
    item_table.columns[2].width = Cm(2.5)  # Unit price
    item_table.columns[3].width = Cm(1.5)  # BHD (Unit)
    item_table.columns[4].width = Cm(1.5)  # Qty
    item_table.columns[5].width = Cm(2.5)  # Total price
    item_table.columns[6].width = Cm(1.5)  # BHD (Total)

    # Set default table style to include all borders for line items
    # You might need to adjust table.style to 'Table Grid' or manually add borders
    item_table.style = 'Table Grid' # This usually adds all standard borders

    # Add header row
    item_table_headers = ["No", "Description", "Unit price", "BHD", "Qty", "Total price", "BHD"]
    hdr_cells = item_table.rows[0].cells
    for i, header_text in enumerate(item_table_headers):
        cell = hdr_cells[i]
        cell.text = header_text
        paragraph = cell.paragraphs[0]
        # Align headers
        if header_text in ["No", "Qty", "BHD"]:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif header_text in ["Unit price", "Total price"]:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT # Right align price headers
        else:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # Apply bold, font, and size for headers
        for run in paragraph.runs:
            run.font.name = 'Arial'
            run.bold = True
            run.font.size = Pt(9)
        # Apply light gray shading to header row cells
        cell.shading.background = RGBColor(0xD9, 0xD9, 0xD9) # Light gray (hex #D9D9D9)

    # Add item rows
    for i, item in enumerate(invoice_data['line_items']):
        row_cells = item_table.add_row().cells
        # No.
        cell_no = row_cells[0]
        cell_no.text = str(i + 1)
        cell_no.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Description
        cell_desc = row_cells[1]
        cell_desc.text = item['description']
        cell_desc.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        # Unit price
        cell_unit_price = row_cells[2]
        cell_unit_price.text = f"{item['unit_price']:.3f}"
        cell_unit_price.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        # BHD (Unit)
        cell_bhd_unit = row_cells[3]
        cell_bhd_unit.text = "BHD"
        cell_bhd_unit.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Qty
        cell_qty = row_cells[4]
        cell_qty.text = str(item['quantity'])
        cell_qty.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Total price
        cell_total_price = row_cells[5]
        cell_total_price.text = f"{item['total_price']:.3f}"
        cell_total_price.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        # BHD (Total)
        cell_bhd_total = row_cells[6]
        cell_bhd_total.text = "BHD"
        cell_bhd_total.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Apply font size for all cells in item rows
        for j in range(7):
            for paragraph in row_cells[j].paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(9)

    # Add empty rows if needed to match min rows in template, otherwise dynamic.
    # For now, we'll assume dynamic based on input.

    # --- Summary Section (Subtotal, VAT, Grand Total) ---
    # This section appears to be a 2-column table that aligns with the right side of the main table.
    summary_table = document.add_table(rows=3, cols=2)
    summary_table.autofit = False
    summary_table.allow_autofit = False

    # Adjust summary table position using left indent for table (requires more advanced XML manipulation)
    # Or, the simpler way is to align its left edge with the 'Unit Price' column of the main table
    # by setting its first column to be very wide (approx. item_table.columns[0] to item_table.columns[4] widths summed)
    # Total width of item table is approx Cm(1.0 + 8.0 + 2.5 + 1.5 + 1.5 + 2.5 + 1.5) = 18.5 cm
    # Summary starts visually after description, maybe 8.0 + 1.0 + 2.5 = 11.5cm from left edge?
    # Let's align its right edge.
    # Total document width (after margins) approx 21cm - 3cm = 18cm.
    # Right column widths for prices are 2.5 + 1.5 = 4cm.
    # So left column width for summary table = (18cm - 4cm) = 14cm approx.

    summary_table.columns[0].width = Cm(14.0) # Label column (Subtotal:, VAT @ 10%:, Grand Total:)
    summary_table.columns[1].width = Cm(4.0)  # Value column (Amounts)

    # Remove all borders for summary table cells, as per typical invoice design.
    for row in summary_table.rows:
        for cell in row.cells:
            set_cell_border(cell, top={'sz': 0, 'val': 'nil'}, bottom={'sz': 0, 'val': 'nil'},
                                left={'sz': 0, 'val': 'nil'}, right={'sz': 0, 'val': 'nil'})

    # Row 1: Subtotal
    subtotal_cell_label = summary_table.rows[0].cells[0]
    subtotal_cell_label.text = "Subtotal:"
    subtotal_cell_label.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    subtotal_cell_value = summary_table.rows[0].cells[1]
    subtotal_cell_value.text = f"{invoice_data['subtotal']:.3f} BHD"
    subtotal_cell_value.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Row 2: VAT
    vat_cell_label = summary_table.rows[1].cells[0]
    vat_cell_label.text = "VAT @ 10%:"
    vat_cell_label.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    vat_cell_value = summary_table.rows[1].cells[1]
    vat_cell_value.text = f"{invoice_data['vat_amount']:.3f} BHD"
    vat_cell_value.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Row 3: Grand Total
    grand_total_cell_label = summary_table.rows[2].cells[0]
    grand_total_cell_label.text = "Grand Total in BHD"
    grand_total_cell_label.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    grand_total_cell_value = summary_table.rows[2].cells[1]
    grand_total_cell_value.text = f"{invoice_data['grand_total']:.3f} BHD"
    grand_total_cell_value.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Apply bold and potentially larger font size for Grand Total
    for p in grand_total_cell_label.paragraphs:
        for run in p.runs:
            run.font.name = 'Arial'
            run.bold = True
            run.font.size = Pt(11) # Slightly larger

    for p in grand_total_cell_value.paragraphs:
        for run in p.runs:
            run.font.name = 'Arial'
            run.bold = True
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0, 0, 0) # Black color for consistency

    document.add_paragraph() # Add some space after summary

    # --- Payment Details (Static Text) ---
    payment_details_para = document.add_paragraph()
    payment_details_para.add_run("Payment will be remitted to the following account:").font.name = 'Arial'
    payment_details_para.add_run("\nSalahuddin Softtech Solutions, Arab Bank, A/C No.: 2002-146072-510,").font.name = 'Arial'
    payment_details_para.add_run("\nIBAN: BH31 ARAB 0200 2146072510, Swift Code: ARAB BHBM,").font.name = 'Arial'
    payment_details_para.add_run("\nAddress: Arab Bank Plc. , P.O Box: 395, Manama, Kingdom of Bahrain.").font.name = 'Arial'
    for run in payment_details_para.runs:
        run.font.size = Pt(9)

    document.add_paragraph() # Spacing

    # --- Terms and Conditions ---
    terms_heading = document.add_paragraph('Terms and Conditions')
    terms_heading.runs[0].bold = True
    terms_heading.runs[0].font.size = Pt(12)
    terms_heading.runs[0].font.name = 'Arial'

    terms_para1 = document.add_paragraph()
    terms_para1.add_run("Payment terms: ").bold = True
    terms_para1.add_run(invoice_data['payment_terms'])
    for run in terms_para1.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(9)

    terms_para2 = document.add_paragraph("Title and property at all above remain ours until full payment is received and we reserve the rights to withdraw goods/services if not paid for when due.")
    for run in terms_para2.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(9)

    terms_para3 = document.add_paragraph("Signing this document implies acceptance of these terms.")
    for run in terms_para3.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(9)

    document.add_paragraph() # Spacing

    # --- Receipt/Acknowledgement ---
    # Using a table for precise alignment of (Name), (Date), (Signature)
    ack_table = document.add_table(rows=1, cols=3)
    ack_table.autofit = False
    ack_table.allow_autofit = False

    # Set widths to spread them out appropriately
    ack_table.columns[0].width = Cm(6) # Name
    ack_table.columns[1].width = Cm(6) # Date
    ack_table.columns[2].width = Cm(6) # Signature

    # No borders for this table
    for row in ack_table.rows:
        for cell in row.cells:
            set_cell_border(cell, top={'sz': 0, 'val': 'nil'}, bottom={'sz': 0, 'val': 'nil'},
                                left={'sz': 0, 'val': 'nil'}, right={'sz': 0, 'val': 'nil'})

    ack_cells = ack_table.rows[0].cells
    ack_cells[0].text = "(Name)"
    ack_cells[1].text = "(Date)"
    ack_cells[2].text = "(Signature)"

    for cell in ack_cells:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in cell.paragraphs[0].runs:
            run.font.name = 'Arial'
            run.font.size = Pt(9)
            # Add a bottom border only for the signature line underneath the text
            # This is more complex and typically done by drawing shapes or using specific paragraph borders.
            # For simplicity, if your template shows a line, you might add a bottom border to the cell,
            # or rely on the user visually interpreting the text placement.
            # A simpler way to get a line is to use an underscore _ within the text, but that's not ideal.


    document.add_paragraph() # Spacing after the ack_table

    document.add_paragraph("For SALAHUDDIN SOFTTECH SOLUTIONS")
    document.add_paragraph("Jobin George")
    document.add_paragraph("Operations Manager")
    # Apply font formatting to these last three paragraphs
    for para in document.paragraphs[-3:]: # Get the last 3 paragraphs
        for run in para.runs:
            run.font.name = 'Arial'
            run.font.size = Pt(10) # Adjust size if needed

    # Save the document to an in-memory byte stream
    byte_io = io.BytesIO()
    document.save(byte_io)
    byte_io.seek(0)
    return byte_io

# --- Streamlit Form Layout (Remains the same as previous corrected version) ---
st.title("Invoice Generator")

st.subheader("Line Items")

# Initialize line_items in session_state if not present
if 'line_items' not in st.session_state:
    st.session_state.line_items = []

# Buttons to add/remove items (OUTSIDE the form)
if st.button("Add New Item"):
    st.session_state.line_items.append({"description": "", "unit_price": 0.000, "quantity": 1, "total_price": 0.000})

total_subtotal = 0.000
items_to_remove = []

for i, item in enumerate(st.session_state.line_items):
    st.write(f"--- Item {i+1} ---")
    item_cols = st.columns([4, 2, 1, 1]) # Description, Unit Price, Qty, Remove Button
    with item_cols[0]:
        item['description'] = st.text_input("Description", item['description'], key=f"desc_{i}")
    with item_cols[1]:
        item['unit_price'] = st.number_input("Unit price (BHD)", min_value=0.000, value=float(item['unit_price']), format="%.3f", key=f"price_{i}")
    with item_cols[2]:
        item['quantity'] = st.number_input("Qty", min_value=1, value=int(item['quantity']), step=1, key=f"qty_{i}")

    item['total_price'] = item['unit_price'] * item['quantity']
    total_subtotal += item['total_price']

    st.text(f"Total Price for Item {i+1}: {item['total_price']:.3f} BHD")
    if item_cols[3].button(f"Remove Item {i+1}", key=f"remove_{i}"):
        items_to_remove.append(i)

for i in sorted(items_to_remove, reverse=True):
    del st.session_state.line_items[i]
if items_to_remove:
    st.experimental_rerun()


# --- Main Invoice Form (for header details and final submission) ---
with st.form("invoice_form"):
    st.subheader("Customer & Invoice Details")

    col1_form, col2_form = st.columns(2)
    with col1_form:
        to_company = st.text_input("To (Company Name)", "ABC Company W.L.L.")
        customer_address = st.text_area("Address", "Building 123, Road 456, Block 789, Manama, Bahrain")
        customer_tel = st.text_input("Tel", "+973 17XXXXXX")
        attn_person = st.text_input("ATTN", "Mr. John Doe")
        customer_email = st.text_input("Email", "john.doe@abccompany.com")
        customer_po = st.text_input("Customer PO#", "PO-12345")

    with col2_form:
        invoice_date = st.date_input("Date", datetime.date.today())
        current_date_str = datetime.date.today().strftime("%y%m%d")
        sss_invoice_no = st.text_input("SSS Invoice No", f"SSS-{current_date_str}-001")
        customer_vat_no = st.text_input("Customer VAT No.", "VAT123456789")

    vat_rate = 0.10
    vat_amount = total_subtotal * vat_rate
    grand_total = total_subtotal + vat_amount

    st.subheader("Summary")
    st.markdown(f"**Subtotal:** {total_subtotal:.3f} BHD")
    st.markdown(f"**VAT @ 10%:** {vat_amount:.3f} BHD")
    st.markdown(f"## **Grand Total:** {grand_total:.3f} BHD")


    st.subheader("Terms and Conditions")
    payment_terms = st.text_area("Payment terms", "30 days from invoice date")

    # --- Submit Button for the Form ---
    submitted = st.form_submit_button("Generate Invoice")

    if submitted:
        if not st.session_state.line_items:
            st.warning("Please add at least one line item to generate the invoice.")
            if 'generated_docx_data' in st.session_state:
                del st.session_state.generated_docx_data
            if 'generated_docx_filename' in st.session_state:
                del st.session_state.generated_docx_filename
        else:
            invoice_data = {
                "to_company": to_company,
                "customer_address": customer_address,
                "customer_tel": customer_tel,
                "attn_person": attn_person,
                "customer_email": customer_email,
                "customer_po": customer_po,
                "invoice_date": invoice_date.strftime("%d-%m-%Y"),
                "sss_invoice_no": sss_invoice_no,
                "customer_vat_no": customer_vat_no,
                "line_items": st.session_state.line_items,
                "subtotal": total_subtotal,
                "vat_amount": vat_amount,
                "grand_total": grand_total,
                "payment_terms": payment_terms
            }

            st.success("Invoice data collected successfully! Generating document...")

            docx_file_bytes = generate_invoice_docx(invoice_data)

            st.session_state.generated_docx_data = docx_file_bytes
            st.session_state.generated_docx_filename = f"Invoice_{sss_invoice_no}.docx"


# --- Download button (OUTSIDE the form, conditional on a generated file existing) ---
if 'generated_docx_data' in st.session_state and st.session_state.generated_docx_data is not None:
    st.download_button(
        label="Download Invoice (Word)",
        data=st.session_state.generated_docx_data,
        file_name=st.session_state.generated_docx_filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    st.info("PDF generation requires external libraries/tools for exact formatting. We can explore options for this next if the Word document is satisfactory.")
