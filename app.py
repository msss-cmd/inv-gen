import streamlit as st
import datetime
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.shared import RGBColor
import io # To handle file in memory

# --- Streamlit Page Setup ---
st.set_page_config(layout="wide", page_title="Invoice Generator")

# --- Function to generate DOCX ---
def generate_invoice_docx(invoice_data):
    document = Document()

    # --- Header Section ---
    title = document.add_paragraph('TAX INVOICE/DELIVERY NOTE', style='Normal')
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title.runs[0]
    run_title.bold = True
    run_title.font.size = Pt(14)

    document.add_paragraph() # Add a blank line for spacing

    # Table for Contact and Invoice Details
    table_header_data = [
        ["To:", invoice_data['to_company'], "Date:", invoice_data['invoice_date']],
        ["Address:", invoice_data['customer_address'], "SSS Invoice No:", invoice_data['sss_invoice_no']],
        ["Tel:", invoice_data['customer_tel'], "Customer VAT No.:", invoice_data['customer_vat_no']],
        ["ATTN:", invoice_data['attn_person'], "", ""],
        ["Email:", invoice_data['customer_email'], "", ""],
        ["Customer PO#:", invoice_data['customer_po'], "", ""]
    ]

    table = document.add_table(rows=len(table_header_data), cols=4)
    table.autofit = False
    table.allow_autofit = False

    col_widths = [Cm(2.5), Cm(6.5), Cm(3), Cm(6)]
    for i, width in enumerate(col_widths):
        table.columns[i].width = width

    for r_idx, row_data in enumerate(table_header_data):
        for c_idx, cell_text in enumerate(row_data):
            cell = table.rows[r_idx].cells[c_idx]
            cell.text = str(cell_text)
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)

    # --- Line Items Table ---
    item_table_headers = ["No", "Description", "Unit price", "BHD", "Qty", "Total price", "BHD"]
    item_table = document.add_table(rows=1, cols=7)
    item_table.autofit = False
    item_table.allow_autofit = False

    item_col_widths = [Cm(1.5), Cm(7), Cm(2.5), Cm(1.5), Cm(1.5), Cm(2.5), Cm(1.5)]
    for i, width in enumerate(item_col_widths):
        item_table.columns[i].width = width

    hdr_cells = item_table.rows[0].cells
    for i, header_text in enumerate(item_table_headers):
        cell = hdr_cells[i]
        cell.text = header_text
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if header_text in ["No", "Qty", "BHD"] else WD_ALIGN_PARAGRAPH.LEFT
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(9)

    for i, item in enumerate(invoice_data['line_items']):
        row_cells = item_table.add_row().cells
        row_cells[0].text = str(i + 1)
        row_cells[1].text = item['description']
        row_cells[2].text = f"{item['unit_price']:.3f}"
        row_cells[3].text = "BHD"
        row_cells[4].text = str(item['quantity'])
        row_cells[5].text = f"{item['total_price']:.3f}"
        row_cells[6].text = "BHD"

        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[5].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

        for j in [0, 1, 2, 3, 4, 5, 6]:
            for paragraph in row_cells[j].paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(9)

    # --- Summary Section (Subtotal, VAT, Grand Total) ---
    summary_table = document.add_table(rows=3, cols=2)
    summary_table.autofit = False
    summary_table.allow_autofit = False

    summary_col_widths = [Cm(15.5), Cm(4.5)]
    summary_table.columns[0].width = summary_col_widths[0]
    summary_table.columns[1].width = summary_col_widths[1]

    subtotal_row = summary_table.rows[0].cells
    subtotal_row[0].text = "Subtotal:"
    subtotal_row[1].text = f"{invoice_data['subtotal']:.3f} BHD"
    subtotal_row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    subtotal_row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    vat_row = summary_table.rows[1].cells
    vat_row[0].text = "VAT @ 10%:"
    vat_row[1].text = f"{invoice_data['vat_amount']:.3f} BHD"
    vat_row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    vat_row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    grand_total_row = summary_table.rows[2].cells
    grand_total_row[0].text = "Grand Total in BHD"
    grand_total_row[1].text = f"{invoice_data['grand_total']:.3f} BHD"
    grand_total_row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    grand_total_row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    for p in grand_total_row[0].paragraphs:
        for run in p.runs:
            run.bold = True
            run.font.size = Pt(11)

    for p in grand_total_row[1].paragraphs:
        for run in p.runs:
            run.bold = True
            run.font.size = Pt(11)

    document.add_paragraph()

    # --- Payment Details (Static Text) ---
    payment_details_text = """
    Payment will be remitted to the following account:
    Salahuddin Softtech Solutions, Arab Bank, A/C No.: 2002-146072-510,
    IBAN: BH31 ARAB 0200 2146072510, Swift Code: ARAB BHBM,
    Address: Arab Bank Plc. , P.O Box: 395, Manama, Kingdom of Bahrain.
    """
    document.add_paragraph(payment_details_text.strip())

    # --- Terms and Conditions ---
    document.add_heading('Terms and Conditions', level=2)
    terms_para = document.add_paragraph()
    terms_para.add_run("Payment terms: ").bold = True
    terms_para.add_run(invoice_data['payment_terms'])

    document.add_paragraph("Title and property at all above remain ours until full payment is received and we reserve the rights to withdraw goods/services if not paid for when due.")
    document.add_paragraph("Signing this document implies acceptance of these terms.")

    # --- Receipt/Acknowledgement ---
    document.add_paragraph()
    received_by_para = document.add_paragraph("Received by")
    received_by_para.add_run("\t\t   (Name)\t\t   (Date)\t          (Signature)").font.name = 'Courier New'

    document.add_paragraph("For SALAHUDDIN SOFTTECH SOLUTIONS")
    document.add_paragraph("Jobin George")
    document.add_paragraph("Operations Manager")

    byte_io = io.BytesIO()
    document.save(byte_io)
    byte_io.seek(0)
    return byte_io

# --- Streamlit Form Layout ---
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

    # Store generated DOCX in session_state if form was submitted successfully
    if submitted:
        if not st.session_state.line_items:
            st.warning("Please add at least one line item to generate the invoice.")
            # Clear any previously stored generated file
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

            # Generate the DOCX file in memory
            docx_file_bytes = generate_invoice_docx(invoice_data)

            # Store the generated file data and filename in session state
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
    # You might want to clear the session state after download if you only want one download per generation
    # del st.session_state.generated_docx_data
    # del st.session_state.generated_docx_filename
