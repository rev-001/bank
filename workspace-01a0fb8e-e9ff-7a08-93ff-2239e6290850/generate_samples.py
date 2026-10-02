"""
Generate sample PDF documents (Invoices, Multi-page POs, Bank Statements, Vendor Bills)
for testing and demonstration of the Offline Agent.
"""

import os
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


def generate_samples():
    os.makedirs("/home/user/samples", exist_ok=True)
    styles = getSampleStyleSheet()

    # Title & Normal Styles
    title_style = ParagraphStyle(
        'InvoiceTitle',
        parent=styles['Heading1'],
        fontSize=20,
        textColor=colors.HexColor('#1E3A8A'),
        spaceAfter=10
    )
    bold_style = ParagraphStyle('BoldText', parent=styles['Normal'], fontName='Helvetica-Bold')
    normal_style = styles['Normal']

    # -------------------------------------------------------------
    # 1. Standard Corporate Invoice
    # -------------------------------------------------------------
    doc1 = SimpleDocTemplate("/home/user/samples/invoice_standard.pdf", pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    story1 = []

    story1.append(Paragraph("INVOICE", title_style))
    story1.append(Spacer(1, 10))

    meta_data = [
        [Paragraph("<b>Vendor:</b> Apex Global Technologies Inc.", normal_style), Paragraph("<b>Invoice No:</b> INV-2026-8842", normal_style)],
        [Paragraph("<b>Address:</b> 100 Innovation Way, Suite 400", normal_style), Paragraph("<b>Invoice Date:</b> 2026-09-15", normal_style)],
        [Paragraph("<b>Bill To:</b> Nexus Enterprise Solutions", normal_style), Paragraph("<b>Due Date:</b> 2026-10-15", normal_style)],
        [Paragraph("<b>Tax ID:</b> US-948271049", normal_style), Paragraph("<b>Payment Terms:</b> Net 30", normal_style)],
    ]
    meta_table = Table(meta_data, colWidths=[300, 240])
    meta_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story1.append(meta_table)
    story1.append(Spacer(1, 20))

    items_data = [
        ["Item Code", "Description", "Quantity", "Unit Price", "Total Price"],
        ["SRV-01", "Cloud Infrastructure Architecture Consulting", "40", "$150.00", "$6,000.00"],
        ["LIC-04", "Enterprise Database Annual License", "5", "$1,200.00", "$6,000.00"],
        ["SUP-12", "24/7 Dedicated Support Tier 3 (Hours)", "25", "$95.00", "$2,375.00"],
        ["DEV-08", "Custom Integration & API Development", "60", "$120.00", "$7,200.00"],
        ["SEC-02", "Cybersecurity Audit & Penetration Testing", "1", "$3,500.00", "$3,500.00"],
    ]
    items_table = Table(items_data, colWidths=[70, 230, 60, 80, 100])
    items_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    story1.append(items_table)
    story1.append(Spacer(1, 15))

    totals_data = [
        ["", "", "Subtotal:", "$25,075.00"],
        ["", "", "Tax Amount (8.0%):", "$2,006.00"],
        ["", "", "Discount:", "$500.00"],
        ["", "", "Grand Total:", "$26,581.00"],
    ]
    totals_table = Table(totals_data, colWidths=[200, 100, 140, 100])
    totals_table.setStyle(TableStyle([
        ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
        ('FONTNAME', (2, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (2, -1), (-1, -1), 11),
        ('LINEABOVE', (2, -1), (-1, -1), 1, colors.HexColor('#1E3A8A')),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story1.append(totals_table)
    doc1.build(story1)

    # -------------------------------------------------------------
    # 2. Multi-Page Purchase Order
    # -------------------------------------------------------------
    doc2 = SimpleDocTemplate("/home/user/samples/purchase_order_multipage.pdf", pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    story2 = []

    story2.append(Paragraph("PURCHASE ORDER", title_style))
    story2.append(Spacer(1, 10))

    po_meta = [
        [Paragraph("<b>Buyer:</b> Global Logistics Corp.", normal_style), Paragraph("<b>PO Number:</b> PO-2026-99120", normal_style)],
        [Paragraph("<b>Vendor:</b> Industrial Hardware Ltd.", normal_style), Paragraph("<b>Order Date:</b> 2026-08-20", normal_style)],
        [Paragraph("<b>Shipping Address:</b> Dock 4, Port Terminal", normal_style), Paragraph("<b>Payment Terms:</b> Net 45", normal_style)],
    ]
    po_table = Table(po_meta, colWidths=[300, 240])
    po_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story2.append(po_table)
    story2.append(Spacer(1, 15))

    po_items_p1 = [
        ["Line #", "Part Number", "Description", "Quantity", "Unit Cost", "Line Total"],
        ["1", "PRT-901", "Heavy Duty Steel Ball Bearings 50mm", "100", "$14.50", "$1,450.00"],
        ["2", "PRT-902", "High Pressure Hydraulic Hose 10m", "25", "$85.00", "$2,125.00"],
        ["3", "PRT-903", "Industrial Valve Actuator Assembly", "8", "$340.00", "$2,720.00"],
        ["4", "PRT-904", "Rotary Shaft Encoder 1024 PPR", "15", "$110.00", "$1,650.00"],
        ["5", "PRT-905", "Pneumatic Cylinder Double Acting", "12", "$175.00", "$2,100.00"],
        ["6", "PRT-906", "Stainless Steel Flange 4-inch", "50", "$32.00", "$1,600.00"],
        ["7", "PRT-907", "Linear Motion Guide Rail 2000mm", "6", "$290.00", "$1,740.00"],
        ["8", "PRT-908", "Precision Gearbox 10:1 Ratio", "4", "$650.00", "$2,600.00"],
    ]
    po_tbl1 = Table(po_items_p1, colWidths=[40, 75, 205, 55, 75, 90])
    po_tbl1.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0F766E')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('ALIGN', (3, 0), (-1, -1), 'RIGHT'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F0FDFA')]),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    story2.append(po_tbl1)
    
    # Page Break for multi-page demonstration
    story2.append(PageBreak())

    story2.append(Paragraph("PURCHASE ORDER (Continued - Page 2)", ParagraphStyle('SubHeader', parent=styles['Heading2'], textColor=colors.HexColor('#0F766E'))))
    story2.append(Spacer(1, 10))

    po_items_p2 = [
        ["Line #", "Part Number", "Description", "Quantity", "Unit Cost", "Line Total"],
        ["9", "PRT-909", "Ceramic Insulator Bushing High Volt", "30", "$45.00", "$1,350.00"],
        ["10", "PRT-910", "Submersible Sump Pump 3HP", "3", "$890.00", "$2,670.00"],
        ["11", "PRT-911", "Fiber Optic Transceiver Module", "20", "$75.00", "$1,500.00"],
        ["12", "PRT-912", "Variable Frequency Drive 15kW", "2", "$1,450.00", "$2,900.00"],
    ]
    po_tbl2 = Table(po_items_p2, colWidths=[40, 75, 205, 55, 75, 90])
    po_tbl2.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0F766E')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('ALIGN', (3, 0), (-1, -1), 'RIGHT'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F0FDFA')]),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    story2.append(po_tbl2)
    story2.append(Spacer(1, 15))

    po_totals = [
        ["", "", "Subtotal:", "$24,405.00"],
        ["", "", "Sales Tax (6.5%):", "$1,586.33"],
        ["", "", "Grand Total:", "$25,991.33"],
    ]
    po_tot_tbl = Table(po_totals, colWidths=[180, 110, 150, 100])
    po_tot_tbl.setStyle(TableStyle([
        ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
        ('FONTNAME', (2, -1), (-1, -1), 'Helvetica-Bold'),
        ('LINEABOVE', (2, -1), (-1, -1), 1, colors.HexColor('#0F766E')),
    ]))
    story2.append(po_tot_tbl)
    doc2.build(story2)

    # -------------------------------------------------------------
    # 3. Noisy Vendor Bill with Custom Header Synonyms
    # -------------------------------------------------------------
    doc3 = SimpleDocTemplate("/home/user/samples/noisy_vendor_bill.pdf", pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    story3 = []

    story3.append(Paragraph("TAX INVOICE / BILL OF SUPPLY", title_style))
    story3.append(Spacer(1, 10))

    bill_meta = [
        [Paragraph("<b>Supplier:</b> Bharat Industrial Components", normal_style), Paragraph("<b>Bill #:</b> BIC-7721", normal_style)],
        [Paragraph("<b>Customer Name:</b> Precision Fabricators", normal_style), Paragraph("<b>Date:</b> 10/09/2026", normal_style)],
        [Paragraph("<b>GSTIN:</b> 33AAAAA0000A1Z5", normal_style), Paragraph("<b>Total Payable:</b> ₹18,400.00", normal_style)],
    ]
    b_table = Table(bill_meta, colWidths=[300, 240])
    b_table.setStyle(TableStyle([('BOTTOMPADDING', (0, 0), (-1, -1), 4)]))
    story3.append(b_table)
    story3.append(Spacer(1, 15))

    # Note non-standard header names: Particulars, Qnty, Rate/Unit, Line Amt
    noisy_items = [
        ["Particulars", "HSN/SAC", "Qnty", "Rate/Unit", "Line Amt"],
        ["Galvanized M12 Hex Bolts (Box of 100)", "7318", "10", "450.00", "4,500.00"],
        ["High Tensile Studs 16x200mm", "7318", "20", "320.00", "6,400.00"],
        ["EPDM Rubber Flange Gasket 4 Inch", "4016", "15", "180.00", "2,700.00"],
        ["Silicone Sealant High Temp 310ml Cartridge", "3214", "12", "400.00", "4,800.00"],
    ]
    noisy_tbl = Table(noisy_items, colWidths=[210, 70, 60, 90, 110])
    noisy_tbl.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#334155')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    story3.append(noisy_tbl)
    story3.append(Spacer(1, 15))

    b_totals = [
        ["", "", "Subtotal:", "₹18,400.00"],
        ["", "", "Grand Total:", "₹18,400.00"],
    ]
    b_tot_tbl = Table(b_totals, colWidths=[180, 110, 150, 100])
    b_tot_tbl.setStyle(TableStyle([
        ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
        ('FONTNAME', (2, -1), (-1, -1), 'Helvetica-Bold'),
    ]))
    story3.append(b_tot_tbl)
    doc3.build(story3)

    # -------------------------------------------------------------
    # 4. Bank Statement
    # -------------------------------------------------------------
    doc4 = SimpleDocTemplate("/home/user/samples/bank_statement.pdf", pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    story4 = []

    story4.append(Paragraph("MONTHLY ACCOUNT STATEMENT", title_style))
    story4.append(Spacer(1, 10))

    bank_meta = [
        [Paragraph("<b>Account Name:</b> Quantum Venture Partners", normal_style), Paragraph("<b>Account #:</b> 9874-5510-2391", normal_style)],
        [Paragraph("<b>Statement Period:</b> 01/08/2026 - 31/08/2026", normal_style), Paragraph("<b>Currency:</b> USD ($)", normal_style)],
    ]
    bank_tbl = Table(bank_meta, colWidths=[300, 240])
    bank_tbl.setStyle(TableStyle([('BOTTOMPADDING', (0, 0), (-1, -1), 4)]))
    story4.append(bank_tbl)
    story4.append(Spacer(1, 15))

    statement_data = [
        ["Date", "Ref No", "Description", "Debit", "Credit", "Balance"],
        ["2026-08-01", "INIT-001", "Opening Balance", "", "", "$45,200.00"],
        ["2026-08-04", "TXN-8812", "Client Retainer Acme Corp", "", "$12,500.00", "$57,700.00"],
        ["2026-08-08", "TXN-8813", "AWS Cloud Services Hosting", "$1,450.00", "", "$56,250.00"],
        ["2026-08-12", "TXN-8814", "Office Lease Downtown Tower", "$4,800.00", "", "$51,450.00"],
        ["2026-08-18", "TXN-8815", "Consulting Invoice INV-8842", "", "$26,581.00", "$78,031.00"],
        ["2026-08-25", "TXN-8816", "Payroll Distribution August", "$18,200.00", "", "$59,831.00"],
        ["2026-08-29", "TXN-8817", "Software Licenses Subscription", "$850.00", "", "$58,981.00"],
    ]
    stmt_table = Table(statement_data, colWidths=[70, 70, 180, 70, 70, 80])
    stmt_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E293B')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('ALIGN', (3, 0), (-1, -1), 'RIGHT'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    story4.append(stmt_table)
    doc4.build(story4)

    print("Sample PDF documents generated successfully in /home/user/samples/")


if __name__ == "__main__":
    generate_samples()
