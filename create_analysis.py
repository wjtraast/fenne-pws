#!/usr/bin/env python3
"""
Create a detachability analysis for Lieke's detail following Sven's methodology
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from datetime import date

def create_analysis_pdf():
    # Create PDF
    pdf_filename = "src/lieke/EVE6.1_Analyse_Losmaakbaarheid_Lieke.pdf"
    doc = SimpleDocTemplate(pdf_filename, pagesize=A4,
                            topMargin=20*mm, bottomMargin=20*mm,
                            leftMargin=20*mm, rightMargin=20*mm)

    # Container for elements
    elements = []

    # Styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#1f4788'),
        spaceAfter=30,
        alignment=TA_CENTER
    )

    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.HexColor('#1f4788'),
        spaceAfter=12,
        spaceBefore=12
    )

    # Title page
    elements.append(Spacer(1, 40*mm))
    elements.append(Paragraph("Analyse: Losmaakbaarheid", title_style))
    elements.append(Paragraph("Detail Zuidgevel Balkon", styles['Heading2']))
    elements.append(Spacer(1, 10*mm))
    elements.append(Paragraph(f"Gebaseerd op EVE 6.1 - Detail Lieke Traast", styles['Normal']))
    elements.append(Paragraph(f"Datum: {date.today().strftime('%d-%m-%Y')}", styles['Normal']))
    elements.append(PageBreak())

    # Page 1: Elements to assess
    elements.append(Paragraph("Analyse: Losmaakbaarheid", heading_style))
    elements.append(Paragraph("Te beoordelen elementen (in bouwvolgorde):", styles['Normal']))
    elements.append(Spacer(1, 10*mm))

    element_list = [
        "1. Betonnen kanaalplaatvloer",
        "2. Geveldrager (met bouten en stelplaat)",
        "3. Gevelopbouw (beton 290mm + dampremmende folie)",
        "4. Isolatie (190mm)",
        "5. Waterkerende folie",
        "6. Spouw (25mm)",
        "7. Metselwerk buitenspouwblad (90mm)",
        "8. Metselwerk binnenspouwblad (90mm)",
        "9. Cement dekvloer (50mm)",
        "10. Balkonvloer opbouw (breedplaatvloer + egalisatielaag + lijmlaag + tegels)"
    ]

    for item in element_list:
        elements.append(Paragraph(item, styles['Normal']))
        elements.append(Spacer(1, 3*mm))

    elements.append(PageBreak())

    # Page 2: Type verbinding (TV)
    elements.append(Paragraph("Analyse: Losmaakbaarheid", heading_style))
    elements.append(Paragraph("Type Verbinding (TV)", styles['Heading3']))
    elements.append(Spacer(1, 5*mm))

    tv_data = [
        ['Element', 'Hoe verbonden', 'Score (TV)'],
        ['1. Betonnen kanaalplaatvloer', 'Stelspecie op draagconstructie', '0,8'],
        ['2. Geveldrager', 'Bouten door stelplaat', '0,8'],
        ['3. Gevelopbouw (beton)', 'Halfen lijmanker type HB-V', '0,6'],
        ['4. Isolatie', 'Klemming in constructie', '0,8'],
        ['5. Waterkerende folie', 'Mechanisch bevestigd', '0,6'],
        ['6. Spouw', 'Niet van toepassing (holle ruimte)', '1,0'],
        ['7. Metselwerk buitenspouwblad', 'Spouwankers + veeranker', '0,6'],
        ['8. Metselwerk binnenspouwblad', 'Cementgebonden aan elkaar', '0,1'],
        ['9. Cement dekvloer', 'Cementgebonden op kanaalplaat', '0,1'],
        ['10. Balkonvloer opbouw', 'Lijm + mechanisch', '0,4'],
    ]

    tv_table = Table(tv_data, colWidths=[60*mm, 70*mm, 30*mm])
    tv_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4788')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (2, 0), (2, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
    ]))

    elements.append(tv_table)
    elements.append(PageBreak())

    # Page 3: Toegankelijkheid verbinding (ToV)
    elements.append(Paragraph("Analyse: Losmaakbaarheid", heading_style))
    elements.append(Paragraph("Toegankelijkheid Verbinding (ToV)", styles['Heading3']))
    elements.append(Spacer(1, 5*mm))

    tov_data = [
        ['Element', 'Toegankelijkheid', 'Score (ToV)'],
        ['1. Betonnen kanaalplaatvloer', 'Gedeeltelijk herstelbare schade', '0,4'],
        ['2. Geveldrager', 'Toegankelijk met extra handeling', '0,8'],
        ['3. Gevelopbouw (beton)', 'Gedeeltelijk herstelbare schade', '0,4'],
        ['4. Isolatie', 'Toegankelijk met extra handeling', '0,8'],
        ['5. Waterkerende folie', 'Vrij toegankelijk', '1,0'],
        ['6. Spouw', 'Vrij toegankelijk', '1,0'],
        ['7. Metselwerk buitenspouwblad', 'Toegankelijk met extra handeling', '0,8'],
        ['8. Metselwerk binnenspouwblad', 'Gedeeltelijk herstelbare schade', '0,4'],
        ['9. Cement dekvloer', 'Niet-herstelbare schade', '0,1'],
        ['10. Balkonvloer opbouw', 'Vrij toegankelijk', '1,0'],
    ]

    tov_table = Table(tov_data, colWidths=[60*mm, 70*mm, 30*mm])
    tov_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4788')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (2, 0), (2, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
    ]))

    elements.append(tov_table)
    elements.append(PageBreak())

    # Page 4: Doorkruisingen (DO)
    elements.append(Paragraph("Analyse: Losmaakbaarheid", heading_style))
    elements.append(Paragraph("Doorkruisingen (DO)", styles['Heading3']))
    elements.append(Spacer(1, 5*mm))

    do_data = [
        ['Element', 'Doorkruising', 'Score (DO)'],
        ['1. Betonnen kanaalplaatvloer', 'Incidentele doorkruising (ankers)', '0,4'],
        ['2. Geveldrager', 'Geen doorkruising', '1,0'],
        ['3. Gevelopbouw (beton)', 'Incidentele doorkruising (ankers)', '0,4'],
        ['4. Isolatie', 'Geen doorkruising', '1,0'],
        ['5. Waterkerende folie', 'Geen doorkruising', '1,0'],
        ['6. Spouw', 'Geen doorkruising', '1,0'],
        ['7. Metselwerk buitenspouwblad', 'Incidentele doorkruising (ankers)', '0,4'],
        ['8. Metselwerk binnenspouwblad', 'Volledige integratie (metselverband)', '0,1'],
        ['9. Cement dekvloer', 'Geen doorkruising', '1,0'],
        ['10. Balkonvloer opbouw', 'Geen doorkruising', '1,0'],
    ]

    do_table = Table(do_data, colWidths=[60*mm, 70*mm, 30*mm])
    do_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4788')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (2, 0), (2, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
    ]))

    elements.append(do_table)
    elements.append(PageBreak())

    # Page 5: Randopsluiting (RO)
    elements.append(Paragraph("Analyse: Losmaakbaarheid", heading_style))
    elements.append(Paragraph("Randopsluiting (RO)", styles['Heading3']))
    elements.append(Spacer(1, 5*mm))

    ro_data = [
        ['Element', 'Randopsluiting', 'Score (RO)'],
        ['1. Betonnen kanaalplaatvloer', 'Overlapping met gevelopbouw', '0,4'],
        ['2. Geveldrager', 'Overlapping', '0,4'],
        ['3. Gevelopbouw (beton)', 'Overlapping met dekvloer', '0,4'],
        ['4. Isolatie', 'Gesloten door constructie', '0,4'],
        ['5. Waterkerende folie', 'Overlapping met slabber', '0,4'],
        ['6. Spouw', 'Open', '1,0'],
        ['7. Metselwerk buitenspouwblad', 'Overlapping', '0,4'],
        ['8. Metselwerk binnenspouwblad', 'Overlapping', '0,4'],
        ['9. Cement dekvloer', 'Gesloten door opbouw', '0,4'],
        ['10. Balkonvloer opbouw', 'Open (met dilatatievoeg)', '1,0'],
    ]

    ro_table = Table(ro_data, colWidths=[60*mm, 70*mm, 30*mm])
    ro_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4788')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (2, 0), (2, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
    ]))

    elements.append(ro_table)
    elements.append(PageBreak())

    # Page 6: Final scores and conclusion
    elements.append(Paragraph("Analyse: Losmaakbaarheid", heading_style))
    elements.append(Paragraph("Totaaloverzicht Scores en Losmaakbaarheid (Llp)", styles['Heading3']))
    elements.append(Spacer(1, 5*mm))

    final_data = [
        ['Element', 'TV', 'ToV', 'DO', 'RO', 'Llp'],
        ['1. Betonnen kanaalplaatvloer', '0,8', '0,4', '0,4', '0,4', '0,48'],
        ['2. Geveldrager', '0,8', '0,8', '1,0', '0,4', '0,73'],
        ['3. Gevelopbouw (beton)', '0,6', '0,4', '0,4', '0,4', '0,43'],
        ['4. Isolatie', '0,8', '0,8', '1,0', '0,4', '0,73'],
        ['5. Waterkerende folie', '0,6', '1,0', '1,0', '0,4', '0,73'],
        ['6. Spouw', '1,0', '1,0', '1,0', '1,0', '1,00'],
        ['7. Metselwerk buitenspouwblad', '0,6', '0,8', '0,4', '0,4', '0,53'],
        ['8. Metselwerk binnenspouwblad', '0,1', '0,4', '0,1', '0,4', '0,20'],
        ['9. Cement dekvloer', '0,1', '0,1', '1,0', '0,4', '0,28'],
        ['10. Balkonvloer opbouw', '0,4', '1,0', '1,0', '1,0', '0,83'],
    ]

    final_table = Table(final_data, colWidths=[55*mm, 20*mm, 20*mm, 20*mm, 20*mm, 25*mm])
    final_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4788')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
        ('BACKGROUND', (5, 0), (5, -1), colors.HexColor('#4a90e2')),
        ('TEXTCOLOR', (5, 0), (5, 0), colors.whitesmoke),
        ('FONTNAME', (5, 1), (5, -1), 'Helvetica-Bold'),
    ]))

    elements.append(final_table)
    elements.append(Spacer(1, 10*mm))

    # Calculate average
    llp_values = [0.48, 0.73, 0.43, 0.73, 0.73, 1.00, 0.53, 0.20, 0.28, 0.83]
    average_llp = sum(llp_values) / len(llp_values)

    elements.append(Paragraph(f"<b>Gemiddelde Llp: {average_llp:.2f}</b>", styles['Normal']))
    elements.append(Spacer(1, 5*mm))

    # Conclusion
    if average_llp >= 0.7:
        conclusion = "goed losmaakbaar"
        color = "green"
    elif average_llp >= 0.4:
        conclusion = "matig losmaakbaar"
        color = "orange"
    else:
        conclusion = "slecht losmaakbaar"
        color = "red"

    conclusion_style = ParagraphStyle(
        'Conclusion',
        parent=styles['Normal'],
        fontSize=12,
        textColor=colors.HexColor('#1f4788'),
        spaceAfter=12,
        spaceBefore=12,
        fontName='Helvetica-Bold'
    )

    elements.append(Paragraph(f"De gemiddelde score voor losmaakbaarheid is <b>{conclusion}</b>.", conclusion_style))
    elements.append(Spacer(1, 10*mm))

    # Analysis notes
    elements.append(Paragraph("Aandachtspunten:", styles['Heading3']))
    notes = [
        "• De spouw (element 6) scoort uitstekend met 1,00 - volledig losmaakbaar",
        "• De balkonvloer opbouw (element 10) is goed losmaakbaar met score 0,83 dankzij de dilatatievoeg en mechanische verbindingen",
        "• Kritieke punten zijn het metselwerk binnenspouwblad (0,20) en de cement dekvloer (0,28) door cementgebonden verbindingen",
        "• De geveldrager, isolatie en waterkerende folie scoren goed (0,73) door toegankelijke en omkeerbare verbindingen",
        "• Voor verbetering: overwegen van mechanische verbindingen in plaats van cementgebonden voor dekvloer en metselwerk",
        "• De Halfen lijmankers en bouten bij de geveldrager maken demontage goed mogelijk"
    ]

    for note in notes:
        elements.append(Paragraph(note, styles['Normal']))
        elements.append(Spacer(1, 2*mm))

    # Build PDF
    doc.build(elements)
    print(f"PDF created successfully: {pdf_filename}")
    return pdf_filename

if __name__ == "__main__":
    create_analysis_pdf()

