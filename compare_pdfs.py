import os
import PyPDF2
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from difflib import ndiff

def read_pdf_lines(file_path):
    lines = []
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            text = page.extract_text()
            lines.extend(text.split('\n'))
    return lines

def get_differences(lines1, lines2):
    differences = []
    for i, diff in enumerate(ndiff(lines1, lines2)):
        if diff[0] in ['+', '-']:
            differences.append(f"Line {i}: {diff}")
    return differences

def compare_pdfs(file1, file2):
    lines1 = read_pdf_lines(file1)
    lines2 = read_pdf_lines(file2)
    is_identical = lines1 == lines2
    differences = [] if is_identical else get_differences(lines1, lines2)
    return is_identical, differences

def generate_report(results, total_files, output_path='comparison_report.pdf'):
    doc = SimpleDocTemplate(output_path, pagesize=letter, rightMargin=30, leftMargin=30)
    styles = getSampleStyleSheet()
    story = []
    
    header_style = ParagraphStyle(
        'CustomHeader',
        parent=styles['Heading1'],
        fontSize=14,
        spaceAfter=30
    )
    
    cell_style = ParagraphStyle(
        'CellStyle',
        parent=styles['Normal'],
        fontSize=9,
        leading=12,
        wordWrap='CJK'
    )
    
    story.append(Paragraph(f"PDF Comparison Report - Total Files: {total_files}", header_style))
    story.append(Spacer(1, 20))
    
    table_data = [['File Name', 'Status', 'Differences']]
    for result in results:
        file_name, is_identical, differences = result
        status = 'Identical' if is_identical else 'Different'
        diff_text = '<br/>'.join(differences) if differences else 'N/A'
        
        # Wrap cell contents in Paragraph objects
        row = [
            Paragraph(file_name, cell_style),
            Paragraph(status, cell_style),
            Paragraph(diff_text, cell_style)
        ]
        table_data.append(row)
    
    table = Table(table_data, colWidths=[120, 60, 330])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ('TOPPADDING', (0, 0), (-1, -1), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    
    story.append(table)
    doc.build(story)

def main():
    location1 = 'C:/Users/U059611/PDFcompare/old'
    location2 = 'C:/Users/U059611/PDFcompare/new'

    pdf_files1 = sorted([f for f in os.listdir(location1) if f.endswith('.pdf')])
    pdf_files2 = sorted([f for f in os.listdir(location2) if f.endswith('.pdf')])

    # Find common files
    common_files = set(pdf_files1).intersection(pdf_files2)
    total_files = len(common_files)
    
    results = []
    for file_name in common_files:
        file1 = os.path.join(location1, file_name)
        file2 = os.path.join(location2, file_name)
        is_identical, differences = compare_pdfs(file1, file2)
        results.append((file_name, is_identical, differences))
        
        # Print results to console
        status = "identical" if is_identical else "different"
        print(f"{file_name} is {status}")
        if differences:
            print("Differences found:")
            for diff in differences:
                print(diff)
            print("-" * 50)
    
    # Generate PDF report
    generate_report(results, total_files)
    print(f"\nReport generated with {total_files} files compared.")

if __name__ == "__main__":
    main()