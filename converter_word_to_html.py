from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


def format_run(run):
    text = run.text
    if not text:
        return ""

    if run.bold and run.italic:
        text = f"<strong><em>{text}</em></strong>"
    elif run.bold:
        text = f"<strong>{text}</strong>"
    elif run.italic:
        text = f"<em>{text}</em>"

    if run.underline:
        text = f"<u>{text}</u>"

    if run.font.strike:
        text = f"<span style=\"text-decoration: line-through;\">{text}</span>"

    return text


def convert_table_to_grid_html(table):
    rows = table.rows
    if not rows:
        return ""

    cols_count = len(rows[0].cells)

    # Cheack short columns
    short_columns = set()
    for col_idx in range(cols_count):
        short_count = 0
        for row in rows:
            text = row.cells[col_idx].text.strip()
            if len(text) <= 8:
                short_count += 1
        if short_count >= len(rows) // 2:
            short_columns.add(col_idx)

    # Create grid-template-columns
    column_styles = []
    for i in range(cols_count):
        if i in short_columns:
            column_styles.append("max-content")
        else:
            column_styles.append("minmax(200px, 1fr)")
    grid_template = ' '.join(column_styles)

    html = (
        f'<div style="display: grid; grid-template-columns: {grid_template}; '
        f'gap: 0;border: 1px solid #999; margin: 12px 0; font-family: sans-serif; '
        f' font-size: 14px; width: max-content;">\n'
    )

    for row in rows:
        for cell in row.cells:
            content = cell.text.strip()
            html += (
                f'<div style="border: 1px solid #999; padding: 8px; text-align: left; '
                f'background-color: #f9f9f9;">{content}</div>\n'
            )

    html += '</div>'
    return html


def get_formatted_paragraph_html(paragraph):
    """
    Форматируем весь текст параграфа, включая форматирование и гиперссылки.
    Проходим по runs подряд.
    We format the entire text of the paragraph,
    including formatting and hyperlinks.
    We go through the runs in a row.
    """
    html = ""

    # Определим для каждого run — внутри ли он гиперссылки
    # В docx гиперссылки — элементы <w:hyperlink> внутри <w:p>
    # Пройдём по XML-структуре параграфа и сопоставим runs с hyperlink runs

    # Let's determine for each run whether it is inside a hyperlink
    # In docx, hyperlinks are <w:hyperlink> elements inside <w:p>
    # Let's go through the paragraph's XML structure
    # and match runs with hyperlink runs

    # Получим все hyperlink элементы
    # Get all hyperlink elements
    hyperlink_elems = list(paragraph._element.findall('.//w:hyperlink', paragraph._element.nsmap))

    # Соберём множества run элементов, которые принадлежат hyperlink-элементам
    # Let's collect sets of run elements that belong to hyperlink elements
    runs_in_hyperlinks = set()
    for h in hyperlink_elems:
        for r in h.findall('.//w:r', paragraph._element.nsmap):
            runs_in_hyperlinks.add(r)

    # Теперь проходим по runs в параграфе,
    # и если run xml-элемент в runs_in_hyperlinks, оборачиваем в <a>
    # Now we go through the runs in the paragraph,
    # and if the run xml element is in runs_in_hyperlinks, we wrap it in <a>
    for run in paragraph.runs:
        run_element = run._element
        text_html = format_run(run)

        if run_element in runs_in_hyperlinks:
            # Найдём hyperlink элемент, которому принадлежит run
            # Let's find the hyperlink element that runs belongs to
            hyperlink = None
            for h in hyperlink_elems:
                if run_element in h.findall('.//w:r', paragraph._element.nsmap):
                    hyperlink = h
                    break
            if hyperlink is not None:
                r_id = hyperlink.get(
                    '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
                )
                href = "#"
                if r_id:
                    rel = paragraph.part.rels.get(r_id)
                    if rel:
                        href = rel.target_ref
                text_html = f'<a href="{href}">{text_html}</a>'

        html += text_html

    return html


def is_ordered_list(paragraph, doc):
    try:
        pPr = paragraph._element.pPr
        if pPr is None or pPr.numPr is None or pPr.numPr.numId is None:
            return False

        num_id = int(pPr.numPr.numId.val)
        numbering = doc.part.numbering_part.numbering_definitions._numbering
        num_elem = numbering.find(
            f'.//w:num[@w:numId="{num_id}"]',
            namespaces=numbering.nsmap)

        if num_elem is not None:
            abstract_id_elem = num_elem.find(
                './w:abstractNumId',
                namespaces=numbering.nsmap)
            if abstract_id_elem is None:
                return False

            abstract_id = abstract_id_elem.get(
                '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'
            )
            abstract_elem = numbering.find(
                f'.//w:abstractNum[@w:abstractNumId="{abstract_id}"]',
                namespaces=numbering.nsmap)

            if abstract_elem is not None:
                numFmt_elem = abstract_elem.find(
                    './/w:numFmt', namespaces=numbering.nsmap)
                if numFmt_elem is not None:
                    fmt = numFmt_elem.get(
                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'
                    )
                    return fmt != 'bullet'  # bullet → <ul>, else <ol>
    except Exception as e:
        print(f"Ошибка определения типа списка: {e}")
    return False


def docx_to_html(input_file, output_file):
    doc = Document(input_file)
    html_lines = []
    list_open = False
    list_type = None  # "ul" or "ol"

    for child in doc.element.body:
        if child.tag.endswith('}p'):  # paragraphs
            paragraph = Paragraph(child, doc)
            text = paragraph.text.strip()
            if not text:
                continue

            style = paragraph.style.name.lower() if paragraph.style else ""
            pPr = paragraph._element.pPr
            has_numbering = pPr is not None and pPr.numPr is not None

            is_list_item = (
                'list' in style or
                has_numbering or
                (text.startswith('- ')) or
                (pPr is not None and pPr.ind is not None)
            )

            is_ordered = is_ordered_list(paragraph, doc)
            content = get_formatted_paragraph_html(paragraph)

            # <h> - titles
            if 'heading 1' in style:
                if list_open:
                    html_lines.append(f"</{list_type}>")
                    list_open = False
                html_lines.append(f"<h1>{content}</h1>")
            elif 'heading 2' in style:
                if list_open:
                    html_lines.append(f"</{list_type}>")
                    list_open = False
                html_lines.append(f"<h2>{content}</h2>")
            elif 'heading 3' in style:
                if list_open:
                    html_lines.append(f"</{list_type}>")
                    list_open = False
                html_lines.append(f"<h3>{content}</h3>")
            elif 'heading 4' in style:
                if list_open:
                    html_lines.append(f"</{list_type}>")
                    list_open = False
                html_lines.append(f"<h4>{content}</h4>")
            elif 'heading 5' in style:
                if list_open:
                    html_lines.append(f"</{list_type}>")
                    list_open = False
                html_lines.append(f"<h5>{content}</h5>")
            elif 'heading 6' in style:
                if list_open:
                    html_lines.append(f"</{list_type}>")
                    list_open = False
                html_lines.append(f"<h6>{content}</h6>")
            elif is_list_item:
                current_list_type = "ol" if is_ordered else "ul"
                if not list_open or list_type != current_list_type:
                    if list_open:
                        html_lines.append(f"</{list_type}>")
                    html_lines.append(f"<{current_list_type}>")
                    list_open = True
                    list_type = current_list_type
                html_lines.append(f"<li>{content}</li>")
            else:
                if list_open:
                    html_lines.append(f"</{list_type}>")
                    list_open = False
                html_lines.append(f"<p>{content}</p>")

        elif child.tag.endswith('}tbl'):  # Table
            table = Table(child, doc)
            if list_open:
                html_lines.append(f"</{list_type}>")
                list_open = False
            html_lines.append(convert_table_to_grid_html(table))

    if list_open:
        html_lines.append(f"</{list_type}>")

    html_content = "\n".join(html_lines)

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html_content)


if __name__ == '__main__':
    try:
        docx_to_html('input.docx', 'output.html')
        print("Конвертация успешно завершена!")
        print("Conversion completed successfully!")
    except Exception as e:
        print(f"Ошибка: {str(e)}")
        print(f"Error: {str(e)}")
