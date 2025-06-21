from docx import Document


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


def get_formatted_paragraph_html(paragraph):
    html = ""
    for child in paragraph._element:
        tag = child.tag.lower()
        if tag.endswith('}r'):  # paragraphs (run)
            for run in paragraph.runs:
                if run.text in child.text or run.text.strip() in child.text:
                    html += format_run(run)
                    break
        elif tag.endswith('}hyperlink'):  # hyperlink
            html += format_hyperlink(child, paragraph)
    return html


def format_hyperlink(hyperlink_elem, paragraph):
    r_id = hyperlink_elem.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    if not r_id:
        return ""

    rel = paragraph.part.rels.get(r_id)
    href = rel.target_ref if rel else "#"

    text_parts = []
    for run_elem in hyperlink_elem.findall('.//w:r', paragraph._element.nsmap):
        texts = run_elem.findall('.//w:t', paragraph._element.nsmap)
        for t in texts:
            text_parts.append(t.text)

    text = ''.join(text_parts)
    return f'<a href="{href}">{text}</a>'


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
        print(f"List type definition error: {e}")
    return False


def docx_to_html(input_file, output_file):
    doc = Document(input_file)
    html_lines = []
    list_open = False
    list_type = None  # "ul" or "ol"

    for paragraph in doc.paragraphs:
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

        # Convert <h> - titles
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
        # convert list: <ol> and <ul>    
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
