import logging
import re

def modify_first_occurrence(doc, old_text, new_text, paragraph_index=None):
    if paragraph_index is not None:
        try:
            paragraph = doc.paragraphs[paragraph_index]
            if old_text in paragraph.text:
                paragraph.text = paragraph.text.replace(old_text, new_text, 1)
                logging.info(f"文档的段落已修改为: {paragraph.text}")
                return
        except IndexError:
            logging.warning(f"指定的段落索引 {paragraph_index} 超出范围")
            return
    else:
        for paragraph in doc.paragraphs:
            if old_text in paragraph.text:
                paragraph.text = paragraph.text.replace(old_text, new_text, 1)
                logging.info(f"文档的段落已修改为: {paragraph.text}")
                return  # 修改后退出
    logging.warning(f"文档中不包含: {old_text}")

def update_section_references(doc, old_pattern, new_pattern):
    regex = re.compile(old_pattern)

    def update_paragraph_text(paragraph):
        for run in paragraph.runs:
            if regex.search(run.text):
                run.text = regex.sub(new_pattern, run.text)

    def traverse_element(element):
        if hasattr(element, 'paragraphs'):
            for paragraph in element.paragraphs:
                update_paragraph_text(paragraph)
        if hasattr(element, 'tables'):
            for table in element.tables:
                for row in table.rows:
                    for cell in row.cells:
                        traverse_element(cell)

    for paragraph in doc.paragraphs:
        update_paragraph_text(paragraph)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                traverse_element(cell)

    for section in doc.sections:
        traverse_element(section.header)
        traverse_element(section.footer)

    logging.info(f"文档的引用已更新")