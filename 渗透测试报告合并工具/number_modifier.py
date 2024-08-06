import re
import win32com.client
import logging
import os
import tempfile

from docx import Document

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def extract_numbered_paragraphs(doc_path):
    try:
        app = win32com.client.Dispatch('Kwps.Application')
        app.Visible = False
        doc = app.Documents.Open(doc_path)
    except Exception as e:
        logging.error(f"打开文档 {doc_path} 时出错: {e}")
        raise

    numbered_paragraphs = []
    pattern = re.compile(r'^\d+(\.\d+)*[\.\)\(]*$')

    for paragraph in doc.Paragraphs:
        if paragraph.Range.ListFormat.ListString:
            numbering_text = paragraph.Range.ListFormat.ListString
            if numbering_text.strip() and numbering_text.replace('.', '').replace(')', '').replace('(', '').isnumeric():
                numbered_paragraphs.append((paragraph, numbering_text))
        else:
            text = paragraph.Range.Text.strip()
            match = pattern.match(text)
            if match and match.start() == 0:
                numbering_text = match.group()
                if numbering_text.strip():
                    numbered_paragraphs.append((paragraph, numbering_text))

    return numbered_paragraphs, doc, app

def get_max_main_chapter_number_from_toc(toc_doc_path):
    try:
        app = win32com.client.Dispatch('Kwps.Application')
        app.Visible = False
        doc = app.Documents.Open(toc_doc_path)
    except Exception as e:
        logging.error(f"打开文档 {toc_doc_path} 时出错: {e}")
        raise

    max_chapter = 0
    pattern = re.compile(r'^\d+(\.\d+)*[\.\)\(]*$')

    for table in doc.TablesOfContents:
        for entry in table.Range.Paragraphs:
            text = entry.Range.Text.strip()
            parts = text.split()
            if not parts:
                continue

            potential_number = parts[0]
            match = pattern.match(potential_number)
            if match and ')' not in potential_number:
                cleaned_numbering_text = re.sub(r'[^\d.]', '', match.group())
                main_chapter = int(cleaned_numbering_text.split('.')[0])
                if main_chapter > max_chapter:
                    max_chapter = main_chapter

    doc.Close(False)
    app.Quit()
    logging.info(f"目录中的最大章节号: {max_chapter}")
    return max_chapter

def copy_content_after_toc(doc_path):
    try:
        app = win32com.client.Dispatch('Kwps.Application')
        app.Visible = False
        doc = app.Documents.Open(doc_path)

        temp_doc_path = os.path.join(tempfile.gettempdir(), 'temp_sub_doc.docx')
        temp_doc = app.Documents.Add()

        toc_end_position = None
        for table in doc.TablesOfContents:
            toc_end_position = table.Range.End

        if toc_end_position is not None:
            range_to_copy = doc.Range(Start=toc_end_position, End=doc.Content.End)
            range_to_copy.Copy()
            temp_doc.Range().Paste()

            temp_doc.SaveAs(temp_doc_path)
            temp_doc.Close(False)
            doc.Close(False)
            app.Quit()
            return temp_doc_path

        # 如果没有目录页，返回原始路径
        doc.Close(False)
        app.Quit()
        return doc_path

    except Exception as e:
        logging.error(f"复制内容时出错: {e}")
        raise

def extract_and_modify_numbered_paragraphs(content_doc_path, toc_doc_path, output_doc_path):
    doc = Document(content_doc_path)
    temp_content_doc_path = content_doc_path
    has_toc = any(table._element for table in doc.tables if 'Table of Contents' in table.style.name)
    if has_toc:
        logging.info("检测到目录页，复制目录之后的内容...")
        temp_content_doc_path = copy_content_after_toc(content_doc_path)

    max_chapter = get_max_main_chapter_number_from_toc(toc_doc_path)
    new_chapter_prefix = f"{max_chapter + 1}."

    numbered_paragraphs, win32_doc, app = extract_numbered_paragraphs(temp_content_doc_path)

    modified_paragraphs = []
    pattern_to_modify = re.compile(r'^\d+(\.\d+)*$')
    skip_patterns = re.compile(r'^\(\d+\)|^\d+\)$')

    for para, num in numbered_paragraphs:
        if para.Range.ListFormat.ListString:
            para.Range.ListFormat.RemoveNumbers()
        else:
            text = para.Range.Text.strip()
            new_text = re.sub(r'^\d+(\.\d+)*[\.\)\(]*\s*', '', text)
            para.Range.Text = new_text

        if para.Range.Tables.Count > 0 or skip_patterns.match(num):
            modified_paragraphs.append((para, num))
        else:
            modified_paragraphs.append((para, f"{new_chapter_prefix}{num}"))

    for paragraph, new_numbering in modified_paragraphs:
        paragraph.Range.InsertBefore(f"{new_numbering}")

    try:
        win32_doc.SaveAs(output_doc_path)
        logging.info(f"修改后的文档已保存到 {output_doc_path}")
    except Exception as e:
        logging.error(f"保存文档 {output_doc_path} 时出错: {e}")
        raise
    finally:
        win32_doc.Close(False)
        app.Quit()

if __name__ == "__main__":
    content_doc_path = 'path_to_content_doc.docx'
    toc_doc_path = 'path_to_toc_doc.docx'
    output_doc_path = 'path_to_output_doc.docx'
    extract_and_modify_numbered_paragraphs(content_doc_path, toc_doc_path, output_doc_path)
