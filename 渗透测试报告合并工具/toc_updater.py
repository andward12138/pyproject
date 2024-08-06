import win32com.client as win32
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def update_table_of_contents(doc_path):
    try:
        word_app = win32.Dispatch("Kwps.Application")
        word_app.Visible = False
        doc = word_app.Documents.Open(doc_path)
    except Exception as e:
        logging.error(f"打开文档 {doc_path} 时出错: {e}")
        raise

    try:
        for toc in doc.TablesOfContents:
            toc.Update()

        doc.Save()
        logging.info(f"文档 {doc_path} 的目录已更新")
    finally:
        doc.Close()
        word_app.Quit()

if __name__ == "__main__":
    doc_path = r'C:\Users\wa\Desktop\4.docx'
    update_table_of_contents(doc_path)
