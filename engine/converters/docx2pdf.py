import os
import logging
from pathlib import Path
import pythoncom
import win32com.client

logger = logging.getLogger("callidus")


class PdfConverter:
    def convert(self, source_path: Path, output_pdf: Path):
        """
        Открывает DOCX через MS Word, принудительно обновляет поля (номера страниц,
        оглавление) и сохраняет идеальный PDF и обновленный DOCX.
        """
        source_str = str(source_path.resolve())
        pdf_str = str(output_pdf.resolve())

        # Обязательная инициализация COM для многопоточных серверов (FastAPI)
        pythoncom.CoInitialize()

        word = None
        doc = None
        try:
            # Запускаем невидимый процесс Word (DispatchEx создает независимый процесс)
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False  # Отключаем всплывающие окна

            logger.debug(f"[PdfConverter] Открытие файла в MS Word: {source_path.name}")
            doc = word.Documents.Open(source_str)

            # 1. Принудительно обновляем все поля (номера страниц, перекрестные ссылки)
            doc.Fields.Update()

            # 2. Принудительно обновляем Оглавление (Table of Contents)
            if doc.TablesOfContents.Count > 0:
                for toc in doc.TablesOfContents:
                    toc.Update()

            # Сохраняем обновленный Word-документ (чтобы при открытии DOCX оглавление уже было готово)
            doc.Save()

            # 3. Экспортируем в PDF. 17 - это константа wdFormatPDF
            logger.debug(f"[PdfConverter] Сохранение PDF: {output_pdf.name}")
            doc.SaveAs(pdf_str, FileFormat=17)

        except Exception as e:
            logger.error(f"[PdfConverter] Ошибка при работе с COM-объектом Word: {e}")
            raise e
        finally:
            # Строго обязательно закрываем документ и процесс, иначе в диспетчере задач повиснут процессы WINWORD.EXE
            if doc is not None:
                doc.Close(SaveChanges=0)  # 0 = wdDoNotSaveChanges
            if word is not None:
                word.Quit()

            # Освобождаем COM-ресурсы потока
            pythoncom.CoUninitialize()

        return output_pdf