import os
import csv
import streamlit as st
import re
import json
from datetime import datetime
from crossref.restful import Works
from docx import Document
from docx.oxml.ns import qn
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
from tqdm import tqdm
from docx.oxml import OxmlElement
import base64
import html
import concurrent.futures
from typing import List, Dict, Tuple, Set, Any, Optional
import hashlib
import time
from collections import Counter

works = Works()

# Полный словарь переводов
TRANSLATIONS = {
    'en': {
        'header': '🎨 Citation Style Constructor',
        'general_settings': '⚙️ General Settings',
        'element_config': '📑 Element Configuration',
        'style_preview': '👀 Style Preview',
        'data_input': '📁 Data Input',
        'data_output': '📤 Data Output',
        'numbering_style': 'Numbering:',
        'author_format': 'Authors:',
        'author_separator': 'Separator:',
        'et_al_limit': 'Et al after:',
        'use_and': "'and'",
        'use_ampersand': "'&'",
        'doi_format': 'DOI format:',
        'doi_hyperlink': 'DOI as hyperlink',
        'page_format': 'Pages:',
        'final_punctuation': 'Final punctuation:',
        'element': 'Element',
        'italic': 'Italic',
        'bold': 'Bold',
        'parentheses': 'Parentheses',
        'separator': 'Separator',
        'input_method': 'Input:',
        'output_method': 'Output:',
        'select_docx': 'Select DOCX',
        'enter_references': 'Enter references (one per line)',
        'references': 'References:',
        'results': 'Results:',
        'process': '🚀 Process',
        'example': 'Example:',
        'error_select_element': 'Select at least one element!',
        'processing': '⏳ Processing...',
        'upload_file': 'Upload a file!',
        'enter_references_error': 'Enter references!',
        'select_docx_output': 'Select DOCX output to download!',
        'doi_txt': '📄 DOI (TXT)',
        'references_docx': '📋 References (DOCX)',
        'found_references': 'Found {} references.',
        'found_references_text': 'Found {} references in text.',
        'statistics': 'Statistics: {} DOI found, {} not found.',
        'language': 'Language:',
        'gost_style': 'Apply GOST Style',
        'export_style': '📤 Export Style',
        'import_style': '📥 Import Style',
        'export_file_name': 'File name:',
        'import_file': 'Select style file:',
        'export_success': 'Style exported successfully!',
        'import_success': 'Style imported successfully!',
        'import_error': 'Error importing style file!',
        'processing_status': 'Processing references...',
        'current_reference': 'Current: {}',
        'processed_stats': 'Processed: {}/{} | Found: {} | Errors: {}',
        'time_remaining': 'Estimated time remaining: {}',
        'duplicate_reference': '🔄 Repeated Reference (See #{})',
        'batch_processing': 'Batch processing DOI...',
        'extracting_metadata': 'Extracting metadata...',
        'checking_duplicates': 'Checking for duplicates...',
        'retrying_failed': 'Retrying failed DOI requests...',
        'bibliographic_search': 'Searching by bibliographic data...',
        'style_presets': 'Style Presets',
        'gost_button': 'GOST',
        'acs_button': 'ACS (MDPI)',
        'rsc_button': 'RSC',
        'cta_button': 'CTA',
        'style_preset_tooltip': 'Here are some styles maintained by individual publishers. For major publishers (Elsevier, Springer Nature, and Wiley), styles vary from journal to journal. To create (or reformat) references for a specific journal, use the Citation Style Constructor.',
        'journal_style': 'Journal style:',
        'full_journal_name': 'Full Journal Name',
        'journal_abbr_with_dots': 'J. Abbr.',
        'journal_abbr_no_dots': 'J Abbr',
        'short_guide_title': 'A short guide for the conversion of doi-based references',
        'step_1': '❶ Select a ready reference style (GOST, ACS(MDPI), RSC, or CTA), or create your own style by selecting the sequence, design, and punctuation of the element configurations',
        'step_1_note': '(!) The punctuation boxes enable various items to be included between element configurations (simple punctuation, Vol., Issue…)',
        'step_2': '❷ Then, use the Style Presets to change certain element configurations for each reformatted reference.',
        'step_3': '❸ The Style Preview function enables users to visualize the final form of their reference style',
        'step_4': '❹ If the final style is appropriate, select the Docx or Text option in the Data Input section and upload the corresponding information. Then, in the Data Output section, select the required options and press "Process" to initiate reformatting.',
        'step_5': '❺ After processing is complete, download the reformatted references in your preferred format.',
        'step_5_note': '(!) Outputting the Docx file is recommended, as it preserves formatting (e.g., bold, italic, and hyperlinks) and includes additional stats at the end of the document.',
        'step_6': '❻ After creating your final version of the style, save it so that you can upload it again in the next session. Use the Style Management section for this purpose.'
    },
    'ru': {
        'header': '🎨 Конструктор стилей цитирования',
        'general_settings': '⚙️ Настройки',
        'element_config': '📑 Конфигурация элементов',
        'style_preview': '👀 Предпросмотр',
        'data_input': '📁 Ввод',
        'data_output': '📤 Вывод',
        'numbering_style': 'Нумерация:',
        'author_format': 'Авторы:',
        'author_separator': 'Разделитель:',
        'et_al_limit': 'Et al после:',
        'use_and': "'и'",
        'use_ampersand': "'&'",
        'doi_format': 'Формат DOI:',
        'doi_hyperlink': 'DOI как ссылка',
        'page_format': 'Страницы:',
        'final_punctuation': 'Конечная пунктуация:',
        'element': 'Элемент',
        'italic': 'Курсив',
        'bold': 'Жирный',
        'parentheses': 'Скобки',
        'separator': 'Разделитель',
        'input_method': 'Ввод:',
        'output_method': 'Вывод:',
        'select_docx': 'Выберите DOCX',
        'enter_references': 'Введите ссылки (по одной на строку)',
        'references': 'Ссылки:',
        'results': 'Результаты:',
        'process': '🚀 Обработать',
        'example': 'Пример:',
        'error_select_element': 'Выберите хотя бы один элемент!',
        'processing': '⏳ Обработка...',
        'upload_file': 'Загрузите файл!',
        'enter_references_error': 'Введите ссылки!',
        'select_docx_output': 'Выберите DOCX для скачивания!',
        'doi_txt': '📄 DOI (TXT)',
        'references_docx': '📋 Ссылки (DOCX)',
        'found_references': 'Найдено {} ссылок.',
        'found_references_text': 'Найдено {} ссылок в тексте.',
        'statistics': 'Статистика: {} DOI найдено, {} не найдено.',
        'language': 'Язык:',
        'gost_style': 'Применить стиль ГОСТ',
        'export_style': '📤 Экспорт стиля',
        'import_style': '📥 Импорт стиля',
        'export_file_name': 'Имя файла:',
        'import_file': 'Выберите файл стиля:',
        'export_success': 'Стиль экспортирован успешно!',
        'import_success': 'Стиль импортирован успешно!',
        'import_error': 'Ошибка импорта файла стиля!',
        'processing_status': 'Обработка ссылок...',
        'current_reference': 'Текущая: {}',
        'processed_stats': 'Обработано: {}/{} | Найдено: {} | Ошибки: {}',
        'time_remaining': 'Примерное время до завершения: {}',
        'duplicate_reference': '🔄 Повторная ссылка (См. #{})',
        'batch_processing': 'Пакетная обработка DOI...',
        'extracting_metadata': 'Извлечение метаданных...',
        'checking_duplicates': 'Проверка на дубликаты...',
        'retrying_failed': 'Повторная попытка для неудачных DOI...',
        'bibliographic_search': 'Поиск по библиографическим данным...',
        'style_presets': 'Готовые стили',
        'gost_button': 'ГОСТ',
        'acs_button': 'ACS (MDPI)',
        'rsc_button': 'RSC',
        'cta_button': 'CTA',
        'style_preset_tooltip': 'Здесь указаны некоторые стили, которые сохраняются в пределах одного издательства. Для ряда крупных издательств (Esevier, Springer Nature, Wiley) стиль отличается от журнала к журналу. Для формирования (или переформатирования) ссылок для конкретного журнала предлагаем воспользоваться конструктором ссылок.',
        'journal_style': 'Стиль журнала:',
        'full_journal_name': 'Полное название журнала',
        'journal_abbr_with_dots': 'J. Abbr.',
        'journal_abbr_no_dots': 'J Abbr',
        'short_guide_title': 'Краткое руководство для конвертации ссылок, имеющих doi',
        'step_1': '❶ Выберите готовый стиль ссылок (ГОСТ, ACS(MDPI), RSC или CTA) или создайте свой собственный стиль, выбрав последовательность, оформление и пунктуацию конфигураций элементов',
        'step_1_note': '(!) Поля пунктуации позволяют включать различные элементы между конфигурациями (простая пунктуация, Том, Выпуск…)',
        'step_2': '❷ Затем используйте готовые стили, чтобы изменить определенные конфигурации элементов для каждой переформатированной ссылки.',
        'step_3': '❸ Функция предпросмотра стиля позволяет визуализировать окончательную форму вашего стиля ссылок',
        'step_4': '❹ Если окончательный стиль подходит, выберите опцию Docx или Текст в разделе ввода данных и загрузите соответствующую информацию. Затем в разделе вывода данных выберите нужные опции и нажмите "Обработать" для начала переформатирования.',
        'step_5': '❺ После завершения обработки загрузите переформатированные ссылки в предпочитаемом формате.',
        'step_5_note': '(!) Рекомендуется выводить файл Docx, так как он сохраняет форматирование (например, жирный шрифт, курсив и гиперссылки) и включает дополнительную статистику в конце документа.',
        'step_6': '❻ После создания окончательной версии стиля сохраните его, чтобы можно было снова загрузить в следующей сессии. Для этого используйте раздел Style Management.'
    }
}

# Константы для стилей
NUMBERING_STYLES = ["No numbering", "1", "1.", "1)", "(1)", "[1]"]
AUTHOR_FORMATS = ["AA Smith", "A.A. Smith", "Smith AA", "Smith A.A", "Smith, A.A."]
PAGE_FORMATS = ["122 - 128", "122-128", "122 – 128", "122–128", "122–8", "122"]
DOI_FORMATS = ["10.10/xxx", "doi:10.10/xxx", "DOI:10.10/xxx", "https://dx.doi.org/10.10/xxx"]
JOURNAL_STYLES = ["{Full Journal Name}", "{J. Abbr.}", "{J Abbr}"]
AVAILABLE_ELEMENTS = ["", "Authors", "Title", "Journal", "Year", "Volume", "Issue", "Pages", "DOI"]

# Хранение текущего языка
if 'current_language' not in st.session_state:
    st.session_state.current_language = 'en'

# Хранение импортированного стиля и флага применения
if 'imported_style' not in st.session_state:
    st.session_state.imported_style = None
if 'style_applied' not in st.session_state:
    st.session_state.style_applied = False

# Флаг для применения импортированного стиля после рендера
if 'apply_imported_style' not in st.session_state:
    st.session_state.apply_imported_style = False

# Для хранения результатов обработки
if 'output_text_value' not in st.session_state:
    st.session_state.output_text_value = ""
if 'show_results' not in st.session_state:
    st.session_state.show_results = False
if 'download_data' not in st.session_state:
    st.session_state.download_data = {}

# Для хранения состояния чекбоксов and/&
if 'use_and_checkbox' not in st.session_state:
    st.session_state.use_and_checkbox = False
if 'use_ampersand_checkbox' not in st.session_state:
    st.session_state.use_ampersand_checkbox = False

# Для хранения стиля журнала
if 'journal_style' not in st.session_state:
    st.session_state.journal_style = '{Full Journal Name}'

class JournalAbbreviation:
    def __init__(self):
        self.ltwa_data = {}
        self.load_ltwa_data()
        # Список аббревиатур, которые всегда пишутся с большой буквы
        self.uppercase_abbreviations = {
            'acs', 'ecs', 'rsc', 'ieee', 'iet', 'acm', 'aims', 'bmc', 'bmj', 'npj'
        }
    
    def load_ltwa_data(self):
        """Загружает данные сокращений из файла ltwa.csv"""
        try:
            csv_path = os.path.join(os.path.dirname(__file__), 'ltwa.csv')
            with open(csv_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f, delimiter='\t')
                next(reader)  # Пропускаем заголовок
                for row in reader:
                    if len(row) >= 2:
                        word = row[0].strip()
                        abbreviation = row[1].strip() if row[1].strip() else None
                        self.ltwa_data[word] = abbreviation
        except FileNotFoundError:
            print("Файл ltwa.csv не найден")
        except Exception as e:
            print(f"Ошибка загрузки ltwa.csv: {e}")
    
    def abbreviate_word(self, word: str) -> str:
        """Сокращает одно слово на основе данных LTWA"""
        word_lower = word.lower()
        
        # Проверяем точное совпадение
        if word_lower in self.ltwa_data:
            abbr = self.ltwa_data[word_lower]
            return abbr if abbr else word
        
        # Проверяем совпадение с дефисом (корневые слова)
        for ltwa_word, abbr in self.ltwa_data.items():
            if ltwa_word.endswith('-') and word_lower.startswith(ltwa_word[:-1]):
                return abbr if abbr else word
        
        return word
    
    def abbreviate_journal_name(self, journal_name: str, style: str = "{J. Abbr.}") -> str:
        """Сокращает название журнала в соответствии с выбранным стилем"""
        if not journal_name:
            return ""
        
        # Удаляем артикли, предлоги и двоеточия
        words_to_remove = {'a', 'an', 'the', 'of', 'in', 'and', '&'}
        words = [word for word in journal_name.split() if word.lower() not in words_to_remove]
        
        # Удаляем двоеточия из отдельных слов
        words = [word.replace(':', '') for word in words]
        
        # Если после удаления артиклей и предлогов осталось только одно слово - не сокращаем
        if len(words) <= 1:
            return journal_name
        
        # Сокращаем каждое слово
        abbreviated_words = []
        for i, word in enumerate(words):
            # Сохраняем регистр первой буквы
            original_first_char = word[0]
            abbreviated = self.abbreviate_word(word.lower())
            
            # Восстанавливаем регистр
            if abbreviated and original_first_char.isupper():
                abbreviated = abbreviated[0].upper() + abbreviated[1:]
            
            # Для первого слова проверяем, является ли оно аббревиатурой, которую нужно писать с большой буквы
            if i == 0 and abbreviated.lower() in self.uppercase_abbreviations:
                abbreviated = abbreviated.upper()
            
            abbreviated_words.append(abbreviated)
        
        # Формируем результат в зависимости от стиля
        if style == "{J. Abbr.}":
            # Аббревиатура с точками
            result = " ".join(abbreviated_words)
        elif style == "{J Abbr}":
            # Аббревиатура без точек
            result = " ".join(abbr.replace('.', '') for abbr in abbreviated_words)
        else:
            # Полное название
            result = journal_name
        
        # Убираем двойные точки
        result = re.sub(r'\.\.+', '.', result)
        
        return result

# Инициализация системы сокращений
journal_abbrev = JournalAbbreviation()

class BaseCitationFormatter:
    """Базовый класс для форматирования цитирования"""
    
    def __init__(self, style_config: Dict[str, Any]):
        self.style_config = style_config
    
    def format_authors(self, authors: List[Dict[str, str]]) -> str:
        """Форматирует список авторов"""
        if not authors:
            return ""
        
        author_format = self.style_config['author_format']
        separator = self.style_config['author_separator']
        et_al_limit = self.style_config['et_al_limit']
        use_and_bool = self.style_config['use_and_bool']
        use_ampersand_bool = self.style_config['use_ampersand_bool']
        
        author_str = ""
        
        # Определяем лимит авторов для отображения
        if use_and_bool or use_ampersand_bool:
            limit = len(authors)
        else:
            limit = et_al_limit if et_al_limit and et_al_limit > 0 else len(authors)
        
        for i, author in enumerate(authors[:limit]):
            given = author['given']
            family = author['family']
            
            # Извлекаем инициалы
            initials = given.split()[:2]
            first_initial = initials[0][0] if initials else ''
            second_initial = initials[1][0].upper() if len(initials) > 1 else ''
            
            # Форматируем автора в зависимости от выбранного формата
            if author_format == "AA Smith":
                formatted_author = f"{first_initial}{second_initial} {family}"
            elif author_format == "A.A. Smith":
                if second_initial:
                    formatted_author = f"{first_initial}.{second_initial}. {family}"
                else:
                    formatted_author = f"{first_initial}. {family}"
            elif author_format == "Smith AA":
                formatted_author = f"{family} {first_initial}{second_initial}"
            elif author_format == "Smith A.A":
                if second_initial:
                    formatted_author = f"{family} {first_initial}.{second_initial}."
                else:
                    formatted_author = f"{family} {first_initial}."
            elif author_format == "Smith, A.A.":
                if second_initial:
                    formatted_author = f"{family}, {first_initial}.{second_initial}."
                else:
                    formatted_author = f"{family}, {first_initial}."
            else:
                formatted_author = f"{first_initial}. {family}"
            
            author_str += formatted_author
            
            # Добавляем разделитель между авторов
            if i < len(authors[:limit]) - 1:
                if i == len(authors[:limit]) - 2 and (use_and_bool or use_ampersand_bool):
                    # Используем "and" или "&" в зависимости от выбора
                    if use_and_bool:
                        author_str += " and "
                    else:  # use_ampersand_bool
                        author_str += " & "
                else:
                    author_str += separator
        
        # Добавляем "et al" если нужно
        if et_al_limit and len(authors) > et_al_limit and not (use_and_bool or use_ampersand_bool):
            author_str += " et al"
        
        return author_str.strip()
    
    def format_pages(self, pages: str, article_number: str, style_type: str = "default") -> str:
        """Форматирует страницы в зависимости от стиля"""
        page_format = self.style_config['page_format']
        
        if pages:
            if style_type == "rsc":
                # Для RSC стиля берем только первую страницу
                if '-' in pages:
                    first_page = pages.split('-')[0].strip()
                    return first_page
                else:
                    return pages.strip()
            elif style_type == "cta":
                # Для стиля CTA сокращаем диапазон страниц (6441–6 вместо 6441–6446)
                if '-' in pages:
                    start, end = pages.split('-')
                    start = start.strip()
                    end = end.strip()
                    
                    # Сокращаем конечную страницу если возможно
                    if len(start) == len(end) and start[:-1] == end[:-1]:
                        return f"{start}–{end[-1]}"
                    elif len(start) > 1 and len(end) > 1 and start[:-2] == end[:-2]:
                        return f"{start}–{end[-2:]}"
                    else:
                        return f"{start}–{end}"
                else:
                    return pages.strip()
            else:
                # Для других стилей используем стандартное форматирование
                if '-' not in pages:
                    return pages
                
                start, end = pages.split('-')
                start = start.strip()
                end = end.strip()
                
                if page_format == "122 - 128":
                    return f"{start} - {end}"
                elif page_format == "122-128":
                    return f"{start}-{end}"
                elif page_format == "122 – 128":
                    return f"{start} – {end}"
                elif page_format == "122–128":
                    return f"{start}–{end}"
                elif page_format == "122–8":
                    i = 0
                    while i < len(start) and i < len(end) and start[i] == end[i]:
                        i += 1
                    return f"{start}–{end[i:]}"
        
        # Если страниц нет, используем номер статьи
        return article_number
    
    def format_doi(self, doi: str) -> Tuple[str, str]:
        """Форматирует DOI и возвращает текст и URL"""
        doi_format = self.style_config['doi_format']
        
        if doi_format == "10.10/xxx":
            value = doi
        elif doi_format == "doi:10.10/xxx":
            value = f"doi:{doi}"
        elif doi_format == "DOI:10.10/xxx":
            value = f"DOI:{doi}"
        elif doi_format == "https://dx.doi.org/10.10/xxx":
            value = f"https://dx.doi.org/{doi}"
        else:
            value = doi
        
        return value, f"https://doi.org/{doi}"
    
    def format_journal_name(self, journal_name: str) -> str:
        """Форматирует название журнала с учетом выбранного стиля"""
        journal_style = self.style_config.get('journal_style', '{Full Journal Name}')
        return journal_abbrev.abbreviate_journal_name(journal_name, journal_style)

class CustomCitationFormatter(BaseCitationFormatter):
    """Форматировщик для пользовательских стилей"""
    
    def format_reference(self, metadata: Dict[str, Any], for_preview: bool = False) -> Tuple[Any, bool]:
        """Форматирует ссылку в пользовательском стиле"""
        if not metadata:
            error_message = "Ошибка: Не удалось отформатировать ссылку." if st.session_state.current_language == 'ru' else "Error: Could not format the reference."
            return (error_message, True)
        
        elements = []
        
        for i, (element, config) in enumerate(self.style_config['elements']):
            value = ""
            doi_value = None
            
            if element == "Authors":
                value = self.format_authors(metadata['authors'])
            elif element == "Title":
                value = metadata['title']
            elif element == "Journal":
                value = self.format_journal_name(metadata['journal'])
            elif element == "Year":
                value = str(metadata['year']) if metadata['year'] else ""
            elif element == "Volume":
                value = metadata['volume']
            elif element == "Issue":
                value = metadata['issue']
            elif element == "Pages":
                value = self.format_pages(metadata['pages'], metadata['article_number'])
            elif element == "DOI":
                doi = metadata['doi']
                doi_value = doi
                value, _ = self.format_doi(doi)
            
            if value:
                # Добавляем скобки если нужно
                if config['parentheses'] and value:
                    value = f"({value})"
                
                # Добавляем разделитель
                separator = config['separator'] if i < len(self.style_config['elements']) - 1 else ''
                
                if for_preview:
                    # Для предпросмотра используем HTML-теги
                    formatted_value = value
                    if config['italic']:
                        formatted_value = f"<i>{formatted_value}</i>"
                    if config['bold']:
                        formatted_value = f"<b>{formatted_value}</b>"
                    
                    elements.append((formatted_value, False, False, separator, False, None))
                else:
                    # Для реального документа сохраняем информацию о форматировании
                    elements.append((value, config['italic'], config['bold'], separator,
                                   (element == "DOI" and self.style_config['doi_hyperlink']), doi_value))
        
        if for_preview:
            # Собираем строку для предпросмотра
            ref_str = ""
            for i, (value, _, _, separator, _, _) in enumerate(elements):
                ref_str += value
                if separator and i < len(elements) - 1:
                    ref_str += separator
                elif i == len(elements) - 1 and self.style_config['final_punctuation']:
                    ref_str = ref_str.rstrip(',.') + "."
            
            # Убираем двойные точки
            ref_str = re.sub(r'\.\.+', '.', ref_str)
            
            return ref_str, False
        else:
            return elements, False

class GOSTCitationFormatter(BaseCitationFormatter):
    """Форматировщик для стиля ГОСТ"""
    
    def format_reference(self, metadata: Dict[str, Any], for_preview: bool = False) -> Tuple[Any, bool]:
        """Форматирует ссылку по стандарту ГОСТ"""
        if not metadata:
            error_message = "Ошибка: Не удалось отформатировать ссылку." if st.session_state.current_language == 'ru' else "Error: Could not format the reference."
            return (error_message, True)
        
        # Форматируем первого автора для основной части
        first_author = ""
        if metadata['authors']:
            author = metadata['authors'][0]
            given = author['given']
            family = author['family']
            initials = given.split()[:2]
            first_initial = initials[0][0] if initials else ''
            second_initial = initials[1][0].upper() if len(initials) > 1 else ''
            
            if second_initial:
                first_author = f"{family}, {first_initial}.{second_initial}."
            else:
                first_author = f"{family}, {first_initial}."
        
        # Форматируем всех авторов для части после /
        all_authors = ""
        for i, author in enumerate(metadata['authors']):
            given = author['given']
            family = author['family']
            initials = given.split()[:2]
            first_initial = initials[0][0] if initials else ''
            second_initial = initials[1][0].upper() if len(initials) > 1 else ''
            
            if second_initial:
                author_str = f"{first_initial}.{second_initial}. {family}"
            else:
                author_str = f"{first_initial}. {family}"
            
            all_authors += author_str
            if i < len(metadata['authors']) - 1:
                all_authors += ", "
        
        # Форматируем страницы с использованием длинного тире вместо дефиса
        pages = metadata['pages']
        article_number = metadata['article_number']
        
        # Определяем язык и устанавливаем метки для томов/страниц/статей
        is_russian = st.session_state.current_language == 'ru'
        volume_label = "Т." if is_russian else "Vol."
        page_label = "С." if is_russian else "P."
        article_label = "Арт." if is_russian else "Art."
        issue_label = "№" if is_russian else "No."
        
        # Форматируем DOI
        doi_url = f"https://doi.org/{metadata['doi']}"
        
        # Для ГОСТ используем полное название журнала (без сокращений)
        journal_name = metadata['journal']
        
        # Строим ссылку ГОСТ с номером выпуска, если доступно
        if metadata['issue']:
            gost_ref = f"{first_author} {metadata['title']} / {all_authors} // {journal_name}. – {metadata['year']}. – {volume_label} {metadata['volume']}. – {issue_label} {metadata['issue']}."
        else:
            gost_ref = f"{first_author} {metadata['title']} / {all_authors} // {journal_name}. – {metadata['year']}. – {volume_label} {metadata['volume']}."
        
        # Добавляем страницы или номер статьи
        if pages:
            if '-' in pages:
                start_page, end_page = pages.split('-')
                pages = f"{start_page.strip()}–{end_page.strip()}"  # Используем длинное тире
            else:
                pages = pages.strip()
            gost_ref += f" – {page_label} {pages}."
        elif article_number:
            gost_ref += f" – {article_label} {article_number}."
        else:
            if is_russian:
                gost_ref += " – [Без пагинации]."
            else:
                gost_ref += " – [No pagination]."
        
        # Добавляем DOI
        gost_ref += f" – {doi_url}"
        
        if for_preview:
            return gost_ref, False
        else:
            # Для реального документа возвращаем как несколько элементов с DOI как гиперссылкой
            elements = []
            
            # Добавляем весь текст до DOI как обычный текст
            text_before_doi = gost_ref.replace(doi_url, "")
            elements.append((text_before_doi, False, False, "", False, None))
            
            # Добавляем DOI как гиперссылку
            elements.append((doi_url, False, False, "", True, metadata['doi']))
            
            return elements, False

class ACSCitationFormatter(BaseCitationFormatter):
    """Форматировщик для стиля ACS (MDPI)"""
    
    def format_reference(self, metadata: Dict[str, Any], for_preview: bool = False) -> Tuple[Any, bool]:
        """Форматирует ссылку в стиле ACS (MDPI)"""
        if not metadata:
            error_message = "Ошибка: Не удалось отформатировать ссылку." if st.session_state.current_language == 'ru' else "Error: Could not format the reference."
            return (error_message, True)
        
        # Форматируем авторов в стиле ACS: Surname, I.I.; Surname, I.I.; ...
        authors_str = ""
        for i, author in enumerate(metadata['authors']):
            given = author['given']
            family = author['family']
            
            # Извлекаем инициалы
            initials = given.split()[:2]
            first_initial = initials[0][0] if initials else ''
            second_initial = initials[1][0].upper() if len(initials) > 1 else ''
            
            # Форматируем автора: Surname, I.I.
            if second_initial:
                author_str = f"{family}, {first_initial}.{second_initial}."
            else:
                author_str = f"{family}, {first_initial}."
            
            authors_str += author_str
            
            # Добавляем разделитель
            if i < len(metadata['authors']) - 1:
                authors_str += "; "
        
        # Форматируем страницы
        pages = metadata['pages']
        article_number = metadata['article_number']
        
        if pages:
            if '-' in pages:
                start_page, end_page = pages.split('-')
                start_page = start_page.strip()
                end_page = end_page.strip()
                # Используем короткий формат для конечной страницы если возможно
                if len(start_page) == len(end_page) and start_page[:-1] == end_page[:-1]:
                    pages_formatted = f"{start_page}−{end_page[-1]}"
                else:
                    pages_formatted = f"{start_page}−{end_page}"
            else:
                pages_formatted = pages
        elif article_number:
            pages_formatted = article_number
        else:
            pages_formatted = ""
        
        # Применяем сокращение названия журнала для стиля ACS
        journal_name = self.format_journal_name(metadata['journal'])
        
        # Собираем ссылку ACS
        acs_ref = f"{authors_str} {metadata['title']}. {journal_name} {metadata['year']}, {metadata['volume']}, {pages_formatted}."
        
        # Убираем двойные точки
        acs_ref = re.sub(r'\.\.+', '.', acs_ref)
        
        if for_preview:
            return acs_ref, False
        else:
            # Для реального документа разбиваем на элементы с форматированием
            elements = []
            
            # Авторы
            elements.append((authors_str, False, False, " ", False, None))
            
            # Название
            elements.append((metadata['title'], False, False, ". ", False, None))
            
            # Журнал (курсив)
            elements.append((journal_name, True, False, " ", False, None))
            
            # Год (жирный)
            elements.append((str(metadata['year']), False, True, ", ", False, None))
            
            # Том (курсив)
            elements.append((metadata['volume'], True, False, ", ", False, None))
            
            # Страницы
            elements.append((pages_formatted, False, False, ".", False, None))
            
            return elements, False

class RSCCitationFormatter(BaseCitationFormatter):
    """Форматировщик для стиля RSC"""
    
    def format_reference(self, metadata: Dict[str, Any], for_preview: bool = False) -> Tuple[Any, bool]:
        """Форматирует ссылку в стиле RSC"""
        if not metadata:
            error_message = "Ошибка: Не удалось отформатировать ссылку." if st.session_state.current_language == 'ru' else "Error: Could not format the reference."
            return (error_message, True)
        
        # Форматируем авторов в стиле RSC: I.I. Surname, I.I. Surname, ... and I.I. Surname
        authors_str = ""
        for i, author in enumerate(metadata['authors']):
            given = author['given']
            family = author['family']
            
            # Извлекаем инициалы
            initials = given.split()[:2]
            first_initial = initials[0][0] if initials else ''
            second_initial = initials[1][0].upper() if len(initials) > 1 else ''
            
            # Форматируем автора: I.I. Surname
            if second_initial:
                author_str = f"{first_initial}.{second_initial}. {family}"
            else:
                author_str = f"{first_initial}. {family}"
            
            authors_str += author_str
            
            # Добавляем разделитель
            if i < len(metadata['authors']) - 1:
                if i == len(metadata['authors']) - 2:
                    authors_str += " and "
                else:
                    authors_str += ", "
        
        # Форматируем страницы - для RSC берем только первую страницу
        pages = metadata['pages']
        article_number = metadata['article_number']
        
        if pages:
            # Для RSC стиля берем только первую страницу
            if '-' in pages:
                first_page = pages.split('-')[0].strip()
                pages_formatted = first_page
            else:
                pages_formatted = pages.strip()
        elif article_number:
            pages_formatted = article_number
        else:
            pages_formatted = ""
        
        # Применяем сокращение названия журнала для стиля RSC
        journal_name = self.format_journal_name(metadata['journal'])
        
        # Собираем ссылку RSC
        rsc_ref = f"{authors_str}, {journal_name}, {metadata['year']}, {metadata['volume']}, {pages_formatted}."
        
        # Убираем двойные точки
        rsc_ref = re.sub(r'\.\.+', '.', rsc_ref)
        
        if for_preview:
            return rsc_ref, False
        else:
            # Для реального документа разбиваем на элементы с форматированием
            elements = []
            
            # Авторы
            elements.append((authors_str, False, False, ", ", False, None))
            
            # Журнал (курсив)
            elements.append((journal_name, True, False, ", ", False, None))
            
            # Год
            elements.append((str(metadata['year']), False, False, ", ", False, None))
            
            # Том (жирный)
            elements.append((metadata['volume'], False, True, ", ", False, None))
            
            # Страницы (только первая страница)
            elements.append((pages_formatted, False, False, ".", False, None))
            
            return elements, False

class CTACitationFormatter(BaseCitationFormatter):
    """Форматировщик для стиля CTA"""
    
    def format_reference(self, metadata: Dict[str, Any], for_preview: bool = False) -> Tuple[Any, bool]:
        """Форматирует ссылку в стиле CTA"""
        if not metadata:
            error_message = "Ошибка: Не удалось отформатировать ссылку." if st.session_state.current_language == 'ru' else "Error: Could not format the reference."
            return (error_message, True)
        
        # Форматируем авторов в стиле CTA: Surname Initials, Surname Initials, ... Surname Initials
        authors_str = ""
        for i, author in enumerate(metadata['authors']):
            given = author['given']
            family = author['family']
            
            # Извлекаем инициалы
            initials = given.split()[:2]
            first_initial = initials[0][0] if initials else ''
            second_initial = initials[1][0].upper() if len(initials) > 1 else ''
            
            # Форматируем автора: Surname Initials (без точек)
            if second_initial:
                author_str = f"{family} {first_initial}{second_initial}"
            else:
                author_str = f"{family} {first_initial}"
            
            authors_str += author_str
            
            # Добавляем разделитель
            if i < len(metadata['authors']) - 1:
                authors_str += ", "
        
        # Форматируем страницы для стиля CTA (сокращаем диапазон)
        pages = metadata['pages']
        article_number = metadata['article_number']
        pages_formatted = self.format_pages(pages, article_number, "cta")
        
        # Применяем сокращение названия журнала для стиля CTA (без точек)
        journal_name = self.format_journal_name(metadata['journal'])
        
        # Форматируем номер выпуска если есть
        issue_part = f"({metadata['issue']})" if metadata['issue'] else ""
        
        # Собираем ссылку CTA
        cta_ref = f"{authors_str}. {metadata['title']}. {journal_name}. {metadata['year']};{metadata['volume']}{issue_part}:{pages_formatted}. doi:{metadata['doi']}"
        
        if for_preview:
            return cta_ref, False
        else:
            # Для реального документа разбиваем на элементы с форматированием
            elements = []
            
            # Авторы
            elements.append((authors_str, False, False, ". ", False, None))
            
            # Название
            elements.append((metadata['title'], False, False, ". ", False, None))
            
            # Журнал (курсив)
            elements.append((journal_name, True, False, ". ", False, None))
            
            # Год
            elements.append((str(metadata['year']), False, False, ";", False, None))
            
            # Том
            elements.append((metadata['volume'], False, False, "", False, None))
            
            # Номер выпуска (если есть)
            if metadata['issue']:
                elements.append((f"({metadata['issue']})", False, False, ":", False, None))
            else:
                elements.append(("", False, False, ":", False, None))
            
            # Страницы
            elements.append((pages_formatted, False, False, ". ", False, None))
            
            # DOI - всегда как гиперссылка в стиле CTA
            doi_text = f"doi:{metadata['doi']}"
            elements.append((doi_text, False, False, "", True, metadata['doi']))
            
            return elements, False

class CitationFormatterFactory:
    """Фабрика для создания форматировщиков цитирования"""
    
    @staticmethod
    def create_formatter(style_config: Dict[str, Any]) -> BaseCitationFormatter:
        """Создает соответствующий форматировщик на основе конфигурации стиля"""
        if style_config.get('gost_style', False):
            return GOSTCitationFormatter(style_config)
        elif style_config.get('acs_style', False):
            return ACSCitationFormatter(style_config)
        elif style_config.get('rsc_style', False):
            return RSCCitationFormatter(style_config)
        elif style_config.get('cta_style', False):
            return CTACitationFormatter(style_config)
        else:
            return CustomCitationFormatter(style_config)

class DocumentGenerator:
    """Класс для генерации DOCX документов"""
    
    @staticmethod
    def add_hyperlink(paragraph, text, url):
        """Добавляет гиперссылку в параграф"""
        part = paragraph.part
        r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
        
        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('r:id'), r_id)
        
        new_run = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        
        # Синий цвет для гиперссылки
        color = OxmlElement('w:color')
        color.set(qn('w:val'), '0000FF')
        rPr.append(color)
        
        # Подчеркивание
        underline = OxmlElement('w:u')
        underline.set(qn('w:val'), 'single')
        rPr.append(underline)
        
        new_run.append(rPr)
        new_text = OxmlElement('w:t')
        new_text.text = text
        new_run.append(new_text)
        
        hyperlink.append(new_run)
        paragraph._p.append(hyperlink)
        
        return hyperlink
    
    @staticmethod
    def apply_yellow_background(run):
        """Применяет желтый фон к тексту"""
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), 'FFFF00')
        run._element.get_or_add_rPr().append(shd)
    
    @staticmethod
    def apply_blue_background(run):
        """Применяет синий фон к тексту"""
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), 'E6F3FF')  # Светло-синий цвет
        run._element.get_or_add_rPr().append(shd)
    
    @staticmethod
    def apply_red_color(run):
        """Применяет красный цвет к тексту"""
        color = OxmlElement('w:color')
        color.set(qn('w:val'), 'FF0000')
        run._element.get_or_add_rPr().append(color)
    
    @staticmethod
    def generate_document(formatted_refs: List[Tuple[Any, bool, Any]], 
                         statistics: Dict[str, Any],
                         style_config: Dict[str, Any],
                         duplicates_info: Dict[int, int] = None) -> io.BytesIO:
        """Генерирует DOCX документ с отформатированными ссылками и статистикой"""
        output_doc = Document()
        
        # Измененный заголовок согласно требованию 1 и 4
        output_doc.add_paragraph('Citation Style Construction / developed by daM©')
        output_doc.add_paragraph('See short stats after the References section')
        output_doc.add_heading('References', level=1)
        
        # Добавляем форматированные ссылки
        DocumentGenerator._add_formatted_references(output_doc, formatted_refs, style_config, duplicates_info)
        
        # Добавляем раздел Stats согласно требованию 5
        DocumentGenerator._add_statistics_section(output_doc, statistics)
        
        output_doc_buffer = io.BytesIO()
        output_doc.save(output_doc_buffer)
        output_doc_buffer.seek(0)
        
        return output_doc_buffer
    
    @staticmethod
    def _add_formatted_references(doc: Document, 
                                formatted_refs: List[Tuple[Any, bool, Any]], 
                                style_config: Dict[str, Any],
                                duplicates_info: Dict[int, int] = None):
        """Добавляет форматированные ссылки в документ"""
        for i, (elements, is_error, metadata) in enumerate(formatted_refs):
            numbering = style_config['numbering_style']
            
            # Формируем префикс нумерации
            if numbering == "No numbering":
                prefix = ""
            elif numbering == "1":
                prefix = f"{i + 1} "
            elif numbering == "1.":
                prefix = f"{i + 1}. "
            elif numbering == "1)":
                prefix = f"{i + 1}) "
            elif numbering == "(1)":
                prefix = f"({i + 1}) "
            elif numbering == "[1]":
                prefix = f"[{i + 1}] "
            else:
                prefix = f"{i + 1}. "
            
            para = doc.add_paragraph(prefix)
            
            if is_error:
                # Показываем оригинальный текст с желтым фоном и сообщением об ошибки
                run = para.add_run(str(elements))
                DocumentGenerator.apply_yellow_background(run)
            elif duplicates_info and i in duplicates_info:
                # Дубликат - выделяем синим и добавляем пометку
                original_index = duplicates_info[i] + 1  # +1 потому что нумерация с 1
                duplicate_note = get_text('duplicate_reference').format(original_index)
                
                if isinstance(elements, str):
                    run = para.add_run(elements)
                    DocumentGenerator.apply_blue_background(run)
                    para.add_run(f" - {duplicate_note}").italic = True
                else:
                    for j, (value, italic, bold, separator, is_doi_hyperlink, doi_value) in enumerate(elements):
                        if is_doi_hyperlink and doi_value:
                            DocumentGenerator.add_hyperlink(para, value, f"https://doi.org/{doi_value}")
                        else:
                            run = para.add_run(value)
                            if italic:
                                run.font.italic = True
                            if bold:
                                run.font.bold = True
                            DocumentGenerator.apply_blue_background(run)
                        
                        if separator and j < len(elements) - 1:
                            para.add_run(separator)
                    
                    para.add_run(f" - {duplicate_note}").italic = True
            else:
                # Обычная ссылка
                if metadata is None:
                    run = para.add_run(str(elements))
                    run.font.italic = True
                else:
                    for j, (value, italic, bold, separator, is_doi_hyperlink, doi_value) in enumerate(elements):
                        if is_doi_hyperlink and doi_value:
                            # Добавляем DOI как гиперссылку
                            DocumentGenerator.add_hyperlink(para, value, f"https://doi.org/{doi_value}")
                        else:
                            run = para.add_run(value)
                            if italic:
                                run.font.italic = True
                            if bold:
                                run.font.bold = True
                        
                        # Добавляем разделитель между элементами
                        if separator and j < len(elements) - 1:
                            para.add_run(separator)
                    
                    # Добавляем конечную пунктуацию
                    if style_config['final_punctuation'] and not is_error:
                        para.add_run(".")
    
    @staticmethod
    def _add_statistics_section(doc: Document, statistics: Dict[str, Any]):
        """Добавляет раздел статистики в документ"""
        doc.add_heading('Stats', level=1)
        
        # Таблица Journal Frequency
        doc.add_heading('Journal Frequency', level=2)
        journal_table = doc.add_table(rows=1, cols=3)
        journal_table.style = 'Table Grid'
        
        # Заголовки таблицы
        hdr_cells = journal_table.rows[0].cells
        hdr_cells[0].text = 'Journal Name'
        hdr_cells[1].text = 'Count'
        hdr_cells[2].text = 'Percentage (%)'
        
        # Данные таблицы
        for journal_stat in statistics['journal_stats']:
            row_cells = journal_table.add_row().cells
            row_cells[0].text = journal_stat['journal']
            row_cells[1].text = str(journal_stat['count'])
            row_cells[2].text = str(journal_stat['percentage'])
        
        # Пустая строка между таблицами
        doc.add_paragraph()
        
        # Таблица Year Distribution
        doc.add_heading('Year Distribution', level=2)
        
        # Добавляем предупреждение если нужно согласно требованию 6
        if statistics['needs_more_recent_references']:
            warning_para = doc.add_paragraph()
            warning_run = warning_para.add_run("To improve the relevance and significance of the research, consider including more recent references published within the last 3-4 years")
            DocumentGenerator.apply_red_color(warning_run)
            doc.add_paragraph()
        
        year_table = doc.add_table(rows=1, cols=3)
        year_table.style = 'Table Grid'
        
        # Заголовки таблицы
        hdr_cells = year_table.rows[0].cells
        hdr_cells[0].text = 'Year'
        hdr_cells[1].text = 'Count'
        hdr_cells[2].text = 'Percentage (%)'
        
        # Данные таблицы
        for year_stat in statistics['year_stats']:
            row_cells = year_table.add_row().cells
            row_cells[0].text = str(year_stat['year'])
            row_cells[1].text = str(year_stat['count'])
            row_cells[2].text = str(year_stat['percentage'])
        
        # Пустая строка между таблицами
        doc.add_paragraph()
        
        # Таблица Author Distribution
        doc.add_heading('Author Distribution', level=2)
        
        # Добавляем предупреждение если нужно согласно требованию 7
        if statistics['has_frequent_author']:
            warning_para = doc.add_paragraph()
            warning_run = warning_para.add_run("The author(s) are referenced frequently. Either reduce the number of references to the author(s), or expand the reference list to include more sources")
            DocumentGenerator.apply_red_color(warning_run)
            doc.add_paragraph()
        
        author_table = doc.add_table(rows=1, cols=3)
        author_table.style = 'Table Grid'
        
        # Заголовки таблицы
        hdr_cells = author_table.rows[0].cells
        hdr_cells[0].text = 'Author'
        hdr_cells[1].text = 'Count'
        hdr_cells[2].text = 'Percentage (%)'
        
        # Данные таблицы
        for author_stat in statistics['author_stats']:
            row_cells = author_table.add_row().cells
            row_cells[0].text = author_stat['author']
            row_cells[1].text = str(author_stat['count'])
            row_cells[2].text = str(author_stat['percentage'])

def get_text(key):
    return TRANSLATIONS[st.session_state.current_language].get(key, key)

def clean_text(text):
    """Очищает текст от HTML тегов и entities"""
    if not text:
        return ""
    
    # Сначала убираем HTML теги, включая sub и sup
    text = re.sub(r'<[^>]+>', '', text)
    
    # Затем декодируем HTML entities
    text = html.unescape(text)
    
    # Убираем оставшиеся XML/HTML entities
    text = re.sub(r'&[^;]+;', '', text)
    
    return text.strip()

def normalize_name(name):
    """Нормализует имя автора с учетом составных фамилий"""
    if not name:
        return ''
    
    # Обрабатываем составные фамилии с дефисами, апострофами и другими разделителями
    if '-' in name or "'" in name or '’' in name:
        # Разбиваем на части по дефисам и апострофам
        parts = re.split(r'([-\'’])', name)
        normalized_parts = []
        
        for i, part in enumerate(parts):
            if part in ['-', "'", '’']:
                normalized_parts.append(part)
            else:
                if part:
                    # Каждую часть имени пишем с большой буквы
                    normalized_parts.append(part[0].upper() + part[1:].lower() if len(part) > 1 else part.upper())
        
        return ''.join(normalized_parts)
    else:
        # Обычное имя
        if len(name) > 1:
            return name[0].upper() + name[1:].lower()
        else:
            return name.upper()

def is_section_header(text):
    """Определяет, является ли текст заголовком раздела"""
    text_upper = text.upper().strip()
    
    # Проверяем только явные паттерны заголовков
    section_patterns = [
        r'^NOTES?\s+AND\s+REFERENCES?$',
        r'^REFERENCES?$',
        r'^BIBLIOGRAPHY$',
        r'^LITERATURE$',
        r'^WORKS?\s+CITED$',
        r'^SOURCES?$',
        r'^CHAPTER\s+\d+$',
        r'^SECTION\s+\d+$',
        r'^PART\s+\d+$'
    ]
    
    for pattern in section_patterns:
        if re.search(pattern, text_upper):
            return True
    
    # Убираем слишком агрессивную проверку на короткие строки
    # DOI могут быть короткими, но это не заголовки
    return False

def find_doi(reference):
    """Находит DOI в строке ссылки"""
    if is_section_header(reference):
        return None
    
    # Улучшенные паттерны для поиска DOI
    doi_patterns = [
        r'https?://doi\.org/(10\.\d{4,9}/[-._;()/:A-Za-z0-9]+)',  # https://doi.org/10.xxx/xxx
        r'doi:\s*(10\.\d{4,9}/[-._;()/:A-Za-z0-9]+)',             # doi:10.xxx/xxx
        r'DOI:\s*(10\.\d{4,9}/[-._;()/:A-Za-z0-9]+)',             # DOI:10.xxx/xxx
        r'\b(10\.\d{4,9}/[-._;()/:A-Za-z0-9]+)\b'                 # 10.xxx/xxx (просто DOI)
    ]
    
    for pattern in doi_patterns:
        match = re.search(pattern, reference, re.IGNORECASE)
        if match:
            doi = match.group(1)
            # Убираем только конечные точки и запятые
            doi = doi.rstrip('.,;:')
            return doi
    
    # Если строка содержит только DOI (без другого текста)
    clean_ref = reference.strip()
    if re.match(r'^(doi:|DOI:)?\s*10\.\d{4,9}/[-._;()/:A-Za-z0-9]+\s*$', clean_ref, re.IGNORECASE):
        doi_match = re.search(r'(10\.\d{4,9}/[-._;()/:A-Za-z0-9]+)', clean_ref)
        if doi_match:
            doi = doi_match.group(1).rstrip('.,;:')
            return doi
    
    # ВАЖНЫЙ БЛОК: Если DOI не найден в явном виде, попробуем найти по библиографическим данным
    clean_ref = re.sub(r'\s*(https?://doi\.org/|doi:|DOI:)\s*[^\s,;]+', '', reference, flags=re.IGNORECASE)
    clean_ref = clean_ref.strip()
    
    if len(clean_ref) < 30:
        return None
    
    try:
        # Используем Crossref API для поиска по библиографическим данным
        query = works.query(bibliographic=clean_ref).sort('relevance').order('desc')
        for result in query:
            if 'DOI' in result:
                return result['DOI']
    except Exception as e:
        print(f"Error in bibliographic search for '{clean_ref}': {e}")
        return None
    
    return None

def normalize_doi(doi):
    """Нормализует DOI к стандартному формату"""
    if not doi:
        return ""
    # Убираем префиксы и приводим к нижнему регистру
    doi = re.sub(r'^(https?://doi\.org/|doi:|DOI:)', '', doi, flags=re.IGNORECASE)
    return doi.lower().strip()

def generate_reference_hash(metadata):
    """Генерирует хеш для идентификации дубликатов ссылок"""
    if not metadata:
        return None
    
    # Создаем строку для хеширования из основных полей
    hash_string = ""
    
    # Авторы (только фамилии в нижнем регистре)
    if metadata.get('authors'):
        authors_hash = "|".join(sorted([author.get('family', '').lower() for author in metadata['authors']]))
        hash_string += authors_hash + "||"
    
    # Название (первые 50 символов в нижнем регистре)
    title = metadata.get('title', '')[:50].lower()
    hash_string += title + "||"
    
    # Журнал и год
    hash_string += (metadata.get('journal', '') + "||").lower()
    hash_string += str(metadata.get('year', '')) + "||"
    
    # Том и страницы
    hash_string += metadata.get('volume', '') + "||"
    hash_string += metadata.get('pages', '') + "||"
    
    # DOI (если есть)
    hash_string += normalize_doi(metadata.get('doi', ''))
    
    # Создаем MD5 хеш
    return hashlib.md5(hash_string.encode('utf-8')).hexdigest()

def extract_metadata_batch(doi_list, progress_callback=None):
    """Пакетное извлечение метаданных через Crossref API с повторными попытками"""
    if not doi_list:
        return []
    
    results = [None] * len(doi_list)
    
    # Первая попытка - пакетная обработка с ThreadPoolExecutor
    with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
        future_to_index = {executor.submit(extract_metadata_sync, doi): i for i, doi in enumerate(doi_list)}
        
        completed = 0
        for future in concurrent.futures.as_completed(future_to_index):
            index = future_to_index[future]
            try:
                result = future.result()
                results[index] = result
            except Exception as e:
                print(f"Error processing DOI at index {index}: {e}")
                results[index] = None
            
            completed += 1
            if progress_callback:
                progress_callback(completed, len(doi_list))
    
    # Вторая попытка - повтор для неудачных запросов
    failed_indices = [i for i, result in enumerate(results) if result is None]
    if failed_indices:
        print(f"Retrying {len(failed_indices)} failed DOI requests...")
        
        if progress_callback:
            progress_callback(len(doi_list) - len(failed_indices), len(doi_list), retry_mode=True)
        
        # Более медленная повторная попытка с меньшим количеством потоков
        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
            retry_futures = {}
            for index in failed_indices:
                doi = doi_list[index]
                future = executor.submit(extract_metadata_sync, doi)
                retry_futures[future] = index
            
            for future in concurrent.futures.as_completed(retry_futures):
                index = retry_futures[future]
                try:
                    result = future.result()
                    results[index] = result
                except Exception as e:
                    print(f"Error in retry processing DOI at index {index}: {e}")
                    results[index] = None
                
                completed = len(doi_list) - len([r for r in results if r is None])
                if progress_callback:
                    progress_callback(completed, len(doi_list), retry_mode=True)
    
    return results

def extract_metadata_sync(doi):
    """Синхронная версия извлечения метаданных"""
    try:
        result = works.doi(doi)
        if not result:
            return None
        
        authors = result.get('author', [])
        author_list = []
        for author in authors:
            given_name = author.get('given', '')
            family_name = normalize_name(author.get('family', ''))
            author_list.append({
                'given': given_name,
                'family': family_name
            })
        
        title = ''
        if 'title' in result and result['title']:
            title = clean_text(result['title'][0])
            # Дополнительная очистка от тегов sub, i, SUB
            title = re.sub(r'</?sub>|</?i>|</?SUB>|</?I>', '', title, flags=re.IGNORECASE)
        
        journal = ''
        if 'container-title' in result and result['container-title']:
            journal = clean_text(result['container-title'][0])
        
        year = None
        if 'published' in result and 'date-parts' in result['published']:
            date_parts = result['published']['date-parts']
            if date_parts and date_parts[0]:
                year = date_parts[0][0]
        
        volume = result.get('volume', '')
        issue = result.get('issue', '')
        pages = result.get('page', '')
        article_number = result.get('article-number', '')
        
        metadata = {
            'authors': author_list,
            'title': title,
            'journal': journal,
            'year': year,
            'volume': volume,
            'issue': issue,
            'pages': pages,
            'article_number': article_number,
            'doi': doi,
            'original_doi': doi
        }
        
        return metadata
        
    except Exception as e:
        print(f"Error extracting metadata for DOI {doi}: {e}")
        return None

def format_reference(metadata, style_config, for_preview=False):
    """Универсальная функция форматирования ссылки"""
    formatter = CitationFormatterFactory.create_formatter(style_config)
    return formatter.format_reference(metadata, for_preview)

def find_duplicate_references(formatted_refs):
    """Находит дубликаты ссылок и возвращает информацию о них"""
    seen_hashes = {}
    duplicates_info = {}
    
    for i, (elements, is_error, metadata) in enumerate(formatted_refs):
        if is_error or not metadata:
            continue
            
        ref_hash = generate_reference_hash(metadata)
        if not ref_hash:
            continue
            
        if ref_hash in seen_hashes:
            # Найден дубликат
            original_index = seen_hashes[ref_hash]
            duplicates_info[i] = original_index
        else:
            # Первое вхождение
            seen_hashes[ref_hash] = i
    
    return duplicates_info

def generate_statistics(formatted_refs):
    """Генерирует статистику по ссылкам"""
    # Собираем данные
    journals = []
    years = []
    authors = []
    
    current_year = datetime.now().year
    
    for _, _, metadata in formatted_refs:
        if not metadata:
            continue
            
        # Журналы
        if metadata.get('journal'):
            journals.append(metadata['journal'])
        
        # Годы
        if metadata.get('year'):
            years.append(metadata['year'])
        
        # Авторы
        if metadata.get('authors'):
            for author in metadata['authors']:
                given = author.get('given', '')
                family = author.get('family', '')
                if family:
                    # Форматируем автора как "Surname FirstInitial"
                    first_initial = given[0] if given else ''
                    author_formatted = f"{family} {first_initial}." if first_initial else family
                    authors.append(author_formatted)
    
    # Уникальные DOI (без дубликатов)
    unique_dois = set()
    for _, _, metadata in formatted_refs:
        if metadata and metadata.get('doi'):
            unique_dois.add(metadata['doi'])
    
    total_unique_dois = len(unique_dois)
    
    # Статистика журналов
    journal_counter = Counter(journals)
    journal_stats = []
    for journal, count in journal_counter.most_common(20):
        percentage = (count / total_unique_dois) * 100 if total_unique_dois > 0 else 0
        journal_stats.append({
            'journal': journal,
            'count': count,
            'percentage': round(percentage, 2)
        })
    
    # Статистика годов
    year_counter = Counter(years)
    year_stats = []
    # Сортируем годы от текущего к 2010
    for year in range(current_year, 2009, -1):
        if year in year_counter:
            count = year_counter[year]
            percentage = (count / total_unique_dois) * 100 if total_unique_dois > 0 else 0
            year_stats.append({
                'year': year,
                'count': count,
                'percentage': round(percentage, 2)
            })
    
    # Проверка актуальности (последние 4 года)
    recent_years = [current_year - i for i in range(4)]
    recent_count = sum(year_counter.get(year, 0) for year in recent_years)
    recent_percentage = (recent_count / total_unique_dois) * 100 if total_unique_dois > 0 else 0
    needs_more_recent_references = recent_percentage < 20
    
    # Статистика авторов
    author_counter = Counter(authors)
    author_stats = []
    for author, count in author_counter.most_common(20):
        percentage = (count / total_unique_dois) * 100 if total_unique_dois > 0 else 0
        author_stats.append({
            'author': author,
            'count': count,
            'percentage': round(percentage, 2)
        })
    
    # Проверка частоты авторов
    has_frequent_author = any(stats['percentage'] > 30 for stats in author_stats)
    
    return {
        'journal_stats': journal_stats,
        'year_stats': year_stats,
        'author_stats': author_stats,
        'total_unique_dois': total_unique_dois,
        'needs_more_recent_references': needs_more_recent_references,
        'has_frequent_author': has_frequent_author
    }

def process_references_with_progress(references, style_config, progress_container, status_container):
    """Обрабатывает список ссылок с отображением прогресса"""
    doi_list = []
    formatted_refs = []
    doi_found_count = 0
    doi_not_found_count = 0
    
    # Собираем все DOI для пакетной обработки
    valid_dois = []
    reference_doi_map = {}  # Сопоставление индекса ссылки с DOI
    
    for i, ref in enumerate(references):
        if is_section_header(ref):
            doi_list.append(f"{ref} [SECTION HEADER - SKIPPED]")
            formatted_refs.append((ref, False, None))
            continue
            
        doi = find_doi(ref)
        if doi:
            valid_dois.append(doi)
            reference_doi_map[i] = doi
            doi_list.append(doi)
        else:
            doi_list.append(f"{ref}\nПроверьте источник и добавьте DOI вручную." if st.session_state.current_language == 'ru' else f"{ref}\nPlease check this source and insert the DOI manually.")
            error_message = f"{ref} [ОШИБКА: DOI не найден. Проверьте ссылку вручную.]" if st.session_state.current_language == 'ru' else f"{ref} [ERROR: DOI not found. Please check reference manually.]"
            formatted_refs.append((error_message, True, None))
            doi_not_found_count += 1
    
    # Пакетная обработка DOI
    if valid_dois:
        status_container.info(get_text('batch_processing'))
        
        # Создаем прогресс-бар для пакетной обработки
        batch_progress_bar = progress_container.progress(0)
        batch_status = status_container.empty()
        
        def update_batch_progress(completed, total, retry_mode=False):
            progress = completed / total
            batch_progress_bar.progress(progress)
            if retry_mode:
                batch_status.text(f"{get_text('retrying_failed')} {completed}/{total}")
            else:
                batch_status.text(f"{get_text('extracting_metadata')} {completed}/{total}")
        
        # Запускаем пакетную обработку с ThreadPoolExecutor
        metadata_results = extract_metadata_batch(valid_dois, update_batch_progress)
        
        # Обрабатываем результаты
        doi_to_metadata = dict(zip(valid_dois, metadata_results))
        
        for i, ref in enumerate(references):
            if i in reference_doi_map:
                doi = reference_doi_map[i]
                metadata = doi_to_metadata.get(doi)
                
                if metadata:
                    formatted_ref, is_error = format_reference(metadata, style_config)
                    formatted_refs.append((formatted_ref, is_error, metadata))
                    
                    if not is_error:
                        doi_found_count += 1
                    else:
                        error_message = f"{ref} [ОШИБКА: Не удалось отформатировать ссылку. Проверьте DOI вручную.]" if st.session_state.current_language == 'ru' else f"{ref} [ERROR: Could not format reference. Please check DOI manually.]"
                        doi_list[doi_list.index(doi)] = f"{doi}\nПроверьте источник и добавьте DOI вручную." if st.session_state.current_language == 'ru' else f"{doi}\nPlease check this source and insert the DOI manually."
                        formatted_refs.append((error_message, True, None))
                        doi_not_found_count += 1
                else:
                    error_message = f"{ref} [ОШИБКА: Не удалось получить метаданные по DOI. Проверьте DOI вручную.]" if st.session_state.current_language == 'ru' else f"{ref} [ERROR: Could not get metadata for DOI. Please check DOI manually.]"
                    doi_list[doi_list.index(doi)] = f"{doi}\nПроверьте источник и добавьте DOI вручную." if st.session_state.current_language == 'ru' else f"{doi}\nPlease check this source and insert the DOI manually."
                    formatted_refs.append((error_message, True, None))
                    doi_not_found_count += 1
    
    # Поиск дубликатов
    status_container.info(get_text('checking_duplicates'))
    duplicates_info = find_duplicate_references(formatted_refs)
    
    # Создаем TXT файл со списком DOI
    output_txt_buffer = io.StringIO()
    for doi in doi_list:
        output_txt_buffer.write(f"{doi}\n")
    output_txt_buffer.seek(0)
    txt_bytes = io.BytesIO(output_txt_buffer.getvalue().encode('utf-8'))
    
    return formatted_refs, txt_bytes, doi_found_count, doi_not_found_count, duplicates_info

def process_docx(input_file, style_config, progress_container, status_container):
    """Обрабатывает DOCX файл с ссылками с прогрессом"""
    doc = Document(input_file)
    references = []
    
    for para in doc.paragraphs:
        if para.text.strip():
            references.append(para.text.strip())
    
    st.write(f"**{get_text('found_references').format(len(references))}**")
    
    # Обрабатываем все ссылки с прогрессом
    formatted_refs, txt_bytes, doi_found_count, doi_not_found_count, duplicates_info = process_references_with_progress(
        references, style_config, progress_container, status_container
    )
    
    # Генерируем статистику
    statistics = generate_statistics(formatted_refs)
    
    # Создаем новый DOCX документ с отформатированными ссылками
    output_doc_buffer = DocumentGenerator.generate_document(
        formatted_refs, statistics, style_config, duplicates_info
    )
    
    return formatted_refs, txt_bytes, output_doc_buffer, doi_found_count, doi_not_found_count, statistics

def export_style(style_config, file_name):
    """Экспорт стиля в JSON файл"""
    try:
        # Добавляем метаданные в конфигурацию стиля
        export_data = {
            'version': '1.0',
            'export_date': str(datetime.now()),
            'style_config': style_config
        }
        
        # Конвертируем в JSON
        json_data = json.dumps(export_data, indent=2, ensure_ascii=False)
        
        # Создаем байты файла
        file_bytes = json_data.encode('utf-8')
        
        return file_bytes
    except Exception as e:
        st.error(f"Export error: {str(e)}")
        return None

def import_style(uploaded_file):
    """Импорт стиля из JSON файла"""
    try:
        # Читаем содержимое файла
        content = uploaded_file.read().decode('utf-8')
        
        # Парсим JSON
        import_data = json.loads(content)
        
        # Проверяем структуру
        if 'style_config' not in import_data:
            st.error(get_text('import_error'))
            return None
            
        return import_data['style_config']
    except Exception as e:
        st.error(f"{get_text('import_error')}: {str(e)}")
        return None

def apply_imported_style(imported_style):
    """Применение импортированной конфигурации стиля"""
    # Используем callback для безопасного обновления session_state
    st.session_state.num = imported_style.get('numbering_style', "No numbering")
    st.session_state.auth = imported_style.get('author_format', "AA Smith")
    st.session_state.sep = imported_style.get('author_separator', ", ")
    st.session_state.etal = imported_style.get('et_al_limit', 0) or 0
    st.session_state.use_and_checkbox = imported_style.get('use_and_bool', False)
    st.session_state.use_ampersand_checkbox = imported_style.get('use_ampersand_bool', False)
    st.session_state.doi = imported_style.get('doi_format', "10.10/xxx")
    st.session_state.doilink = imported_style.get('doi_hyperlink', True)
    st.session_state.page = imported_style.get('page_format', "122–128")
    st.session_state.punct = imported_style.get('final_punctuation', "")
    st.session_state.gost_style = imported_style.get('gost_style', False)
    st.session_state.acs_style = imported_style.get('acs_style', False)
    st.session_state.rsc_style = imported_style.get('rsc_style', False)
    st.session_state.cta_style = imported_style.get('cta_style', False)
    st.session_state.journal_style = imported_style.get('journal_style', '{Full Journal Name}')
    
    # Применяем элементы
    elements = imported_style.get('elements', [])
    for i in range(8):
        if i < len(elements):
            element, config = elements[i]
            st.session_state[f"el{i}"] = element
            st.session_state[f"it{i}"] = config.get('italic', False)
            st.session_state[f"bd{i}"] = config.get('bold', False)
            st.session_state[f"pr{i}"] = config.get('parentheses', False)
            st.session_state[f"sp{i}"] = config.get('separator', ". ")
        else:
            st.session_state[f"el{i}"] = ""
            st.session_state[f"it{i}"] = False
            st.session_state[f"bd{i}"] = False
            st.session_state[f"pr{i}"] = False
            st.session_state[f"sp{i}"] = ". "
    
    st.session_state.style_applied = True

def main():
    st.set_page_config(layout="wide")
    st.markdown("""
        <style>
        .block-container { padding: 0.2rem; }
        .stSelectbox, .stTextInput, .stNumberInput, .stCheckbox, .stRadio, .stFileUploader, .stTextArea {
            margin-bottom: 0.02rem;
        }
        .stTextArea { height: 40px !important; font-size: 0.7rem; }
        .stButton > button { width: 100%; padding: 0.05rem; font-size: 0.7rem; margin: 0.02rem; }
        h1 { font-size: 1.0rem; margin-bottom: 0.05rem; }
        h2 { font-size: 0.9rem; margin-bottom: 0.05rem; }
        h3 { font-size: 0.8rem; margin-bottom: 0.02rem; }
        label { font-size: 0.65rem !important; }
        .stMarkdown { font-size: 0.65rem; }
        .stCheckbox > label { font-size: 0.6rem; }
        .stRadio > label { font-size: 0.65rem; }
        .stDownloadButton > button { font-size: 0.7rem; padding: 0.05rem; margin: 0.02rem; }
        .element-row { margin: 0.01rem; padding: 0.01rem; }
        .processing-header { font-size: 0.8rem; font-weight: bold; margin-bottom: 0.1rem; }
        .processing-status { font-size: 0.7rem; margin-bottom: 0.05rem; }
        .compact-row { margin-bottom: 0.1rem; }
        .guide-text { font-size: 0.55rem !important; line-height: 1.1; margin-bottom: 0.1rem; }
        .guide-title { font-size: 0.6rem !important; font-weight: bold; margin-bottom: 0.1rem; }
        .guide-step { font-size: 0.55rem !important; line-height: 1.1; margin-bottom: 0.05rem; }
        .guide-note { font-size: 0.5rem !important; font-style: italic; line-height: 1.1; margin-bottom: 0.05rem; margin-left: 0.5rem; }
        </style>
    """, unsafe_allow_html=True)

    # Применяем импортированный стиль если нужно
    if st.session_state.apply_imported_style and st.session_state.imported_style:
        apply_imported_style(st.session_state.imported_style)
        st.session_state.apply_imported_style = False
        st.rerun()

    # Переключение языка
    language_options = [('English', 'en'), ('Русский', 'ru')]
    selected_language = st.selectbox(
        get_text('language'), 
        language_options, 
        format_func=lambda x: x[0], 
        index=0 if st.session_state.current_language == 'en' else 1,
        key="language_selector"
    )
    st.session_state.current_language = selected_language[1]

    st.title(get_text('header'))

    # Трёхколоночный макет
    col1, col2, col3 = st.columns([1, 1, 1])

    with col1:
        st.subheader(get_text('general_settings'))
        
        # Стили пресеты с тултипом
        col_preset, col_info = st.columns([3, 1])
        with col_preset:
            st.markdown(f"**{get_text('style_presets')}**")
        with col_info:
            st.markdown(f"<span title='{get_text('style_preset_tooltip')}'>ℹ️</span>", unsafe_allow_html=True)
        
        # Кнопки стилей в колонках
        col_gost, col_acs, col_rsc, col_cta = st.columns(4)
        
        with col_gost:
            if st.button(get_text('gost_button'), use_container_width=True, key="gost_button"):
                # Устанавливаем конфигурацию стиля ГОСТ
                st.session_state.num = "No numbering"  # Без автоматической нумерации
                st.session_state.auth = "Smith, A.A."
                st.session_state.sep = ", "
                st.session_state.etal = 0
                st.session_state.use_and_checkbox = False
                st.session_state.use_ampersand_checkbox = False
                st.session_state.doi = "https://dx.doi.org/10.10/xxx"
                st.session_state.doilink = True
                st.session_state.page = "122–128"
                st.session_state.punct = ""
                st.session_state.journal_style = "{Full Journal Name}"  # Полное название для ГОСТ
                
                # Очищаем все конфигурации элементов
                for i in range(8):
                    st.session_state[f"el{i}"] = ""
                    st.session_state[f"it{i}"] = False
                    st.session_state[f"bd{i}"] = False
                    st.session_state[f"pr{i}"] = False
                    st.session_state[f"sp{i}"] = ". "
                
                # Устанавливаем флаг стиля ГОСТ
                st.session_state.gost_style = True
                st.session_state.acs_style = False
                st.session_state.rsc_style = False
                st.session_state.cta_style = False
                st.session_state.style_applied = True
                st.rerun()
        
        with col_acs:
            if st.button(get_text('acs_button'), use_container_width=True, key="acs_button"):
                # Устанавливаем конфигурацию стиля ACS
                st.session_state.num = "No numbering"  # Без автоматической нумерации
                st.session_state.auth = "Smith, A.A."
                st.session_state.sep = "; "
                st.session_state.etal = 0
                st.session_state.use_and_checkbox = False
                st.session_state.use_ampersand_checkbox = False
                st.session_state.doi = "10.10/xxx"
                st.session_state.doilink = True
                st.session_state.page = "122–128"
                st.session_state.punct = "."
                st.session_state.journal_style = "{J. Abbr.}"  # Сокращения с точками для ACS
                
                # Очищаем все конфигурации элементов
                for i in range(8):
                    st.session_state[f"el{i}"] = ""
                    st.session_state[f"it{i}"] = False
                    st.session_state[f"bd{i}"] = False
                    st.session_state[f"pr{i}"] = False
                    st.session_state[f"sp{i}"] = ". "
                
                # Устанавливаем флаг стиля ACS
                st.session_state.gost_style = False
                st.session_state.acs_style = True
                st.session_state.rsc_style = False
                st.session_state.cta_style = False
                st.session_state.style_applied = True
                st.rerun()
        
        with col_rsc:
            if st.button(get_text('rsc_button'), use_container_width=True, key="rsc_button"):
                # Устанавливаем конфигурацию стиля RSC
                st.session_state.num = "No numbering"  # Без автоматической нумерации
                st.session_state.auth = "A.A. Smith"
                st.session_state.sep = ", "
                st.session_state.etal = 0
                st.session_state.use_and_checkbox = True
                st.session_state.use_ampersand_checkbox = False
                st.session_state.doi = "10.10/xxx"
                st.session_state.doilink = True
                st.session_state.page = "122"  # Только первая страница
                st.session_state.punct = "."
                st.session_state.journal_style = "{J. Abbr.}"  # Сокращения с точками для RSC
                
                # Очищаем все конфигурации элементов
                for i in range(8):
                    st.session_state[f"el{i}"] = ""
                    st.session_state[f"it{i}"] = False
                    st.session_state[f"bd{i}"] = False
                    st.session_state[f"pr{i}"] = False
                    st.session_state[f"sp{i}"] = ". "
                
                # Устанавливаем флаг стиля RSC
                st.session_state.gost_style = False
                st.session_state.acs_style = False
                st.session_state.rsc_style = True
                st.session_state.cta_style = False
                st.session_state.style_applied = True
                st.rerun()
        
        with col_cta:
            if st.button(get_text('cta_button'), use_container_width=True, key="cta_button"):
                # Устанавливаем конфигурацию стиля CTA
                st.session_state.num = "No numbering"  # Без автоматической нумерации
                st.session_state.auth = "Smith AA"
                st.session_state.sep = ", "
                st.session_state.etal = 0
                st.session_state.use_and_checkbox = False
                st.session_state.use_ampersand_checkbox = False
                st.session_state.doi = "doi:10.10/xxx"
                st.session_state.doilink = True
                st.session_state.page = "122–8"  # Сокращенный формат страниц
                st.session_state.punct = ""
                st.session_state.journal_style = "{J Abbr}"  # Сокращения без точек для CTA
                
                # Очищаем все конфигурации элементов
                for i in range(8):
                    st.session_state[f"el{i}"] = ""
                    st.session_state[f"it{i}"] = False
                    st.session_state[f"bd{i}"] = False
                    st.session_state[f"pr{i}"] = False
                    st.session_state[f"sp{i}"] = ". "
                
                # Устанавливаем флаг стиля CTA
                st.session_state.gost_style = False
                st.session_state.acs_style = False
                st.session_state.rsc_style = False
                st.session_state.cta_style = True
                st.session_state.style_applied = True
                st.rerun()
        
        # Инициализация значений по умолчанию
        default_values = {
            'num': "No numbering",
            'auth': "AA Smith", 
            'sep': ", ",
            'etal': 0,
            'use_and_checkbox': False,
            'use_ampersand_checkbox': False,
            'doi': "10.10/xxx",
            'doilink': True,
            'page': "122–128",
            'punct': "",
            'gost_style': False,
            'acs_style': False,
            'rsc_style': False,
            'cta_style': False,
            'journal_style': '{Full Journal Name}'
        }
        
        for key, default in default_values.items():
            if key not in st.session_state:
                st.session_state[key] = default
        
        # Настройки нумерации
        numbering_style = st.selectbox(
            get_text('numbering_style'), 
            NUMBERING_STYLES, 
            key="num", 
            index=NUMBERING_STYLES.index(st.session_state.num)
        )
        
        # Настройки авторов в одной строке
        col_authors = st.columns([1, 1, 1])
        with col_authors[0]:
            author_format = st.selectbox(
                get_text('author_format'), 
                AUTHOR_FORMATS, 
                key="auth", 
                index=AUTHOR_FORMATS.index(st.session_state.auth)
            )
        with col_authors[1]:
            author_separator = st.selectbox(
                get_text('author_separator'), 
                [", ", "; "], 
                key="sep", 
                index=[", ", "; "].index(st.session_state.sep)
            )
        with col_authors[2]:
            et_al_limit = st.number_input(
                get_text('et_al_limit'), 
                min_value=0, 
                step=1, 
                key="etal", 
                value=st.session_state.etal
            )
        
        # Чекбоксы для разделителей авторов в одной строке
        col_and_amp = st.columns(2)
        with col_and_amp[0]:
            use_and_checkbox = st.checkbox(
                get_text('use_and'), 
                key="use_and_checkbox", 
                value=st.session_state.use_and_checkbox,
                disabled=st.session_state.use_ampersand_checkbox
            )
        with col_and_amp[1]:
            use_ampersand_checkbox = st.checkbox(
                get_text('use_ampersand'), 
                key="use_ampersand_checkbox", 
                value=st.session_state.use_ampersand_checkbox,
                disabled=st.session_state.use_and_checkbox
            )
        
        # Стиль журнала
        journal_style = st.selectbox(
            get_text('journal_style'),
            JOURNAL_STYLES,
            key="journal_style",
            index=JOURNAL_STYLES.index(st.session_state.journal_style),
            format_func=lambda x: {
                "{Full Journal Name}": get_text('full_journal_name'),
                "{J. Abbr.}": get_text('journal_abbr_with_dots'),
                "{J Abbr}": get_text('journal_abbr_no_dots')
            }[x]
        )
        
        # Настройки страниц
        # Безопасное получение индекса для page_format
        current_page = st.session_state.page
        page_index = 3  # Значение по умолчанию "122–128"
        if current_page in PAGE_FORMATS:
            page_index = PAGE_FORMATS.index(current_page)
        
        page_format = st.selectbox(
            get_text('page_format'), 
            PAGE_FORMATS, 
            key="page", 
            index=page_index
        )
        
        # Настройки DOI в одной строке
        col_doi = st.columns([2, 1])
        with col_doi[0]:
            doi_format = st.selectbox(
                get_text('doi_format'), 
                DOI_FORMATS, 
                key="doi", 
                index=DOI_FORMATS.index(st.session_state.doi)
            )
        with col_doi[1]:
            doi_hyperlink = st.checkbox(
                get_text('doi_hyperlink'), 
                key="doilink", 
                value=st.session_state.doilink
            )
        
        # Конечная пунктуация
        final_punctuation = st.selectbox(
            get_text('final_punctuation'), 
            ["", "."], 
            key="punct", 
            index=["", "."].index(st.session_state.punct)
        )

    with col2:
        st.subheader(get_text('element_config'))
        element_configs = []
        used_elements = set()
        
        st.markdown(
            f"<small>{get_text('element')} | {get_text('italic')} | {get_text('bold')} | {get_text('parentheses')} | {get_text('separator')}</small>", 
            unsafe_allow_html=True
        )
        
        # Инициализация элементов
        for i in range(8):
            for prop in ['el', 'it', 'bd', 'pr', 'sp']:
                key = f"{prop}{i}"
                if key not in st.session_state:
                    if prop == 'sp':
                        st.session_state[key] = ". "
                    elif prop == 'el':
                        st.session_state[key] = ""
                    else:
                        st.session_state[key] = False
        
        # Конфигурация элементов
        for i in range(8):
            cols = st.columns([2, 1, 1, 1, 2])
            
            with cols[0]:
                element = st.selectbox(
                    "", 
                    AVAILABLE_ELEMENTS, 
                    key=f"el{i}", 
                    label_visibility="collapsed",
                    index=AVAILABLE_ELEMENTS.index(st.session_state[f"el{i}"]) if st.session_state[f"el{i}"] in AVAILABLE_ELEMENTS else 0
                )
            
            with cols[1]:
                italic = st.checkbox(
                    "", 
                    key=f"it{i}", 
                    help=get_text('italic'), 
                    value=st.session_state[f"it{i}"]
                )
            
            with cols[2]:
                bold = st.checkbox(
                    "", 
                    key=f"bd{i}", 
                    help=get_text('bold'), 
                    value=st.session_state[f"bd{i}"]
                )
            
            with cols[3]:
                parentheses = st.checkbox(
                    "", 
                    key=f"pr{i}", 
                    help=get_text('parentheses'), 
                    value=st.session_state[f"pr{i}"]
                )
            
            with cols[4]:
                separator = st.text_input(
                    "", 
                    value=st.session_state[f"sp{i}"], 
                    key=f"sp{i}", 
                    label_visibility="collapsed"
                )
            
            if element and element not in used_elements:
                element_configs.append((
                    element, 
                    {
                        'italic': italic, 
                        'bold': bold, 
                        'parentheses': parentheses, 
                        'separator': separator
                    }
                ))
                used_elements.add(element)

        # Добавляем краткое руководство
        st.markdown("---")
        st.markdown(f"<div class='guide-title'>{get_text('short_guide_title')}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='guide-step'>{get_text('step_1')}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='guide-note'>{get_text('step_1_note')}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='guide-step'>{get_text('step_2')}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='guide-step'>{get_text('step_3')}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='guide-step'>{get_text('step_4')}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='guide-step'>{get_text('step_5')}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='guide-note'>{get_text('step_5_note')}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='guide-step'>{get_text('step_6')}</div>", unsafe_allow_html=True)

    with col3:
        # Предпросмотр стиля
        st.subheader(get_text('style_preview'))
        
        # Собираем конфигурацию стиля
        style_config = {
            'author_format': st.session_state.auth,
            'author_separator': st.session_state.sep,
            'et_al_limit': st.session_state.etal if st.session_state.etal > 0 else None,
            'use_and_bool': st.session_state.use_and_checkbox,
            'use_ampersand_bool': st.session_state.use_ampersand_checkbox,
            'doi_format': st.session_state.doi,
            'doi_hyperlink': st.session_state.doilink,
            'page_format': st.session_state.page,
            'final_punctuation': st.session_state.punct,
            'numbering_style': st.session_state.num,
            'journal_style': st.session_state.journal_style,
            'elements': element_configs,
            'gost_style': st.session_state.get('gost_style', False),
            'acs_style': st.session_state.get('acs_style', False),
            'rsc_style': st.session_state.get('rsc_style', False),
            'cta_style': st.session_state.get('cta_style', False)
        }
        
        # Показываем пример форматирования
        if st.session_state.get('gost_style', False):
            # Пример для стиля ГОСТ
            preview_metadata = {
                'authors': [
                    {
                        'given': 'John A.', 
                        'family': 'Smith'
                    }, 
                    {
                        'given': 'Alice B.', 
                        'family': 'Doe'
                    }
                ],
                'title': 'Article Title',
                'journal': 'Journal of the American Chemical Society',
                'year': 2020,
                'volume': '15',
                'issue': '3',
                'pages': '122-128',
                'article_number': '',
                'doi': '10.1000/xyz123'
            }
            preview_ref, _ = format_reference(preview_metadata, style_config, for_preview=True)
            
            numbering = style_config['numbering_style']
            if numbering == "No numbering":
                preview_ref_with_numbering = preview_ref
            else:
                if numbering == "1":
                    preview_ref_with_numbering = f"1 {preview_ref}"
                elif numbering == "1.":
                    preview_ref_with_numbering = f"1. {preview_ref}"
                elif numbering == "1)":
                    preview_ref_with_numbering = f"1) {preview_ref}"
                elif numbering == "(1)":
                    preview_ref_with_numbering = f"(1) {preview_ref}"
                elif numbering == "[1]":
                    preview_ref_with_numbering = f"[1] {preview_ref}"
                else:
                    preview_ref_with_numbering = f"1. {preview_ref}"
            
            st.markdown(f"<small>{get_text('example')} {preview_ref_with_numbering}</small>", unsafe_allow_html=True)
        
        elif st.session_state.get('acs_style', False):
            # Пример для стиля ACS
            preview_metadata = {
                'authors': [
                    {
                        'given': 'John A.', 
                        'family': 'Smith'
                    }, 
                    {
                        'given': 'Alice B.', 
                        'family': 'Doe'
                    }
                ],
                'title': 'Article Title',
                'journal': 'Journal of the American Chemical Society',
                'year': 2020,
                'volume': '15',
                'issue': '3',
                'pages': '122-128',
                'article_number': '',
                'doi': '10.1000/xyz123'
            }
            preview_ref, _ = format_reference(preview_metadata, style_config, for_preview=True)
            
            numbering = style_config['numbering_style']
            if numbering == "No numbering":
                preview_ref_with_numbering = preview_ref
            else:
                if numbering == "1":
                    preview_ref_with_numbering = f"1 {preview_ref}"
                elif numbering == "1.":
                    preview_ref_with_numbering = f"1. {preview_ref}"
                elif numbering == "1)":
                    preview_ref_with_numbering = f"1) {preview_ref}"
                elif numbering == "(1)":
                    preview_ref_with_numbering = f"(1) {preview_ref}"
                elif numbering == "[1]":
                    preview_ref_with_numbering = f"[1] {preview_ref}"
                else:
                    preview_ref_with_numbering = f"1. {preview_ref}"
            
            # Исправление для требования 3 - показываем форматирование в предпросмотре
            preview_html = preview_ref_with_numbering
            # Добавляем HTML теги для форматирования
            preview_html = preview_html.replace("J. Am. Chem. Soc.", "<i>J. Am. Chem. Soc.</i>")
            preview_html = preview_html.replace("2020", "<b>2020</b>")
            preview_html = preview_html.replace("15", "<i>15</i>")
            
            st.markdown(f"<small>{get_text('example')} {preview_html}</small>", unsafe_allow_html=True)
        
        elif st.session_state.get('rsc_style', False):
            # Пример для стиля RSC
            preview_metadata = {
                'authors': [
                    {
                        'given': 'John A.', 
                        'family': 'Smith'
                    }, 
                    {
                        'given': 'Alice B.', 
                        'family': 'Doe'
                    }
                ],
                'title': 'Article Title',
                'journal': 'Chemical Communications',
                'year': 2020,
                'volume': '15',
                'issue': '3',
                'pages': '122-128',
                'article_number': '',
                'doi': '10.1000/xyz123'
            }
            preview_ref, _ = format_reference(preview_metadata, style_config, for_preview=True)
            
            numbering = style_config['numbering_style']
            if numbering == "No numbering":
                preview_ref_with_numbering = preview_ref
            else:
                if numbering == "1":
                    preview_ref_with_numbering = f"1 {preview_ref}"
                elif numbering == "1.":
                    preview_ref_with_numbering = f"1. {preview_ref}"
                elif numbering == "1)":
                    preview_ref_with_numbering = f"1) {preview_ref}"
                elif numbering == "(1)":
                    preview_ref_with_numbering = f"(1) {preview_ref}"
                elif numbering == "[1]":
                    preview_ref_with_numbering = f"[1] {preview_ref}"
                else:
                    preview_ref_with_numbering = f"1. {preview_ref}"
            
            # Исправление для требования 3 - показываем форматирование в предпросмотре
            preview_html = preview_ref_with_numbering
            # Добавляем HTML теги для форматирования
            preview_html = preview_html.replace("Chem. Commun.", "<i>Chem. Commun.</i>")
            preview_html = preview_html.replace("15", "<b>15</b>")
            
            st.markdown(f"<small>{get_text('example')} {preview_html}</small>", unsafe_allow_html=True)
        
        elif st.session_state.get('cta_style', False):
            # Пример для стиля CTA
            preview_metadata = {
                'authors': [
                    {
                        'given': 'Fei', 
                        'family': 'He'
                    }, 
                    {
                        'given': 'Feng', 
                        'family': 'Ma'
                    },
                    {
                        'given': 'Juan', 
                        'family': 'Li'
                    },
                    {
                        'given': 'Tao', 
                        'family': 'Li'
                    },
                    {
                        'given': 'Guangshe', 
                        'family': 'Li'
                    }
                ],
                'title': 'Effect of calcination temperature on the structural properties and photocatalytic activities of solvothermal synthesized TiO2 hollow nanoparticles',
                'journal': 'Ceramics International',
                'year': 2014,
                'volume': '40',
                'issue': '5',
                'pages': '6441-6446',
                'article_number': '',
                'doi': '10.1016/j.ceramint.2013.11.094'
            }
            preview_ref, _ = format_reference(preview_metadata, style_config, for_preview=True)
            
            numbering = style_config['numbering_style']
            if numbering == "No numbering":
                preview_ref_with_numbering = preview_ref
            else:
                if numbering == "1":
                    preview_ref_with_numbering = f"1 {preview_ref}"
                elif numbering == "1.":
                    preview_ref_with_numbering = f"1. {preview_ref}"
                elif numbering == "1)":
                    preview_ref_with_numbering = f"1) {preview_ref}"
                elif numbering == "(1)":
                    preview_ref_with_numbering = f"(1) {preview_ref}"
                elif numbering == "[1]":
                    preview_ref_with_numbering = f"[1] {preview_ref}"
                else:
                    preview_ref_with_numbering = f"1. {preview_ref}"
            
            st.markdown(f"<small>{get_text('example')} {preview_ref_with_numbering}</small>", unsafe_allow_html=True)
        
        elif not style_config['elements']:
            st.markdown(
                f"<b style='color:red; font-size: 0.7rem;'>{get_text('error_select_element')}</b>", 
                unsafe_allow_html=True
            )
        else:
            # Пример для обычного стиля
            preview_metadata = {
                'authors': [
                    {
                        'given': 'John A.', 
                        'family': 'Smith'
                    }, 
                    {
                        'given': 'Alice B.', 
                        'family': 'Doe'
                    }
                ],
                'title': 'Article Title',
                'journal': 'Journal of the American Chemical Society',
                'year': 2020,
                'volume': '15',
                'issue': '3',
                'pages': '122-128',
                'article_number': 'e12345',
                'doi': '10.1000/xyz123'
            }
            preview_ref, _ = format_reference(preview_metadata, style_config, for_preview=True)
            
            numbering = style_config['numbering_style']
            if numbering == "No numbering":
                preview_ref_with_numbering = preview_ref
            else:
                if numbering == "1":
                    preview_ref_with_numbering = f"1 {preview_ref}"
                elif numbering == "1.":
                    preview_ref_with_numbering = f"1. {preview_ref}"
                elif numbering == "1)":
                    preview_ref_with_numbering = f"1) {preview_ref}"
                elif numbering == "(1)":
                    preview_ref_with_numbering = f"(1) {preview_ref}"
                elif numbering == "[1]":
                    preview_ref_with_numbering = f"[1] {preview_ref}"
                else:
                    preview_ref_with_numbering = f"1. {preview_ref}"
            
            st.markdown(f"<small>{get_text('example')} {preview_ref_with_numbering}</small>", unsafe_allow_html=True)

        # Ввод данных
        st.subheader(get_text('data_input'))
        input_method = st.radio(
            get_text('input_method'), 
            ['DOCX', 'Text' if st.session_state.current_language == 'en' else 'Текст'], 
            horizontal=True, 
            key="input_method"
        )
        
        if input_method == 'DOCX':
            uploaded_file = st.file_uploader(
                get_text('select_docx'), 
                type=['docx'], 
                label_visibility="collapsed", 
                key="docx_uploader"
            )
        else:
            references_input = st.text_area(
                get_text('references'), 
                placeholder=get_text('enter_references'), 
                height=40, 
                label_visibility="collapsed", 
                key="references_input"
            )

        # Вывод данных
        st.subheader(get_text('data_output'))
        output_method = st.radio(
            get_text('output_method'), 
            ['DOCX', 'Text' if st.session_state.current_language == 'en' else 'Текст'], 
            horizontal=True, 
            key="output_method"
        )
        
        # Текстовое поле для результатов (показывается только если выбран текстовый вывод)
        if output_method == 'Text' if st.session_state.current_language == 'en' else 'Текст':
            output_text_value = st.session_state.output_text_value if st.session_state.show_results else ""
            st.text_area(
                get_text('results'), 
                value=output_text_value, 
                height=40, 
                disabled=True, 
                label_visibility="collapsed", 
                key="output_text"
            )

        # Кнопка обработки
        if st.button(get_text('process'), use_container_width=True, key="process_button"):
            if not style_config['elements'] and not style_config.get('gost_style', False) and not style_config.get('acs_style', False) and not style_config.get('rsc_style', False) and not style_config.get('cta_style', False):
                st.error(get_text('error_select_element'))
                return
                
            # Создаем контейнеры для прогресса
            progress_container = st.empty()
            status_container = st.empty()
            
            try:
                if input_method == 'DOCX':
                    if not uploaded_file:
                        st.error(get_text('upload_file'))
                        return
                    
                    with st.spinner(get_text('processing')):
                        formatted_refs, txt_bytes, output_doc_buffer, doi_found_count, doi_not_found_count, statistics = process_docx(
                            uploaded_file, style_config, progress_container, status_container
                        )
                else:
                    if not references_input.strip():
                        st.error(get_text('enter_references_error'))
                        return
                    
                    references = [ref.strip() for ref in references_input.split('\n') if ref.strip()]
                    st.write(f"**{get_text('found_references_text').format(len(references))}**")
                    
                    with st.spinner(get_text('processing')):
                        formatted_refs, txt_bytes, doi_found_count, doi_not_found_count, duplicates_info = process_references_with_progress(
                            references, style_config, progress_container, status_container
                        )
                        
                        # Генерируем статистику
                        statistics = generate_statistics(formatted_refs)
                        
                        # Создаем DOCX документ для текстового ввода
                        output_doc_buffer = DocumentGenerator.generate_document(
                            formatted_refs, statistics, style_config, duplicates_info
                        )

                # Очищаем контейнеры прогресса
                progress_container.empty()
                status_container.empty()
                
                # Показываем статистику
                st.write(f"**{get_text('statistics').format(doi_found_count, doi_not_found_count)}**")
                
                # Подготовка данных для вывода
                if output_method == 'Text' if st.session_state.current_language == 'en' else 'Текст':
                    output_text_value = ""
                    for i, (elements, is_error, metadata) in enumerate(formatted_refs):
                        numbering = style_config['numbering_style']
                        
                        if numbering == "No numbering":
                            prefix = ""
                        elif numbering == "1":
                            prefix = f"{i + 1} "
                        elif numbering == "1.":
                            prefix = f"{i + 1}. "
                        elif numbering == "1)":
                            prefix = f"{i + 1}) "
                        elif numbering == "(1)":
                            prefix = f"({i + 1}) "
                        elif numbering == "[1]":
                            prefix = f"[{i + 1}] "
                        else:
                            prefix = f"{i + 1}. "
                        
                        if is_error:
                            output_text_value += f"{prefix}{elements}\n"
                        else:
                            if isinstance(elements, str):
                                output_text_value += f"{prefix}{elements}\n"
                            else:
                                ref_str = ""
                                for j, element_data in enumerate(elements):
                                    if len(element_data) == 6:
                                        value, _, _, separator, _, _ = element_data
                                        ref_str += value
                                        if separator and j < len(elements) - 1:
                                            ref_str += separator
                                    else:
                                        ref_str += str(element_data)
                                
                                if style_config['final_punctuation'] and not is_error:
                                    ref_str = ref_str.rstrip(',.') + "."
                                
                                output_text_value += f"{prefix}{ref_str}\n"
                    
                    # Сохраняем данные для отображения
                    st.session_state.output_text_value = output_text_value
                    st.session_state.show_results = True
                else:
                    st.session_state.output_text_value = ""
                    st.session_state.show_results = False

                # Сохраняем данные для скачивания
                st.session_state.download_data = {
                    'txt_bytes': txt_bytes,
                    'output_doc_buffer': output_doc_buffer if output_method == 'DOCX' else None
                }
                
            except Exception as e:
                st.error(f"Processing error: {str(e)}")
                return
            
            st.rerun()

        # Кнопки скачивания в одной строке
        if st.session_state.download_data:
            col_download = st.columns(2)
            with col_download[0]:
                st.download_button(
                    label=get_text('doi_txt'),
                    data=st.session_state.download_data['txt_bytes'],
                    file_name='doi_list.txt',
                    mime='text/plain',
                    key="doi_download",
                    use_container_width=True
                )
            
            with col_download[1]:
                if output_method == 'DOCX' and st.session_state.download_data.get('output_doc_buffer'):
                    st.download_button(
                        label=get_text('references_docx'),
                        data=st.session_state.download_data['output_doc_buffer'],
                        file_name='Reformatted references.docx',  # Изменено согласно требованию 3
                        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                        key="docx_download",
                        use_container_width=True
                    )

        # Управление стилями
        st.subheader("💾 Style Management")
        
        # Экспорт стиля в одной строке
        col_export = st.columns([2, 1])
        with col_export[0]:
            export_file_name = st.text_input(
                get_text('export_file_name'), 
                value="my_citation_style", 
                placeholder="Enter file name", 
                key="export_name"
            )
        
        with col_export[1]:
            # Создаем конфигурацию текущего стиля для экспорта
            current_style_config = {
                'author_format': st.session_state.auth,
                'author_separator': st.session_state.sep,
                'et_al_limit': st.session_state.etal if st.session_state.etal > 0 else None,
                'use_and_bool': st.session_state.use_and_checkbox,
                'use_ampersand_bool': st.session_state.use_ampersand_checkbox,
                'doi_format': st.session_state.doi,
                'doi_hyperlink': st.session_state.doilink,
                'page_format': st.session_state.page,
                'final_punctuation': st.session_state.punct,
                'numbering_style': st.session_state.num,
                'journal_style': st.session_state.journal_style,
                'elements': element_configs,
                'gost_style': st.session_state.get('gost_style', False),
                'acs_style': st.session_state.get('acs_style', False),
                'rsc_style': st.session_state.get('rsc_style', False),
                'cta_style': st.session_state.get('cta_style', False)
            }
            
            # Кнопка экспорта
            export_data = export_style(current_style_config, export_file_name)
            if export_data:
                st.download_button(
                    label=get_text('export_style'),
                    data=export_data,
                    file_name=f"{export_file_name}.json",
                    mime="application/json",
                    use_container_width=True,
                    key="export_button"
                )
        
        # Импорт стиля
        imported_file = st.file_uploader(
            get_text('import_file'), 
            type=['json'], 
            label_visibility="collapsed", 
            key="style_importer"
        )
        
        if imported_file is not None and not st.session_state.style_applied:
            imported_style = import_style(imported_file)
            if imported_style:
                # Сохраняем импортированный стиль и устанавливаем флаг для применения
                st.session_state.imported_style = imported_style
                st.session_state.apply_imported_style = True
                st.success(get_text('import_success'))
                st.rerun()

if __name__ == "__main__":
    main()


