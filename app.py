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
        'import_error': 'Error importing style file!'
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
        'import_error': 'Ошибка импорта файла стиля!'
    }
}

# Хранение текущего языка
if 'current_language' not in st.session_state:
    st.session_state.current_language = 'ru'

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

def get_text(key):
    return TRANSLATIONS[st.session_state.current_language].get(key, key)

def clean_text(text):
    return re.sub(r'<[^>]+>|&[^;]+;', '', text).strip()

def normalize_name(name):
    if not name:
        return ''
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
    
    return None

def extract_metadata(doi):
    """Извлекает метаданные по DOI через Crossref API"""
    try:
        print(f"Extracting metadata for DOI: {doi}")
        result = works.doi(doi)
        if not result:
            print(f"No result for DOI: {doi}")
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
        
        journal = ''
        if 'container-title' in result and result['container-title']:
            journal = result['container-title'][0]
        
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
            'doi': doi
        }
        
        print(f"Successfully extracted metadata: {metadata['title'][:50]}...")
        return metadata
        
    except Exception as e:
        print(f"Error extracting metadata for DOI {doi}: {e}")
        return None

def format_authors(authors, author_format, separator, et_al_limit, use_and_bool, use_ampersand_bool):
    if not authors:
        return ""
    
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
        
        # Добавляем разделитель между авторами
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

def format_pages(pages, article_number, page_format):
    if pages:
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

def add_hyperlink(paragraph, text, url):
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

def format_reference(metadata, style_config, for_preview=False):
    if not metadata:
        error_message = "Ошибка: Не удалось отформатировать ссылку." if st.session_state.current_language == 'ru' else "Error: Could not format the reference."
        return (error_message, True)
    
    # Проверяем, включен ли стиль ГОСТ
    if style_config.get('gost_style', False):
        return format_gost_reference(metadata, style_config, for_preview)
    
    elements = []
    
    for i, (element, config) in enumerate(style_config['elements']):
        value = ""
        doi_value = None
        
        if element == "Authors":
            value = format_authors(
                metadata['authors'],
                style_config['author_format'],
                style_config['author_separator'],
                style_config['et_al_limit'],
                style_config['use_and_bool'],
                style_config['use_ampersand_bool']
            )
        elif element == "Title":
            value = metadata['title']
        elif element == "Journal":
            value = metadata['journal']
        elif element == "Year":
            value = str(metadata['year']) if metadata['year'] else ""
        elif element == "Volume":
            value = metadata['volume']
        elif element == "Issue":
            value = metadata['issue']
        elif element == "Pages":
            value = format_pages(metadata['pages'], metadata['article_number'], style_config['page_format'])
        elif element == "DOI":
            doi = metadata['doi']
            doi_value = doi
            if style_config['doi_format'] == "10.10/xxx":
                value = doi
            elif style_config['doi_format'] == "doi:10.10/xxx":
                value = f"doi:{doi}"
            elif style_config['doi_format'] == "DOI:10.10/xxx":
                value = f"DOI:{doi}"
            elif style_config['doi_format'] == "https://dx.doi.org/10.10/xxx":
                value = f"https://dx.doi.org/{doi}"
        
        if value:
            # Добавляем скобки если нужно
            if config['parentheses'] and value:
                value = f"({value})"
            
            # Добавляем разделитель
            separator = config['separator'] if i < len(style_config['elements']) - 1 else ''
            
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
                               (element == "DOI" and style_config['doi_hyperlink']), doi_value))
    
    if for_preview:
        # Собираем строку для предпросмотра
        ref_str = ""
        for i, (value, _, _, separator, _, _) in enumerate(elements):
            ref_str += value
            if separator and i < len(elements) - 1:
                ref_str += separator
            elif i == len(elements) - 1 and style_config['final_punctuation']:
                ref_str = ref_str.rstrip(',.') + "."
        
        return ref_str, False
    else:
        return elements, False

def format_gost_reference(metadata, style_config, for_preview=False):
    """Форматирование ссылки по стандарту ГОСТ"""
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
    
    # Строим ссылку ГОСТ с номером выпуска, если доступно
    if metadata['issue']:
        gost_ref = f"{first_author} {metadata['title']} / {all_authors} // {metadata['journal']}. – {metadata['year']}. – {volume_label} {metadata['volume']}. – {issue_label} {metadata['issue']}."
    else:
        gost_ref = f"{first_author} {metadata['title']} / {all_authors} // {metadata['journal']}. – {metadata['year']}. – {volume_label} {metadata['volume']}."
    
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

def apply_yellow_background(run):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), 'FFFF00')
    run._element.get_or_add_rPr().append(shd)

def process_references(references, style_config):
    """Обрабатывает список ссылок и возвращает отформатированные результаты"""
    doi_list = []
    formatted_refs = []
    doi_found_count = 0
    doi_not_found_count = 0
    
    progress_bar = tqdm(total=len(references), desc=get_text('processing'))
    
    for ref in references:
        if is_section_header(ref):
            doi_list.append(f"{ref} [SECTION HEADER - SKIPPED]")
            formatted_refs.append((ref, False, None))
            progress_bar.update(1)
            continue
            
        doi = find_doi(ref)
        print(f"Processing reference: '{ref}' -> DOI: {doi}")
        
        if doi:
            doi_list.append(doi)
            metadata = extract_metadata(doi)
            
            if metadata:
                print(f"Successfully got metadata for DOI: {doi}")
                formatted_ref, is_error = format_reference(metadata, style_config)
                formatted_refs.append((formatted_ref, is_error, metadata))
                
                if not is_error:
                    doi_found_count += 1
                else:
                    error_message = f"{ref} [ОШИБКА: Не удалось отформатировать ссылку. Проверьте DOI вручную.]" if st.session_state.current_language == 'ru' else f"{ref} [ERROR: Could not format reference. Please check DOI manually.]"
                    doi_list[-1] = f"{doi}\nПроверьте источник и добавьте DOI вручную." if st.session_state.current_language == 'ru' else f"{doi}\nPlease check this source and insert the DOI manually."
                    formatted_refs.append((error_message, True, None))
                    doi_not_found_count += 1
            else:
                print(f"Failed to get metadata for DOI: {doi}")
                error_message = f"{ref} [ОШИБКА: Не удалось получить метаданные по DOI. Проверьте DOI вручную.]" if st.session_state.current_language == 'ru' else f"{ref} [ERROR: Could not get metadata for DOI. Please check DOI manually.]"
                doi_list[-1] = f"{doi}\nПроверьте источник и добавьте DOI вручную." if st.session_state.current_language == 'ru' else f"{doi}\nPlease check this source and insert the DOI manually."
                formatted_refs.append((error_message, True, None))
                doi_not_found_count += 1
        else:
            print(f"No DOI found in reference: '{ref}'")
            error_message = f"{ref} [ОШИБКА: DOI не найден. Проверьте ссылку вручную.]" if st.session_state.current_language == 'ru' else f"{ref} [ERROR: DOI not found. Please check reference manually.]"
            doi_list.append(f"{ref}\nПроверьте источник и добавьте DOI вручную." if st.session_state.current_language == 'ru' else f"{ref}\nPlease check this source and insert the DOI manually.")
            formatted_refs.append((error_message, True, None))
            doi_not_found_count += 1
            
        progress_bar.update(1)
        
    progress_bar.close()
    
    # Выводим статистику
    st.write(f"**{get_text('statistics').format(doi_found_count, doi_not_found_count)}**")
    
    # Создаем TXT файл со списком DOI
    output_txt_buffer = io.StringIO()
    for doi in doi_list:
        output_txt_buffer.write(f"{doi}\n")
    output_txt_buffer.seek(0)
    txt_bytes = io.BytesIO(output_txt_buffer.getvalue().encode('utf-8'))
    
    return formatted_refs, txt_bytes, doi_found_count, doi_not_found_count

def process_docx(input_file, style_config):
    """Обрабатывает DOCX файл с ссылками"""
    doc = Document(input_file)
    references = []
    
    for para in doc.paragraphs:
        if para.text.strip():
            references.append(para.text.strip())
    
    st.write(f"**{get_text('found_references').format(len(references))}**")
    
    # Обрабатываем все ссылки напрямую
    formatted_refs, txt_bytes, doi_found_count, doi_not_found_count = process_references(references, style_config)
    
    # Создаем новый DOCX документ с отформатированными ссылками
    output_doc = Document()
    
    if st.session_state.current_language == 'en':
        output_doc.add_heading('References in Custom Style', level=1)
    else:
        output_doc.add_heading('Ссылки в пользовательском стиле', level=1)
    
    for i, (elements, is_error, metadata) in enumerate(formatted_refs, 1):
        numbering = style_config['numbering_style']
        
        # Формируем префикс нумерации
        if numbering == "No numbering":
            prefix = ""
        elif numbering == "1":
            prefix = f"{i} "
        elif numbering == "1.":
            prefix = f"{i}. "
        elif numbering == "1)":
            prefix = f"{i}) "
        elif numbering == "(1)":
            prefix = f"({i}) "
        elif numbering == "[1]":
            prefix = f"[{i}] "
        else:
            prefix = f"{i}. "
        
        para = output_doc.add_paragraph(prefix)
        
        if is_error:
            # Показываем оригинальный текст с желтым фоном и сообщением об ошибке
            run = para.add_run(str(elements))
            apply_yellow_background(run)
        else:
            if metadata is None:
                run = para.add_run(str(elements))
                run.font.italic = True
            else:
                for j, (value, italic, bold, separator, is_doi_hyperlink, doi_value) in enumerate(elements):
                    if is_doi_hyperlink and doi_value:
                        # Добавляем DOI как гиперссылку
                        add_hyperlink(para, value, f"https://doi.org/{doi_value}")
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
    
    output_doc_buffer = io.BytesIO()
    output_doc.save(output_doc_buffer)
    output_doc_buffer.seek(0)
    
    return formatted_refs, txt_bytes, output_doc_buffer

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
        </style>
    """, unsafe_allow_html=True)

    # Применяем импортированный стиль если нужно
    if st.session_state.apply_imported_style and st.session_state.imported_style:
        apply_imported_style(st.session_state.imported_style)
        st.session_state.apply_imported_style = False
        st.rerun()

    # Переключение языка
    language_options = [('Русский', 'ru'), ('English', 'en')]
    selected_language = st.selectbox(
        get_text('language'), 
        language_options, 
        format_func=lambda x: x[0], 
        index=0 if st.session_state.current_language == 'ru' else 1,
        key="language_selector"
    )
    st.session_state.current_language = selected_language[1]

    st.title(get_text('header'))

    # Трёхколоночный макет
    col1, col2, col3 = st.columns([1, 1, 1])

    with col1:
        st.subheader(get_text('general_settings'))
        
        # Кнопка применения ГОСТ стиля
        if st.button(get_text('gost_style'), use_container_width=True, key="gost_button"):
            # Устанавливаем конфигурацию стиля ГОСТ
            st.session_state.num = "1."
            st.session_state.auth = "Smith, A.A."
            st.session_state.sep = ", "
            st.session_state.etal = 0
            st.session_state.use_and_checkbox = False
            st.session_state.use_ampersand_checkbox = False
            st.session_state.doi = "https://dx.doi.org/10.10/xxx"
            st.session_state.doilink = True
            st.session_state.page = "122–128"
            st.session_state.punct = ""
            
            # Очищаем все конфигурации элементов
            for i in range(8):
                st.session_state[f"el{i}"] = ""
                st.session_state[f"it{i}"] = False
                st.session_state[f"bd{i}"] = False
                st.session_state[f"pr{i}"] = False
                st.session_state[f"sp{i}"] = ". "
            
            # Устанавливаем флаг стиля ГОСТ
            st.session_state.gost_style = True
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
            'gost_style': False
        }
        
        for key, default in default_values.items():
            if key not in st.session_state:
                st.session_state[key] = default
        
        # Настройки нумерации
        numbering_style = st.selectbox(
            get_text('numbering_style'), 
            ["No numbering", "1", "1.", "1)", "(1)", "[1]"], 
            key="num", 
            index=["No numbering", "1", "1.", "1)", "(1)", "[1]"].index(st.session_state.num)
        )
        
        # Настройки авторов
        author_format = st.selectbox(
            get_text('author_format'), 
            ["AA Smith", "A.A. Smith", "Smith AA", "Smith A.A", "Smith, A.A."], 
            key="auth", 
            index=["AA Smith", "A.A. Smith", "Smith AA", "Smith A.A", "Smith, A.A."].index(st.session_state.auth)
        )
        
        author_separator = st.selectbox(
            get_text('author_separator'), 
            [", ", "; "], 
            key="sep", 
            index=[", ", "; "].index(st.session_state.sep)
        )
        
        et_al_limit = st.number_input(
            get_text('et_al_limit'), 
            min_value=0, 
            step=1, 
            key="etal", 
            value=st.session_state.etal
        )
        
        # Чекбоксы для разделителей авторов
        col_and, col_amp = st.columns(2)
        with col_and:
            use_and_checkbox = st.checkbox(
                get_text('use_and'), 
                key="use_and_checkbox", 
                value=st.session_state.use_and_checkbox,
                disabled=st.session_state.use_ampersand_checkbox
            )
        with col_amp:
            use_ampersand_checkbox = st.checkbox(
                get_text('use_ampersand'), 
                key="use_ampersand_checkbox", 
                value=st.session_state.use_ampersand_checkbox,
                disabled=st.session_state.use_and_checkbox
            )
        
        # Настройки DOI
        doi_format = st.selectbox(
            get_text('doi_format'), 
            ["10.10/xxx", "doi:10.10/xxx", "DOI:10.10/xxx", "https://dx.doi.org/10.10/xxx"], 
            key="doi", 
            index=["10.10/xxx", "doi:10.10/xxx", "DOI:10.10/xxx", "https://dx.doi.org/10.10/xxx"].index(st.session_state.doi)
        )
        
        doi_hyperlink = st.checkbox(
            get_text('doi_hyperlink'), 
            key="doilink", 
            value=st.session_state.doilink
        )
        
        # Настройки страниц
        page_format = st.selectbox(
            get_text('page_format'), 
            ["122 - 128", "122-128", "122 – 128", "122–128", "122–8"], 
            key="page", 
            index=["122 - 128", "122-128", "122 – 128", "122–128", "122–8"].index(st.session_state.page)
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
        available_elements = ["", "Authors", "Title", "Journal", "Year", "Volume", "Issue", "Pages", "DOI"]
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
                    available_elements, 
                    key=f"el{i}", 
                    label_visibility="collapsed",
                    index=available_elements.index(st.session_state[f"el{i}"]) if st.session_state[f"el{i}"] in available_elements else 0
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
            'elements': element_configs,
            'gost_style': st.session_state.get('gost_style', False)
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
                'journal': 'Journal Name',
                'year': 2020,
                'volume': '15',
                'issue': '3',
                'pages': '',
                'article_number': 'e12345',
                'doi': '10.1000/xyz123'
            }
            preview_ref, _ = format_gost_reference(preview_metadata, style_config, for_preview=True)
            
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
                'journal': 'Journal Name',
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
            if not style_config['elements'] and not style_config.get('gost_style', False):
                st.error(get_text('error_select_element'))
                return
                
            if input_method == 'DOCX':
                if not uploaded_file:
                    st.error(get_text('upload_file'))
                    return
                
                with st.spinner(get_text('processing')):
                    formatted_refs, txt_bytes, output_doc_buffer = process_docx(uploaded_file, style_config)
            else:
                if not references_input.strip():
                    st.error(get_text('enter_references_error'))
                    return
                
                references = [ref.strip() for ref in references_input.split('\n') if ref.strip()]
                st.write(f"**{get_text('found_references_text').format(len(references))}**")
                
                with st.spinner(get_text('processing')):
                    formatted_refs, txt_bytes, _, _ = process_references(references, style_config)
                    
                    # Создаем DOCX документ для текстового ввода
                    output_doc = Document()
                    
                    if st.session_state.current_language == 'en':
                        output_doc.add_heading('References in Custom Style', level=1)
                    else:
                        output_doc.add_heading('Ссылки в пользовательском стиле', level=1)
                    
                    for i, (elements, is_error, metadata) in enumerate(formatted_refs, 1):
                        numbering = style_config['numbering_style']
                        
                        if numbering == "No numbering":
                            prefix = ""
                        elif numbering == "1":
                            prefix = f"{i} "
                        elif numbering == "1.":
                            prefix = f"{i}. "
                        elif numbering == "1)":
                            prefix = f"{i}) "
                        elif numbering == "(1)":
                            prefix = f"({i}) "
                        elif numbering == "[1]":
                            prefix = f"[{i}] "
                        else:
                            prefix = f"{i}. "
                        
                        para = output_doc.add_paragraph(prefix)
                        
                        if is_error:
                            run = para.add_run(str(elements))
                            apply_yellow_background(run)
                        else:
                            for j, (value, italic, bold, separator, is_doi_hyperlink, doi_value) in enumerate(elements):
                                if is_doi_hyperlink and doi_value:
                                    add_hyperlink(para, value, f"https://doi.org/{doi_value}")
                                else:
                                    run = para.add_run(value)
                                    if italic:
                                        run.font.italic = True
                                    if bold:
                                        run.font.bold = True
                                
                                if separator and j < len(elements) - 1:
                                    para.add_run(separator)
                            
                            if style_config['final_punctuation'] and not is_error:
                                para.add_run(".")
                    
                    output_doc_buffer = io.BytesIO()
                    output_doc.save(output_doc_buffer)
                    output_doc_buffer.seek(0)

            # Подготовка данных для вывода
            if output_method == 'Text' if st.session_state.current_language == 'en' else 'Текст':
                output_text_value = ""
                for i, (elements, is_error, metadata) in enumerate(formatted_refs, 1):
                    numbering = style_config['numbering_style']
                    
                    if numbering == "No numbering":
                        prefix = ""
                    elif numbering == "1":
                        prefix = f"{i} "
                    elif numbering == "1.":
                        prefix = f"{i}. "
                    elif numbering == "1)":
                        prefix = f"{i}) "
                    elif numbering == "(1)":
                        prefix = f"({i}) "
                    elif numbering == "[1]":
                        prefix = f"[{i}] "
                    else:
                        prefix = f"{i}. "
                    
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
            
            st.rerun()

        # Кнопки скачивания
        if st.session_state.download_data:
            st.download_button(
                label=get_text('doi_txt'),
                data=st.session_state.download_data['txt_bytes'],
                file_name='doi_list.txt',
                mime='text/plain',
                key="doi_download"
            )
            
            if output_method == 'DOCX' and st.session_state.download_data.get('output_doc_buffer'):
                st.download_button(
                    label=get_text('references_docx'),
                    data=st.session_state.download_data['output_doc_buffer'],
                    file_name='references_custom.docx',
                    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    key="docx_download"
                )

        # Управление стилями
        st.subheader("💾 Style Management")
        
        # Экспорт текущего стиля
        export_file_name = st.text_input(
            get_text('export_file_name'), 
            value="my_citation_style", 
            placeholder="Enter file name", 
            key="export_name"
        )
        
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
            'elements': element_configs,
            'gost_style': st.session_state.get('gost_style', False)
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
