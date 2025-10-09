import streamlit as st
import re
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
        'doi_format': 'DOI format:',
        'doi_hyperlink': 'DOI as hyperlink',
        'page_format': 'Pages:',
        'final_punctuation': 'Punctuation:',
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
        'language': 'Language:'
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
        'doi_format': 'Формат DOI:',
        'doi_hyperlink': 'DOI как ссылка',
        'page_format': 'Страницы:',
        'final_punctuation': 'Пунктуация:',
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
        'language': 'Язык:'
    }
}

# Хранение текущего языка
if 'current_language' not in st.session_state:
    st.session_state.current_language = 'ru'

def get_text(key):
    return TRANSLATIONS[st.session_state.current_language].get(key, key)

# Функции обработки (без изменений)
def clean_text(text):
    return re.sub(r'<[^>]+>|&[^;]+;', '', text).strip()

def normalize_name(name):
    return name[0].upper() + name[1:].lower() if name and len(name) > 1 else name.upper() if name else ''

def is_section_header(text):
    text_upper = text.upper().strip()
    section_patterns = [
        r'^NOTES?\s+AND\s+REFERENCES?', r'^REFERENCES?', r'^BIBLIOGRAPHY', r'^LITERATURE',
        r'^WORKS?\s+CITED', r'^SOURCES?', r'^CHAPTER\s+\d+', r'^SECTION\s+\d+', r'^PART\s+\d+'
    ]
    for pattern in section_patterns:
        if re.search(pattern, text_upper):
            return True
    if len(text.strip()) < 50 and len(text.strip().split()) <= 5:
        return True
    return False

def find_doi(reference):
    if is_section_header(reference):
        return None
    doi_pattern = r'(?:(?:https?://doi\.org/)|(?:doi:|DOI:))?(\d+\.\d+/[^\s,;]+)'
    match = re.search(doi_pattern, reference)
    if match:
        return match.group(1).rstrip('.')
    clean_ref = re.sub(r'\s*https?://doi\.org/[^\s]+', '', reference)
    clean_ref = re.sub(r'\s*DOI:\s*[^\s]+', '', clean_ref).strip()
    if len(clean_ref) < 30:
        return None
    try:
        query = works.query(bibliographic=clean_ref).sort('relevance').order('desc')
        for result in query:
            if 'DOI' in result:
                return result['DOI']
    except:
        return None
    return None

def extract_metadata(doi):
    try:
        result = works.doi(doi)
        if not result:
            return None
        authors = result.get('author', [])
        author_list = [{'given': a.get('given', ''), 'family': normalize_name(a.get('family', ''))} for a in authors]
        return {
            'authors': author_list,
            'title': clean_text(result.get('title', [''])[0]),
            'journal': result.get('container-title', [''])[0],
            'year': result.get('published', {}).get('date-parts', [[None]])[0][0],
            'volume': result.get('volume', ''),
            'issue': result.get('issue', ''),
            'pages': result.get('page', ''),
            'article_number': result.get('article-number', ''),
            'doi': doi
        }
    except:
        return None

def format_authors(authors, author_format, separator, et_al_limit, use_and):
    if not authors:
        return ""
    author_str = ""
    limit = et_al_limit if et_al_limit and not use_and else len(authors)
    for i, author in enumerate(authors[:limit]):
        given = author['given']
        family = author['family']
        initials = given.split()[:2]
        first_initial = initials[0][0] if initials else ''
        second_initial = initials[1][0].upper() if len(initials) > 1 else ''
        if author_format == "AA Smith":
            author_str += f"{first_initial}{second_initial} {family}"
        elif author_format == "A.A. Smith":
            author_str += f"{first_initial}.{second_initial}. {family}" if second_initial else f"{first_initial}. {family}"
        elif author_format == "Smith AA":
            author_str += f"{family} {first_initial}{second_initial}"
        elif author_format == "Smith A.A":
            author_str += f"{family} {first_initial}.{second_initial}." if second_initial else f"{family} {first_initial}."
        elif author_format == "Smith, A.A.":
            author_str += f"{family}, {first_initial}.{second_initial}." if second_initial else f"{family}, {first_initial}."
        if i < len(authors[:limit]) - 1:
            if i == len(authors[:limit]) - 2 and use_and:
                author_str += " and " if st.session_state.current_language == 'en' else " и "
            else:
                author_str += separator
    if et_al_limit and len(authors) > et_al_limit and not use_and:
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
    return article_number

def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0000FF')
    rPr.append(color)
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
        return ("Ошибка: Не удалось отформатировать ссылку." if st.session_state.current_language == 'ru' else "Error: Could not format the reference.", True)
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
                style_config['use_and']
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
            if config['parentheses'] and value:
                value = f"({value})"
            separator = config['separator'] if i < len(style_config['elements']) - 1 else ''
            if for_preview:
                formatted_value = value
                if config['italic']:
                    formatted_value = f"<i>{formatted_value}</i>"
                if config['bold']:
                    formatted_value = f"<b>{formatted_value}</b>"
                elements.append((formatted_value, False, False, separator, False, None))
            else:
                elements.append((value, config['italic'], config['bold'], separator,
                               (element == "DOI" and style_config['doi_hyperlink']), doi_value))
    if for_preview:
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

def apply_yellow_background(run):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), 'FFFF00')
    run._element.get_or_add_rPr().append(shd)

def process_references(references, style_config):
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
        if doi:
            doi_list.append(doi)
            metadata = extract_metadata(doi)
            if metadata:
                formatted_ref, is_error = format_reference(metadata, style_config)
                formatted_refs.append((formatted_ref, is_error, metadata))
                if not is_error:
                    doi_found_count += 1
                else:
                    doi_list[-1] = f"{ref}\nПроверьте источник и добавьте DOI вручную." if st.session_state.current_language == 'ru' else f"{ref}\nPlease check this source and insert the DOI manually."
                    formatted_refs.append((f"{ref} Проверьте источник и добавьте DOI вручную." if st.session_state.current_language == 'ru' else f"{ref} Please check this source and insert the DOI manually.", True, None))
                    doi_not_found_count += 1
            else:
                doi_list[-1] = f"{ref}\nПроверьте источник и добавьте DOI вручную." if st.session_state.current_language == 'ru' else f"{ref}\nPlease check this source and insert the DOI manually."
                formatted_refs.append((f"{ref} Проверьте источник и добавьте DOI вручную." if st.session_state.current_language == 'ru' else f"{ref} Please check this source and insert the DOI manually.", True, None))
                doi_not_found_count += 1
        else:
            doi_list.append(f"{ref}\nПроверьте источник и добавьте DOI вручную." if st.session_state.current_language == 'ru' else f"{ref}\nPlease check this source and insert the DOI manually.")
            formatted_refs.append((f"{ref} Проверьте источник и добавьте DOI вручную." if st.session_state.current_language == 'ru' else f"{ref} Please check this source and insert the DOI manually.", True, None))
            doi_not_found_count += 1
        progress_bar.update(1)
    progress_bar.close()
    st.write(f"**{get_text('statistics').format(doi_found_count, doi_not_found_count)}**")
    output_txt_buffer = io.StringIO()
    for doi in doi_list:
        output_txt_buffer.write(f"{doi}\n")
    output_txt_buffer.seek(0)
    txt_bytes = io.BytesIO(output_txt_buffer.getvalue().encode('utf-8'))
    return formatted_refs, txt_bytes, doi_found_count, doi_not_found_count

def process_docx(input_file, style_config):
    doc = Document(input_file)
    references = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
    st.write(f"**{get_text('found_references').format(len(references))}**")
    formatted_refs, txt_bytes, doi_found_count, doi_not_found_count = process_references(references, style_config)
    output_doc = Document()
    output_doc.add_heading('References in Custom Style' if st.session_state.current_language == 'en' else 'Ссылки в пользовательском стиле', level=1)
    for i, (elements, is_error, metadata) in enumerate(formatted_refs, 1):
        numbering = style_config['numbering_style']
        prefix = "" if numbering == "No numbering" else f"{i}{numbering[-1] if numbering != '1' else ''} "
        para = output_doc.add_paragraph(prefix)
        if is_error:
            run = para.add_run(str(elements))
            apply_yellow_background(run)
        else:
            if metadata is None:
                run = para.add_run(str(elements))
                run.font.italic = True
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
    return formatted_refs, txt_bytes, output_doc_buffer

# Компактный интерфейс Streamlit
def main():
    st.set_page_config(layout="wide")
    st.markdown("""
        <style>
        .block-container { padding: 0.3rem; }
        .stSelectbox, .stTextInput, .stNumberInput, .stCheckbox, .stRadio, .stFileUploader, .stTextArea {
            margin-bottom: 0.05rem;
        }
        .stTextArea { height: 50px !important; }
        .stButton > button { width: 100%; padding: 0.1rem; font-size: 0.8rem; }
        h1 { font-size: 1.2rem; margin-bottom: 0.1rem; }
        h3 { font-size: 0.9rem; margin-bottom: 0.05rem; }
        label { font-size: 0.75rem !important; }
        .stMarkdown { font-size: 0.75rem; }
        .stCheckbox > label { font-size: 0.7rem; }
        </style>
    """, unsafe_allow_html=True)

    # Переключение языка
    language = st.selectbox(get_text('language'), [('Русский', 'ru'), ('English', 'en')], 
                            format_func=lambda x: x[0], 
                            index=0 if st.session_state.current_language == 'ru' else 1,
                            label_visibility="visible")
    st.session_state.current_language = language[1]

    st.title(get_text('header'), help="Выберите настройки, укажите элементы, загрузите DOCX или введите ссылки, затем нажмите 'Обработать'.")

    # Трёхколоночный макет
    col1, col2, col3 = st.columns([1, 1, 1])

    with col1:
        st.subheader(get_text('general_settings'))
        numbering_style = st.selectbox(get_text('numbering_style'), ["No numbering", "1", "1.", "1)", "(1)", "[1]"], key="num")
        author_format = st.selectbox(get_text('author_format'), ["AA Smith", "A.A. Smith", "Smith AA", "Smith A.A", "Smith, A.A."], key="auth")
        author_separator = st.selectbox(get_text('author_separator'), [", ", "; "], key="sep")
        et_al_limit = st.number_input(get_text('et_al_limit'), min_value=0, step=1, key="etal")
        use_and = st.checkbox(get_text('use_and'), key="and")
        doi_format = st.selectbox(get_text('doi_format'), ["10.10/xxx", "doi:10.10/xxx", "DOI:10.10/xxx", "https://dx.doi.org/10.10/xxx"], key="doi")
        doi_hyperlink = st.checkbox(get_text('doi_hyperlink'), key="doilink")
        page_format = st.selectbox(get_text('page_format'), ["122 - 128", "122-128", "122 – 128", "122–128", "122–8"], key="page")
        final_punctuation = st.selectbox(get_text('final_punctuation'), ["", "."], key="punct")

    with col2:
        st.subheader(get_text('element_config'))
        available_elements = ["", "Authors", "Title", "Journal", "Year", "Volume", "Issue", "Pages", "DOI"]
        element_configs = []
        used_elements = set()
        st.markdown(f"<small>{get_text('element')} | {get_text('italic')} | {get_text('bold')} | {get_text('parentheses')} | {get_text('separator')}</small>", unsafe_allow_html=True)
        for i in range(8):  # Все 8 элементов
            cols = st.columns([2, 1, 1, 1, 2])
            with cols[0]:
                element = st.selectbox("", available_elements, key=f"el{i}", label_visibility="collapsed")
            with cols[1]:
                italic = st.checkbox(get_text('italic'), key=f"it{i}")
            with cols[2]:
                bold = st.checkbox(get_text('bold'), key=f"bd{i}")
            with cols[3]:
                parentheses = st.checkbox(get_text('parentheses'), key=f"pr{i}")
            with cols[4]:
                separator = st.text_input("", value=". ", key=f"sp{i}", label_visibility="collapsed")
            if element and element not in used_elements:
                element_configs.append((element, {'italic': italic, 'bold': bold, 'parentheses': parentheses, 'separator': separator}))
                used_elements.add(element)

    with col3:
        # Предпросмотр
        st.subheader(get_text('style_preview'))
        style_config = {
            'author_format': author_format,
            'author_separator': author_separator,
            'et_al_limit': et_al_limit if et_al_limit > 0 else None,
            'use_and': use_and,
            'doi_format': doi_format,
            'doi_hyperlink': doi_hyperlink,
            'page_format': page_format,
            'final_punctuation': final_punctuation,
            'numbering_style': numbering_style,
            'elements': element_configs
        }
        if not style_config['elements']:
            st.markdown(f"<b style='color:red;'>{get_text('error_select_element')}</b>", unsafe_allow_html=True)
        else:
            preview_metadata = {
                'authors': [{'given': 'John A.' if st.session_state.current_language == 'en' else 'Иван А.', 'family': 'Smith' if st.session_state.current_language == 'en' else 'Иванов'}, 
                            {'given': 'Alice B.' if st.session_state.current_language == 'en' else 'Анна Б.', 'family': 'Doe' if st.session_state.current_language == 'en' else 'Петрова'}],
                'title': 'Article Title' if st.session_state.current_language == 'en' else 'Название статьи',
                'journal': 'Journal Name' if st.session_state.current_language == 'en' else 'Название журнала',
                'year': 2020,
                'volume': '15',
                'issue': '3',
                'pages': '122-128',
                'article_number': 'e12345',
                'doi': '10.1000/xyz123'
            }
            preview_ref, _ = format_reference(preview_metadata, style_config, for_preview=True)
            numbering = style_config['numbering_style']
            preview_ref = preview_ref if numbering == "No numbering" else f"1{numbering[-1] if numbering != '1' else ''} {preview_ref}"
            st.markdown(f"<small>{get_text('example')} {preview_ref}</small>", unsafe_allow_html=True)

        # Ввод
        st.subheader(get_text('data_input'))
        input_method = st.radio(get_text('input_method'), ['DOCX', 'Text' if st.session_state.current_language == 'en' else 'Текст'], horizontal=True)
        if input_method == 'DOCX':
            uploaded_file = st.file_uploader(get_text('select_docx'), type=['docx'], label_visibility="collapsed")
        else:
            references_input = st.text_area(get_text('references'), placeholder=get_text('enter_references'), height=50)

        # Вывод
        st.subheader(get_text('data_output'))
        output_method = st.radio(get_text('output_method'), ['DOCX', 'Text' if st.session_state.current_language == 'en' else 'Текст'], horizontal=True)
        output_text = st.text_area(get_text('results'), placeholder=get_text('results'), height=50, disabled=True)

        # Кнопка обработки
        if st.button(get_text('process')):
            if not style_config['elements']:
                st.error(get_text('error_select_element'))
                return
            if input_method == 'DOCX':
                if not uploaded_file:
                    st.error(get_text('upload_file'))
                    return
                formatted_refs, txt_bytes, output_doc_buffer = process_docx(uploaded_file, style_config)
            else:
                if not references_input.strip():
                    st.error(get_text('enter_references_error'))
                    return
                references = [ref.strip() for ref in references_input.split('\n') if ref.strip()]
                st.write(f"**{get_text('found_references_text').format(len(references))}**")
                formatted_refs, txt_bytes, _, _ = process_references(references, style_config)
                output_doc = Document()
                output_doc.add_heading('References in Custom Style' if st.session_state.current_language == 'en' else 'Ссылки в пользовательском стиле', level=1)
                for i, (elements, is_error, metadata) in enumerate(formatted_refs, 1):
                    numbering = style_config['numbering_style']
                    prefix = "" if numbering == "No numbering" else f"{i}{numbering[-1] if numbering != '1' else ''} "
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

            if output_method == 'Text' if st.session_state.current_language == 'en' else 'Текст':
                output_text_value = ""
                for i, (elements, is_error, metadata) in enumerate(formatted_refs, 1):
                    numbering = style_config['numbering_style']
                    prefix = "" if numbering == "No numbering" else f"{i}{numbering[-1] if numbering != '1' else ''} "
                    if is_error:
                        output_text_value += f"{prefix}{elements}\n"
                    else:
                        ref_str = ""
                        for j, (value, _, _, separator, _, _) in enumerate(elements):
                            ref_str += value
                            if separator and j < len(elements) - 1:
                                ref_str += separator
                            elif j == len(elements) - 1 and style_config['final_punctuation']:
                                ref_str = ref_str.rstrip(',.') + "."
                        output_text_value += f"{prefix}{ref_str}\n"
                st.session_state['output_text'] = output_text_value
            else:
                st.session_state['output_text'] = ""

            # Кнопки скачивания
            st.download_button(
                label=get_text('doi_txt'),
                data=txt_bytes,
                file_name='doi_list.txt',
                mime='text/plain'
            )
            if output_method == 'DOCX':
                st.download_button(
                    label=get_text('references_docx'),
                    data=output_doc_buffer,
                    file_name='references_custom.docx',
                    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                )
            else:
                st.error(get_text('select_docx_output'))

    # Обновление текстового поля результатов
    if 'output_text' in st.session_state:
        st.text_area(get_text('results'), value=st.session_state['output_text'], height=50, disabled=True)

if __name__ == "__main__":
    main()

