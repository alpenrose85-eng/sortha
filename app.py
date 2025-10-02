import streamlit as st
import pandas as pd
import re
import io
from typing import List, Dict, Tuple
import docx

# =============================================================================
# ФУНКЦИИ ЧТЕНИЯ ФАЙЛОВ
# =============================================================================

def read_uploaded_file(file):
    """
    Читает содержимое загруженного файла в зависимости от его типа
    """
    try:
        file.seek(0)
        
        if file.name.endswith('.docx'):
            doc = docx.Document(file)
            full_text = []
            
            for paragraph in doc.paragraphs:
                if paragraph.text.strip():
                    full_text.append(paragraph.text)
            
            for table in doc.tables:
                for row in table.rows:
                    row_text = []
                    for cell in row.cells:
                        if cell.text.strip():
                            row_text.append(cell.text)
                    if row_text:
                        full_text.append(" | ".join(row_text))
            
            return "\n".join(full_text)
        else:
            return file.getvalue().decode("utf-8")
            
    except Exception as e:
        st.error(f"Ошибка чтения файла {file.name}: {str(e)}")
        return None

# =============================================================================
# ФУНКЦИИ ОБРАБОТКИ ДАННЫХ
# =============================================================================

def parse_correct_order(file_content: str) -> List[Dict]:
    """
    Парсит правильный порядок образцов из файла с вырезками
    """
    if not file_content:
        return []
        
    lines = file_content.split('\n')
    correct_samples = []
    
    for line in lines:
        line = line.strip()
        
        if re.match(r'^[-]+$', line) or not line:
            continue
            
        clean_line = re.sub(r'\[(.*?)\]\{\.mark\}', r'\1', line)
        
        match = re.match(r'^\s*(\d+)\s+(.+)$', clean_line)
        if match:
            sample_number = int(match.group(1))
            sample_name = match.group(2).strip()
            
            correct_samples.append({
                'order': sample_number,
                'correct_name': sample_name,
                'key': create_sample_key(sample_name)
            })
    
    return correct_samples

def create_sample_key(sample_name: str) -> str:
    """
    Создает ключ для сопоставления образцов из разных источников
    """
    normalized = re.sub(r'\s+', ' ', sample_name.strip()).lower()
    
    # Паттерны для извлечения ключевых компонентов
    patterns = [
        r'([а-я]+)\s*([а-я]+)?\s*\((\d+[-\d]*),\s*([а-я])\)',  # КПП ВД(50,А)
        r'([а-я]+)\s*([а-я]+)?-?(\d+)?\s*\((\d+),\s*([а-я])\)', # КПП НД-1(19,А)
        r'([а-я]+)\s*([а-я]+)?\s*(\d+)',  # НГ ШПП 4
        r'(\d+)[,_]\s*([а-я]+)',  # 28_КПП ВД
    ]
    
    for pattern in patterns:
        match = re.search(pattern, normalized)
        if match:
            parts = [p for p in match.groups() if p]
            return '_'.join(parts)
    
    # Если паттерн не найден, ищем числа в названии
    numbers = re.findall(r'\d+', normalized)
    if numbers:
        return f"sample_{numbers[-1]}"
    
    return normalized

def parse_chemical_tables_simple(file_content: str) -> Dict[str, List[Dict]]:
    """
    ПРОСТОЙ И НАДЕЖНЫЙ парсинг химического анализа
    """
    if not file_content:
        return {}
    
    st.info("🔍 Начинаю анализ файла с химическим анализом...")
    
    lines = file_content.split('\n')
    tables = {}
    current_steel_grade = None
    
    # Сначала найдем все марки стали
    steel_grades = []
    for line in lines:
        steel_match = re.search(r'Марка стали:\s*([^\n]+)', line, re.IGNORECASE)
        if steel_match:
            steel_grade = steel_match.group(1).strip()
            steel_grades.append(steel_grade)
    
    st.write(f"📋 Найдены марки стали: {steel_grades}")
    
    # Для каждой марки стали найдем образцы
    for steel_grade in steel_grades:
        st.write(f"🔍 Ищу образцы для марки: {steel_grade}")
        samples = []
        
        # Ищем начало таблицы для этой марки стали
        start_index = -1
        for i, line in enumerate(lines):
            if steel_grade in line and 'Марка стали' in line:
                start_index = i
                break
        
        if start_index == -1:
            continue
            
        # Ищем образцы после марки стали
        for i in range(start_index + 1, min(start_index + 50, len(lines))):
            line = lines[i].strip()
            
            # Пропускаем технические строки
            if any(x in line for x in ['Требования ТУ', '14-3Р-55-2001', '---', '###']):
                continue
                
            # Ищем строки с образцами - более гибкие критерии
            if (re.match(r'^\s*\d+\s+[а-яА-Я]', line) or 
                re.search(r'[кК][пП][пП]|[шШ][пП][пП]|[нН][гГ]|[нН][бБ]', line)):
                
                # Разные способы разделения данных
                parts = []
                
                # Способ 1: разделение по множественным пробелам
                temp_parts = re.split(r'\s{2,}', line)
                if len(temp_parts) >= 2:
                    parts = temp_parts
                
                # Способ 2: если не сработало, пробуем разделить по одиночным пробелам
                if not parts:
                    temp_parts = line.split()
                    if len(temp_parts) >= 2:
                        parts = temp_parts
                
                if parts and len(parts) >= 2:
                    # Извлекаем название образца (вторая часть)
                    sample_name = parts[1]
                    measurements = parts[2:] if len(parts) > 2 else []
                    
                    sample_data = {
                        'original_name': sample_name,
                        'measurements': measurements,
                        'key': create_sample_key(sample_name)
                    }
                    samples.append(sample_data)
                    st.write(f"   ✅ Найден образец: {sample_name}")
        
        if samples:
            tables[steel_grade] = samples
            st.success(f"🎉 Для марки {steel_grade} найдено {len(samples)} образцов")
        else:
            st.warning(f"⚠️ Для марки {steel_grade} образцы не найдены")
    
    return tables

def create_comparison_table(correct_samples: List[Dict], chemical_tables: Dict[str, List[Dict]]) -> pd.DataFrame:
    """
    Создает таблицу сравнения названий образцов
    """
    data = []
    
    # Собираем все образцы из химического анализа
    all_analysis_samples = []
    for steel_grade, samples in chemical_tables.items():
        for sample in samples:
            all_analysis_samples.append({
                'name': sample['original_name'],
                'key': sample['key'],
                'steel_grade': steel_grade
            })
    
    # Создаем строки для правильных образцов
    for correct in correct_samples:
        data.append({
            'Тип': 'Правильный порядок',
            '№ п/п': correct['order'],
            'Название образца': correct['correct_name'],
            'Ключ': correct['key'],
            'Марка стали': '-'
        })
    
    # Создаем строки для образцов из анализа
    for i, analysis in enumerate(all_analysis_samples):
        data.append({
            'Тип': 'Химический анализ',
            '№ п/п': i + 1,
            'Название образца': analysis['name'],
            'Ключ': analysis['key'],
            'Марка стали': analysis['steel_grade']
        })
    
    return pd.DataFrame(data)

def match_samples_simple(correct_samples: List[Dict], analysis_samples: List[Dict]) -> List[Dict]:
    """
    Простое сопоставление образцов по числам в названиях
    """
    matched = []
    
    for correct in correct_samples:
        correct_numbers = re.findall(r'\d+', correct['correct_name'])
        
        for analysis in analysis_samples:
            analysis_numbers = re.findall(r'\d+', analysis['original_name'])
            
            # Если есть общие числа - считаем что это один образец
            common_numbers = set(correct_numbers) & set(analysis_numbers)
            if common_numbers:
                matched.append({
                    'correct_name': correct['correct_name'],
                    'original_name': analysis['original_name'],
                    'measurements': analysis['measurements'],
                    'order': correct['order'],
                    'common_numbers': list(common_numbers)
                })
    
    return matched

# =============================================================================
# STREAMLIT ИНТЕРФЕЙС
# =============================================================================

st.set_page_config(
    page_title="Сортировка химического анализа",
    page_icon="🔬",
    layout="wide"
)

st.markdown("""
<style>
    .main-header { 
        font-size: 2.5rem; 
        color: #1f77b4; 
        text-align: center; 
        margin-bottom: 2rem;
        font-weight: bold;
    }
    .debug-box {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

def main():
    st.markdown('<div class="main-header">🔬 Автоматическая сортировка химического анализа</div>', unsafe_allow_html=True)
    
    with st.expander("📖 ИНСТРУКЦИЯ", expanded=False):
        st.markdown("""
        **Как использовать:**
        1. Загрузите файл с правильным порядком образцов
        2. Загрузите файл с результатами химического анализа  
        3. Нажмите кнопку "Обработать данные"
        4. Просмотрите таблицу сравнения названий
        5. Скачайте отсортированные результаты

        **Цель:** Сначала увидеть ВСЕ данные, потом научиться их сопоставлять
        """)
    
    # Загрузка файлов
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📋 Файл с правильным порядком")
        correct_order_file = st.file_uploader(
            "Загрузите файл с порядком образцов",
            type=['txt', 'docx'],
            key="correct_order"
        )
        
        if correct_order_file:
            st.success(f"✅ Загружен: {correct_order_file.name}")
    
    with col2:
        st.subheader("🧪 Файл с химическим анализом")
        chemical_analysis_file = st.file_uploader(
            "Загрузите файл с анализом", 
            type=['txt', 'docx'],
            key="chemical_analysis"
        )
        
        if chemical_analysis_file:
            st.success(f"✅ Загружен: {chemical_analysis_file.name}")
    
    # Показ сырых данных
    show_raw_data = st.checkbox("📄 Показать сырые данные файлов")
    
    # Кнопка обработки
    if st.button("🚀 ОБРАБОТАТЬ ДАННЫЕ", type="primary", use_container_width=True):
        if not correct_order_file or not chemical_analysis_file:
            st.error("❌ Загрузите оба файла")
            return
        
        try:
            # Чтение файлов
            with st.spinner("📖 Читаю файлы..."):
                correct_order_content = read_uploaded_file(correct_order_file)
                chemical_analysis_content = read_uploaded_file(chemical_analysis_file)
            
            if not correct_order_content or not chemical_analysis_content:
                st.error("❌ Ошибка чтения файлов")
                return
            
            # Показ сырых данных если нужно
            if show_raw_data:
                st.subheader("📄 СЫРЫЕ ДАННЫЕ ФАЙЛОВ")
                col1, col2 = st.columns(2)
                with col1:
                    st.text_area("Файл правильного порядка:", correct_order_content, height=300)
                with col2:
                    st.text_area("Файл химического анализа:", chemical_analysis_content, height=300)
            
            # Парсинг правильного порядка
            with st.spinner("🔍 Анализирую правильный порядок..."):
                correct_samples = parse_correct_order(correct_order_content)
            
            if not correct_samples:
                st.error("❌ Не найдены образцы в файле порядка")
                return
            
            st.success(f"✅ В правильном порядке найдено {len(correct_samples)} образцов")
            
            # Парсинг химического анализа
            with st.spinner("🔍 Анализирую химический анализ..."):
                chemical_tables = parse_chemical_tables_simple(chemical_analysis_content)
            
            if not chemical_tables:
                st.error("❌ Не найдены таблицы анализа")
                
                # Покажем отладочную информацию
                st.markdown('<div class="debug-box">', unsafe_allow_html=True)
                st.write("**Отладочная информация по файлу анализа:**")
                
                lines = chemical_analysis_content.split('\n')
                potential_samples = []
                
                for i, line in enumerate(lines):
                    if re.search(r'[кК][пП][пП]|[шШ][пП][пП]|[нН][гГ]|[нН][бБ]|\d+,\d+', line):
                        potential_samples.append((i, line))
                
                st.write(f"Найдено потенциальных строк: {len(potential_samples)}")
                for i, (line_num, line) in enumerate(potential_samples[:20]):
                    st.write(f"{line_num}: {line}")
                st.markdown('</div>', unsafe_allow_html=True)
                return
            
            total_analysis_samples = sum(len(samples) for samples in chemical_tables.values())
            st.success(f"✅ В химическом анализе найдено {total_analysis_samples} образцов")
            
            # ПОКАЗЫВАЕМ ТАБЛИЦУ СРАВНЕНИЯ
            st.subheader("🔍 ТАБЛИЦА СРАВНЕНИЯ НАЗВАНИЙ")
            
            comparison_df = create_comparison_table(correct_samples, chemical_tables)
            st.dataframe(comparison_df, use_container_width=True)
            
            # Статистика
            st.subheader("📊 СТАТИСТИКА")
            col1, col2 = st.columns(2)
            
            with col1:
                st.metric("Образцов в правильном порядке", len(correct_samples))
            with col2:
                st.metric("Образцов в анализе", total_analysis_samples)
            
            # Показываем детальную информацию о ключах
            st.subheader("🔑 КЛЮЧИ СОПОСТАВЛЕНИЯ")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Из правильного порядка:**")
                for sample in correct_samples:
                    st.write(f"`{sample['key']}` → \"{sample['correct_name']}\"")
            
            with col2:
                st.write("**Из химического анализа:**")
                all_analysis_samples = []
                for steel_grade, samples in chemical_tables.items():
                    all_analysis_samples.extend(samples)
                for sample in all_analysis_samples:
                    st.write(f"`{sample['key']}` → \"{sample['original_name']}\"")
            
            # Пробуем простое сопоставление по числам
            st.subheader("🔄 ПРОБУЕМ СОПОСТАВИТЬ ПО ЧИСЛАМ")
            
            all_analysis_samples = []
            for steel_grade, samples in chemical_tables.items():
                all_analysis_samples.extend(samples)
            
            matched_samples = match_samples_simple(correct_samples, all_analysis_samples)
            
            if matched_samples:
                st.success(f"✅ Найдено {len(matched_samples)} возможных сопоставлений!")
                
                # Показываем таблицу сопоставления
                matched_data = []
                for sample in matched_samples:
                    matched_data.append({
                        '№ п/п': sample['order'],
                        'Правильное название': sample['correct_name'],
                        'Исходное название': sample['original_name'],
                        'Общие числа': ', '.join(map(str, sample['common_numbers']))
                    })
                
                st.dataframe(pd.DataFrame(matched_data), use_container_width=True)
                
                # Создаем отсортированные таблицы
                final_tables = {}
                for steel_grade, samples in chemical_tables.items():
                    # Для каждой марки стали создаем таблицу с сопоставленными образцами
                    steel_matched = [s for s in matched_samples 
                                   if any(analysis['original_name'] == s['original_name'] 
                                         for analysis in samples)]
                    
                    if steel_matched:
                        # Сортируем по порядку из правильного списка
                        steel_matched.sort(key=lambda x: x['order'])
                        
                        # Создаем DataFrame
                        if steel_matched and steel_matched[0]['measurements']:
                            num_measurements = len(steel_matched[0]['measurements'])
                            columns = ['№', 'Образец'] + [f'Измерение {i+1}' for i in range(num_measurements)]
                            
                            data = []
                            for i, sample in enumerate(steel_matched):
                                row = [i+1, sample['correct_name']] + sample['measurements']
                                data.append(row)
                            
                            final_tables[steel_grade] = pd.DataFrame(data, columns=columns)
                
                # Показываем отсортированные результаты
                if final_tables:
                    st.subheader("📋 ОТСОРТИРОВАННЫЕ РЕЗУЛЬТАТЫ")
                    
                    for steel_grade, table in final_tables.items():
                        with st.expander(f"🔩 {steel_grade} ({len(table)} образцов)", expanded=True):
                            st.dataframe(table, use_container_width=True)
                            
                            # Скачивание
                            excel_buffer = io.BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                table.to_excel(writer, index=False, sheet_name=steel_grade[:30])
                            excel_buffer.seek(0)
                            
                            st.download_button(
                                label="📥 Скачать Excel",
                                data=excel_buffer,
                                file_name=f"отсортировано_{steel_grade}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"download_{steel_grade}"
                            )
            else:
                st.warning("""
                ⚠️ **Не удалось автоматически сопоставить образцы**
                
                **Что видим из таблицы:**
                - Слева - правильные названия из файла порядка
                - Справа - названия из химического анализа
                - В столбце "Ключ" видно как программа понимает названия
                
                **Следующие шаги:**
                1. Сравните названия вручную
                2. Определите по каким правилам они должны сопоставляться
                3. Мы улучшим алгоритм на основе этих правил
                """)
                
        except Exception as e:
            st.error(f"❌ Ошибка: {str(e)}")
            import traceback
            st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
