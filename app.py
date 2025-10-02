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
    
    patterns = [
        r'([а-я]+)\s*([а-я]+)?\s*\((\d+[-\d]*),\s*([а-я])\)',
        r'([а-я]+)\s*([а-я]+)?-?(\d+)?\s*\((\d+),\s*([а-я])\)',
        r'(\d+)[,_]\s*([а-я]+)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, normalized)
        if match:
            parts = [p for p in match.groups() if p]
            return '_'.join(parts)
    
    numbers = re.findall(r'\d+', normalized)
    if numbers:
        return f"sample_{numbers[-1]}"
    
    return normalized

def parse_chemical_tables(file_content: str) -> Dict[str, List[Dict]]:
    """
    Парсит все таблицы химического анализа из файла
    """
    if not file_content:
        return {}
        
    lines = file_content.split('\n')
    tables = {}
    current_steel_grade = None
    current_table = []
    
    for line in lines:
        line = line.strip()
        
        # Ищем марку стали
        steel_match = re.search(r'Марка стали:\s*([^\n]+)', line, re.IGNORECASE)
        if steel_match:
            if current_steel_grade and current_table:
                tables[current_steel_grade] = current_table
                current_table = []
            current_steel_grade = steel_match.group(1).strip()
            continue
        
        # Более гибкий поиск строк с образцами
        if current_steel_grade:
            # Пропускаем технические строки
            if any(x in line for x in ['Требования ТУ', '14-3Р-55-2001', '---', '###']):
                continue
            
            # Ищем строки, которые содержат номер и название образца
            # Более гибкий паттерн: число, потом текст, потом числа с запятыми
            if re.match(r'^\s*\d+\s+[^\d]', line) and re.search(r'\d+,\d+', line):
                # Пробуем разные способы разделения
                parts = []
                
                # Попробуем разделить по множественным пробелам
                temp_parts = re.split(r'\s{2,}', line)
                if len(temp_parts) >= 3:
                    parts = temp_parts
                else:
                    # Если не получилось, пробуем разделить по одиночным пробелам
                    temp_parts = line.split()
                    if len(temp_parts) >= 3:
                        # Найдем где заканчивается название и начинаются измерения
                        # Измерения - это числа с запятыми
                        for i in range(2, len(temp_parts)):
                            if re.match(r'^\d+,\d+$', temp_parts[i]):
                                name_parts = temp_parts[1:i]
                                measurement_parts = temp_parts[i:]
                                parts = [temp_parts[0], ' '.join(name_parts)] + measurement_parts
                                break
                
                if len(parts) >= 3:
                    sample_data = {
                        'original_name': parts[1],
                        'measurements': parts[2:],
                        'key': create_sample_key(parts[1])
                    }
                    current_table.append(sample_data)
    
    # Добавляем последнюю таблицу
    if current_steel_grade and current_table:
        tables[current_steel_grade] = current_table
    
    return tables

def match_and_sort_samples(original_samples: List[Dict], correct_samples: List[Dict]) -> List[Dict]:
    """
    Сопоставляет и сортирует образцы по правильному порядку
    """
    key_to_correct = {}
    for correct in correct_samples:
        key_to_correct[correct['key']] = {
            'correct_name': correct['correct_name'],
            'order': correct['order']
        }
    
    matched_samples = []
    for original in original_samples:
        if original['key'] in key_to_correct:
            matched_samples.append({
                'correct_name': key_to_correct[original['key']]['correct_name'],
                'measurements': original['measurements'],
                'order': key_to_correct[original['key']]['order'],
                'original_name': original['original_name'],
                'key': original['key']
            })
    
    matched_samples.sort(key=lambda x: x['order'])
    return matched_samples

def create_matching_table(matched_samples: List[Dict]) -> pd.DataFrame:
    """
    Создает таблицу сопоставления названий
    """
    data = []
    for sample in matched_samples:
        data.append({
            '№ п/п': sample['order'],
            'Правильное название': sample['correct_name'],
            'Название из анализа': sample['original_name'],
            'Ключ сопоставления': sample['key']
        })
    
    return pd.DataFrame(data)

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
        4. Просмотрите таблицу сопоставления и отсортированные результаты
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
    
    # Отладочная информация
    debug_mode = st.checkbox("🔧 Режим отладки (показать сырые данные)")
    
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
            
            if debug_mode:
                st.subheader("🔍 СЫРЫЕ ДАННЫЕ")
                col1, col2 = st.columns(2)
                with col1:
                    st.text_area("Файл порядка:", correct_order_content, height=200)
                with col2:
                    st.text_area("Файл анализа:", chemical_analysis_content, height=200)
            
            # Парсинг
            with st.spinner("🔍 Анализирую данные..."):
                correct_samples = parse_correct_order(correct_order_content)
                chemical_tables = parse_chemical_tables(chemical_analysis_content)
            
            if not correct_samples:
                st.error("❌ Не найдены образцы в файле порядка")
                return
                
            if not chemical_tables:
                st.error("❌ Не найдены таблицы анализа")
                
                # Показываем отладочную информацию
                st.markdown('<div class="debug-box">', unsafe_allow_html=True)
                st.write("**Отладочная информация:**")
                st.write("Попробую найти образцы вручную...")
                
                # Альтернативный парсинг
                lines = chemical_analysis_content.split('\n')
                found_samples = []
                for i, line in enumerate(lines):
                    if re.search(r'КПП|труба|\d+,\d+', line) and not any(x in line for x in ['Требования', 'Марка стали']):
                        st.write(f"Строка {i}: {line[:100]}...")
                        found_samples.append(line)
                
                if found_samples:
                    st.write(f"Найдено потенциальных строк: {len(found_samples)}")
                st.markdown('</div>', unsafe_allow_html=True)
                return
            
            # Отладочная информация о найденных данных
            if debug_mode:
                st.subheader("🔍 ПАРСИНГ ДАННЫХ")
                st.write(f"Найдено образцов в правильном порядке: {len(correct_samples)}")
                for sample in correct_samples[:5]:
                    st.write(f" - №{sample['order']}: {sample['correct_name']} → ключ: {sample['key']}")
                
                st.write(f"Найдено таблиц анализа: {len(chemical_tables)}")
                for steel_grade, samples in chemical_tables.items():
                    st.write(f" - {steel_grade}: {len(samples)} образцов")
                    for sample in samples[:3]:
                        st.write(f"   * {sample['original_name']} → ключ: {sample['key']}")
            
            # Обработка
            with st.spinner("🔄 Сопоставляю и сортирую..."):
                final_tables = {}
                all_matched_samples = []
                
                for steel_grade, samples in chemical_tables.items():
                    sorted_samples = match_and_sort_samples(samples, correct_samples)
                    if sorted_samples:
                        num_measurements = len(sorted_samples[0]['measurements'])
                        columns = ['№', 'Образец'] + [f'Измерение {i+1}' for i in range(num_measurements)]
                        
                        data = []
                        for i, sample in enumerate(sorted_samples):
                            row = [i+1, sample['correct_name']] + sample['measurements']
                            data.append(row)
                        
                        final_tables[steel_grade] = pd.DataFrame(data, columns=columns)
                        all_matched_samples.extend(sorted_samples)
            
            # Статистика
            total_matched = len(all_matched_samples)
            total_chemical = sum(len(samples) for samples in chemical_tables.values())
            total_correct = len(correct_samples)
            matching_rate = (total_matched / total_chemical) * 100 if total_chemical > 0 else 0
            
            # Показываем статистику
            st.subheader("📊 СТАТИСТИКА")
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Образцов в порядке", total_correct)
            with col2:
                st.metric("Образцов в анализе", total_chemical)
            with col3:
                st.metric("Сопоставлено", total_matched)
            with col4:
                st.metric("Процент сопоставления", f"{matching_rate:.1f}%")
            
            # ТАБЛИЦА СОПОСТАВЛЕНИЯ
            if all_matched_samples:
                st.subheader("🔍 ТАБЛИЦА СОПОСТАВЛЕНИЯ")
                matching_df = create_matching_table(all_matched_samples)
                st.dataframe(matching_df, use_container_width=True)
            
            # ОТСОРТИРОВАННЫЕ РЕЗУЛЬТАТЫ
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
                
                # Пакетное скачивание
                if len(final_tables) > 1:
                    st.subheader("💾 ПАКЕТНОЕ СКАЧИВАНИЕ")
                    excel_buffer_all = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer_all, engine='openpyxl') as writer:
                        for steel_grade, table in final_tables.items():
                            table.to_excel(writer, index=False, sheet_name=steel_grade[:30])
                    excel_buffer_all.seek(0)
                    
                    st.download_button(
                        label="📦 Скачать все таблицы (Excel)",
                        data=ex_buffer_all,
                        file_name="все_отсортированные_таблицы.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_all",
                        use_container_width=True
                    )
            else:
                st.warning("⚠️ Не удалось сопоставить образцы")
                
        except Exception as e:
            st.error(f"❌ Ошибка: {str(e)}")
            import traceback
            st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
