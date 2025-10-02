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
    
    # Паттерны для правильных названий (из файла порядка)
    patterns_correct = [
        r'([а-я]+)\s*([а-я]+)?\s*\((\d+[-\d]*),\s*([а-я])\)',  # КПП ВД(50,А)
        r'([а-я]+)\s*([а-я]+)?-?(\d+)?\s*\((\d+),\s*([а-я])\)', # КПП НД-1(19,А)
    ]
    
    # Паттерны для названий из химического анализа
    patterns_analysis = [
        r'([а-я]+)\s*([а-я]+)?\s*(\d+)',  # НГ ШПП 4
        r'(\d+)[,_]\s*([а-я]+)',  # 28_КПП ВД
    ]
    
    # Сначала пробуем паттерны для правильных названий
    for pattern in patterns_correct:
        match = re.search(pattern, normalized)
        if match:
            parts = [p for p in match.groups() if p]
            return '_'.join(parts)
    
    # Затем пробуем паттерны для названий анализа
    for pattern in patterns_analysis:
        match = re.search(pattern, normalized)
        if match:
            parts = [p for p in match.groups() if p]
            return '_'.join(parts)
    
    # Если паттерн не найден, ищем числа в названии
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
        
        # Пропускаем технические строки
        if any(x in line for x in ['Требования ТУ', '14-3Р-55-2001', '---', '###']):
            continue
        
        # Более простой и надежный поиск строк с образцами
        if current_steel_grade and re.search(r'[кК][пП][пП]|[шШ][пП][пП]|[эЭ][пП][кК]', line):
            # Ищем строки, которые содержат номер и название образца
            if re.match(r'^\s*\d+\s+[^\d]', line) and re.search(r'\d', line):
                # Разделяем по множественным пробелам
                parts = re.split(r'\s{2,}', line)
                if len(parts) >= 2:
                    # Берем только название образца (вторая часть)
                    sample_name = parts[1]
                    measurements = parts[2:] if len(parts) > 2 else []
                    
                    sample_data = {
                        'original_name': sample_name,
                        'measurements': measurements,
                        'key': create_sample_key(sample_name)
                    }
                    current_table.append(sample_data)
    
    # Добавляем последнюю таблицу
    if current_steel_grade and current_table:
        tables[current_steel_grade] = current_table
    
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
    .comparison-table {
        margin: 20px 0;
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

        **Следующий шаг:** Будем улучшать алгоритм сопоставления на основе увиденных данных
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
            
            # Парсинг
            with st.spinner("🔍 Анализирую данные..."):
                correct_samples = parse_correct_order(correct_order_content)
                chemical_tables = parse_chemical_tables(chemical_analysis_content)
            
            if not correct_samples:
                st.error("❌ Не найдены образцы в файле порядка")
                return
                
            if not chemical_tables:
                st.error("❌ Не найдены таблицы анализа")
                return
            
            # ПОКАЗЫВАЕМ ТАБЛИЦУ СРАВНЕНИЯ
            st.subheader("🔍 ТАБЛИЦА СРАВНЕНИЯ НАЗВАНИЙ")
            
            comparison_df = create_comparison_table(correct_samples, chemical_tables)
            st.dataframe(comparison_df, use_container_width=True)
            
            # Статистика
            total_correct = len(correct_samples)
            total_analysis = sum(len(samples) for samples in chemical_tables.values())
            
            st.subheader("📊 СТАТИСТИКА")
            col1, col2 = st.columns(2)
            
            with col1:
                st.metric("Образцов в правильном порядке", total_correct)
            with col2:
                st.metric("Образцов в анализе", total_analysis)
            
            # Показываем детальную информацию о ключах
            st.subheader("🔑 ИНФОРМАЦИЯ О КЛЮЧАХ СОПОСТАВЛЕНИЯ")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Ключи из правильного порядка:**")
                for sample in correct_samples[:10]:  # Показываем первые 10
                    st.write(f"`{sample['key']}` → \"{sample['correct_name']}\"")
            
            with col2:
                st.write("**Ключи из химического анализа:**")
                all_analysis_samples = []
                for steel_grade, samples in chemical_tables.items():
                    all_analysis_samples.extend(samples)
                for sample in all_analysis_samples[:10]:  # Показываем первые 10
                    st.write(f"`{sample['key']}` → \"{sample['original_name']}\"")
            
            # Пробуем сопоставить
            st.subheader("🔄 ПОПЫТКА СОПОСТАВЛЕНИЯ")
            
            all_matched_samples = []
            final_tables = {}
            
            for steel_grade, samples in chemical_tables.items():
                sorted_samples = match_and_sort_samples(samples, correct_samples)
                all_matched_samples.extend(sorted_samples)
                
                if sorted_samples:
                    num_measurements = len(sorted_samples[0]['measurements'])
                    columns = ['№', 'Образец'] + [f'Измерение {i+1}' for i in range(num_measurements)]
                    
                    data = []
                    for i, sample in enumerate(sorted_samples):
                        row = [i+1, sample['correct_name']] + sample['measurements']
                        data.append(row)
                    
                    final_tables[steel_grade] = pd.DataFrame(data, columns=columns)
            
            total_matched = len(all_matched_samples)
            matching_rate = (total_matched / total_analysis) * 100 if total_analysis > 0 else 0
            
            st.metric("Сопоставлено образцов", f"{total_matched} ({matching_rate:.1f}%)")
            
            if total_matched > 0:
                st.success(f"✅ Удалось сопоставить {total_matched} образцов!")
                
                # Показываем таблицу сопоставления
                st.subheader("📋 ТАБЛИЦА СОПОСТАВЛЕННЫХ ОБРАЗЦОВ")
                matched_data = []
                for sample in all_matched_samples:
                    matched_data.append({
                        '№ п/п': sample['order'],
                        'Правильное название': sample['correct_name'],
                        'Исходное название': sample['original_name'],
                        'Ключ': sample['key']
                    })
                st.dataframe(pd.DataFrame(matched_data), use_container_width=True)
            
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
                        data=excel_buffer_all,
                        file_name="все_отсортированные_таблицы.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_all",
                        use_container_width=True
                    )
            else:
                st.warning("""
                ⚠️ **Не удалось автоматически сопоставить образцы**
                
                **Что дальше?**
                1. Посмотрите на таблицу сравнения выше
                2. Проанализируйте ключи сопоставления
                3. Мы улучшим алгоритм на основе этих данных
                
                **Пожалуйста, сообщите:**
                - Какие образцы должны быть сопоставлены?
                - По каким признакам они должны сопоставляться (цифры, буквы, комбинации)?
                """)
                
        except Exception as e:
            st.error(f"❌ Ошибка: {str(e)}")
            import traceback
            st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
