import streamlit as st
import pandas as pd
import re
import io
import base64
from typing import List, Dict, Tuple

# =============================================================================
# ФУНКЦИИ ОБРАБОТКИ ДАННЫХ (ранее в utils/processing.py)
# =============================================================================

def parse_correct_order(file_content: str) -> List[Dict]:
    """
    Парсит правильный порядок образцов из файла с вырезками
    """
    lines = file_content.split('\n')
    correct_samples = []
    
    for line in lines:
        # Ищем строки с номерами и названиями образцов
        match = re.match(r'^\s*(\d+)\s+([^\d].*)$', line.strip())
        if match:
            sample_number = int(match.group(1))
            sample_name = match.group(2).strip()
            
            # Убираем разметку типа [ ]{.mark}
            sample_name = re.sub(r'\[(.*?)\]\{\.mark\}', r'\1', sample_name)
            
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
    # Нормализуем название
    normalized = re.sub(r'\s+', ' ', sample_name.strip()).lower()
    
    # Извлекаем ключевые компоненты для сопоставления
    patterns = [
        r'([а-я]+)\s*([а-я]+)?\s*\((\d+[-\d]*),\s*([а-я])\)',  # КПП ВД(50,А)
        r'([а-я]+)\s*([а-я]+)?\s*\((\d+)\)',  # Резервные паттерны
        r'(\d+)[,_]\s*([а-я]+)',  # Для формата "28_КПП ВД"
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

def parse_chemical_tables(file_content: str) -> Dict[str, List[Dict]]:
    """
    Парсит все таблицы химического анализа из файла
    Возвращает словарь: {марка_стали: [список_образцов]}
    """
    lines = file_content.split('\n')
    tables = {}
    current_steel_grade = None
    current_table = []
    in_table = False
    header_found = False
    
    for line in lines:
        # Ищем начало новой марки стали
        steel_match = re.search(r'Марка стали:\s*([^\n]+)', line)
        if steel_match:
            # Сохраняем предыдущую таблицу, если есть
            if current_steel_grade and current_table:
                tables[current_steel_grade] = current_table
                current_table = []
            
            current_steel_grade = steel_match.group(1).strip()
            in_table = False
            header_found = False
            continue
        
        # Ищем начало таблицы (строка с разделителями)
        if re.match(r'^-+\s+-+', line) and current_steel_grade:
            if not in_table:
                in_table = True
            elif header_found:
                # Конец таблицы
                if current_table:
                    tables[current_steel_grade] = current_table
                    current_table = []
                in_table = False
                header_found = False
            continue
        
        # Парсим строки с данными образцов
        if in_table and current_steel_grade:
            # Пропускаем строку с требованиями ТУ
            if 'Требования ТУ' in line or '14-3Р-55-2001' in line:
                continue
            
            # Ищем строки с образцами (содержат цифры, запятые и названия)
            if re.match(r'^\s*\d+\s+[а-я]', line.lower()):
                parts = re.split(r'\s{2,}', line.strip())
                if len(parts) >= 3:  # Как минимум номер, название и одно измерение
                    sample_data = {
                        'original_name': parts[1],
                        'measurements': parts[2:],  # Все измерения
                        'key': create_sample_key(parts[1])
                    }
                    current_table.append(sample_data)
                    header_found = True
    
    # Добавляем последнюю таблицу
    if current_steel_grade and current_table:
        tables[current_steel_grade] = current_table
    
    return tables

def match_and_sort_samples(original_samples: List[Dict], correct_samples: List[Dict]) -> List[Dict]:
    """
    Сопоставляет и сортирует образцы по правильному порядку
    """
    # Создаем маппинг ключей -> правильные названия и порядок
    key_to_correct = {}
    for correct in correct_samples:
        key_to_correct[correct['key']] = {
            'correct_name': correct['correct_name'],
            'order': correct['order']
        }
    
    # Сопоставляем образцы
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
    
    # Сортируем по порядку из правильного списка
    matched_samples.sort(key=lambda x: x['order'])
    
    return matched_samples

def create_final_tables(sorted_samples_dict: Dict[str, List[Dict]]) -> Dict[str, pd.DataFrame]:
    """
    Создает финальные таблицы для каждой марки стали
    """
    final_tables = {}
    
    for steel_grade, samples in sorted_samples_dict.items():
        if not samples:
            continue
            
        # Определяем количество столбцов измерений
        num_measurements = len(samples[0]['measurements'])
        columns = ['№', 'Образец'] + [f'Измерение {i+1}' for i in range(num_measurements)]
        
        data = []
        for i, sample in enumerate(samples):
            row = [i+1, sample['correct_name']] + sample['measurements']
            data.append(row)
        
        final_tables[steel_grade] = pd.DataFrame(data, columns=columns)
    
    return final_tables

def process_multiple_files(correct_order_content: str, chemical_analysis_content: str) -> Dict[str, pd.DataFrame]:
    """
    Основная функция для обработки данных из нескольких файлов
    """
    # Парсим правильный порядок
    correct_samples = parse_correct_order(correct_order_content)
    
    # Парсим таблицы химического анализа
    chemical_tables = parse_chemical_tables(chemical_analysis_content)
    
    # Обрабатываем каждую таблицу отдельно
    sorted_samples_dict = {}
    
    for steel_grade, samples in chemical_tables.items():
        sorted_samples = match_and_sort_samples(samples, correct_samples)
        sorted_samples_dict[steel_grade] = sorted_samples
    
    # Создаем финальные таблицы
    final_tables = create_final_tables(sorted_samples_dict)
    
    return final_tables, correct_samples, chemical_tables

def get_matching_stats(correct_samples: List[Dict], chemical_tables: Dict[str, List[Dict]], final_tables: Dict[str, pd.DataFrame]) -> Dict:
    """
    Возвращает статистику по сопоставлению образцов
    """
    total_correct = len(correct_samples)
    
    total_chemical = 0
    for samples in chemical_tables.values():
        total_chemical += len(samples)
    
    total_matched = 0
    for table in final_tables.values():
        total_matched += len(table)
    
    return {
        'total_correct_samples': total_correct,
        'total_chemical_samples': total_chemical,
        'total_matched_samples': total_matched,
        'matching_rate': round((total_matched / total_chemical) * 100, 2) if total_chemical > 0 else 0
    }

# =============================================================================
# STREAMLIT ИНТЕРФЕЙС (ранее основной app.py)
# =============================================================================

# Настройка страницы
st.set_page_config(
    page_title="Сортировка химического анализа",
    page_icon="🔬",
    layout="wide"
)

# Стили CSS для улучшенного отображения
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .stats-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 0.5rem 0;
    }
    .success-text {
        color: #28a745;
        font-weight: bold;
    }
    .warning-text {
        color: #ffc107;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

def main():
    st.markdown('<div class="main-header">🔬 Автоматическая сортировка химического анализа</div>', unsafe_allow_html=True)
    
    # Описание приложения
    with st.expander("📖 Инструкция по использованию"):
        st.markdown("""
        **Как использовать:**
        1. Загрузите файл с правильным порядком образцов (формат DOCX или TXT)
        2. Загрузите файл с результатами химического анализа (формат DOCX или TXT)  
        3. Нажмите кнопку "Обработать данные"
        4. Просмотрите результаты и скачайте отсортированные таблицы

        **Поддерживаемые форматы:**
        - Текстовые файлы (.txt)
        - Документы Word (.docx)
        - Файлы должны содержать таблицы в текстовом представлении

        **Примеры названий для сопоставления:**
        - "КПП ВД 2, труба 28" → "КПП ВД(28,Г)"
        - "НГ 28_КПП ВД" → "КПП ВД(28,Г)" 
        - "КПП ВД 2, труба 122" → "КПП ВД(50,А)"
        """)
    
    # Загрузка файлов
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📋 Файл с правильным порядком")
        correct_order_file = st.file_uploader(
            "Загрузите файл с правильным порядком образцов",
            type=['txt', 'docx'],
            key="correct_order"
        )
        
        if correct_order_file:
            st.success(f"✅ Файл загружен: {correct_order_file.name}")
            
            # Предпросмотр содержимого
            if st.checkbox("Показать содержимое файла с порядком"):
                try:
                    content = correct_order_file.getvalue().decode("utf-8")
                    st.text_area("Содержимое файла:", content, height=200)
                except:
                    st.warning("Не удалось прочитать файл как текст. Возможно, это бинарный DOCX.")

    with col2:
        st.subheader("🧪 Файл с химическим анализом")
        chemical_analysis_file = st.file_uploader(
            "Загрузите файл с результатами химического анализа",
            type=['txt', 'docx'],
            key="chemical_analysis"
        )
        
        if chemical_analysis_file:
            st.success(f"✅ Файл загружен: {chemical_analysis_file.name}")
            
            # Предпросмотр содержимого
            if st.checkbox("Показать содержимое файла с анализом"):
                try:
                    content = chemical_analysis_file.getvalue().decode("utf-8")
                    st.text_area("Содержимое файла:", content, height=200)
                except:
                    st.warning("Не удалось прочитать файл как текст. Возможно, это бинарный DOCX.")
    
    # Кнопка обработки
    if st.button("🚀 Обработать данные", type="primary"):
        if not correct_order_file or not chemical_analysis_file:
            st.error("❌ Пожалуйста, загрузите оба файла")
            return
        
        try:
            # Чтение файлов
            correct_order_content = correct_order_file.getvalue().decode("utf-8")
            chemical_analysis_content = chemical_analysis_file.getvalue().decode("utf-8")
            
            # Обработка данных
            with st.spinner("🔍 Обрабатываю данные..."):
                final_tables, correct_samples, chemical_tables = process_multiple_files(
                    correct_order_content, 
                    chemical_analysis_content
                )
            
            # Статистика
            stats = get_matching_stats(correct_samples, chemical_tables, final_tables)
            
            # Отображение статистики
            st.subheader("📊 Статистика обработки")
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown(f'<div class="stats-card">Образцов в правильном порядке: <span class="success-text">{stats["total_correct_samples"]}</span></div>', unsafe_allow_html=True)
            with col2:
                st.markdown(f'<div class="stats-card">Образцов в анализе: <span class="success-text">{stats["total_chemical_samples"]}</span></div>', unsafe_allow_html=True)
            with col3:
                st.markdown(f'<div class="stats-card">Сопоставлено образцов: <span class="success-text">{stats["total_matched_samples"]}</span></div>', unsafe_allow_html=True)
            with col4:
                color_class = "success-text" if stats["matching_rate"] > 80 else "warning-text"
                st.markdown(f'<div class="stats-card">Процент сопоставления: <span class="{color_class}">{stats["matching_rate"]}%</span></div>', unsafe_allow_html=True)
            
            # Отображение результатов
            st.subheader("📋 Результаты сортировки")
            
            if not final_tables:
                st.warning("❌ Не удалось сопоставить ни одного образца. Проверьте формат названий в файлах.")
            else:
                for steel_grade, table in final_tables.items():
                    with st.expander(f"Марка стали: {steel_grade} ({len(table)} образцов)"):
                        st.dataframe(table, use_container_width=True)
                        
                        # Кнопка скачивания для каждой таблицы
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            table.to_excel(writer, sheet_name=steel_grade[:30], index=False)
                        
                        excel_buffer.seek(0)
                        b64 = base64.b64encode(excel_buffer.read()).decode()
                        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="отсортированный_{steel_grade}.xlsx">📥 Скачать таблицу Excel</a>'
                        st.markdown(href, unsafe_allow_html=True)
                
                # Скачивание всех таблиц одним файлом
                st.subheader("💾 Пакетное скачивание")
                excel_buffer_all = io.BytesIO()
                with pd.ExcelWriter(excel_buffer_all, engine='openpyxl') as writer:
                    for steel_grade, table in final_tables.items():
                        table.to_excel(writer, sheet_name=steel_grade[:30], index=False)
                
                excel_buffer_all.seek(0)
                b64_all = base64.b64encode(excel_buffer_all.read()).decode()
                href_all = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_all}" download="все_отсортированные_таблицы.xlsx">📦 Скачать все таблицы (Excel)</a>'
                st.markdown(href_all, unsafe_allow_html=True)
        
        except Exception as e:
            st.error(f"❌ Произошла ошибка при обработке: {str(e)}")
            st.info("💡 Проверьте, что файлы имеют правильный формат и содержат таблицы в текстовом представлении")
    
    # Информация о проекте
    st.markdown("---")
    st.markdown("""
    ### ℹ️ О проекте
    Это веб-приложение автоматически сортирует результаты химического анализа 
    согласно заданному порядку образцов и обновляет названия на правильные.
    
    **Особенности:**
    - Автоматическое сопоставление образцов по идентификаторам
    - Обработка нескольких таблиц в одном файле
    - Поддержка различных форматов названий
    - Экспорт результатов в Excel
    
    **Технологии:** Streamlit, Pandas, Python
    """)

if __name__ == "__main__":
    main()