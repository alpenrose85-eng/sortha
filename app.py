import streamlit as st
import pandas as pd
import re
import io
from typing import List, Dict, Tuple, Optional
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
            
            # Парсим структурированную информацию из названия
            parsed_info = parse_structured_name(sample_name)
            
            correct_samples.append({
                'order': sample_number,
                'correct_name': sample_name,
                'parsed': parsed_info
            })
    
    return correct_samples

def parse_structured_name(name: str) -> Dict:
    """
    Парсит структурированную информацию из названия образца
    Возвращает: {surface, number, letter, original}
    """
    # Нормализуем название
    normalized = name.strip()
    
    # Паттерны для структурированных названий (правильный порядок)
    patterns_correct = [
        # ЭПК(1,А), КПП ВД(50,А), КПП НД-1(19,А)
        r'([а-яА-Я\s-]+)\((\d+),\s*([а-яА-Я])\)',
    ]
    
    for pattern in patterns_correct:
        match = re.search(pattern, normalized)
        if match:
            surface = match.group(1).strip()
            number = match.group(2)
            letter = match.group(3).upper()
            
            return {
                'surface': surface,
                'number': number,
                'letter': letter,
                'original': normalized,
                'type': 'structured'
            }
    
    # Паттерны для названий из химического анализа
    patterns_analysis = [
        # НГ ШПП 4, НБ ШПП 6
        r'([а-яА-Я]{2})\s+([а-яА-Я]+)\s+(\d+)',
        # КПП ВД 2, труба 13
        r'([а-яА-Я\s-]+)\s+(\d+)[,\s]+\S*\s*(\d+)',
        # КПП ВД 2, труба 13 (упрощенный)
        r'([а-яА-Я\s-]+?)\s+\d+[,\s]+\S*\s*(\d+)',
        # Просто число в конце
        r'(.+?)\s+(\d+)$',
    ]
    
    for pattern in patterns_analysis:
        match = re.search(pattern, normalized)
        if match:
            surface = match.group(1).strip()
            number = match.group(2) if len(match.groups()) >= 2 else None
            letter = None  # В химическом анализе буква обычно не указана
            
            return {
                'surface': surface,
                'number': number,
                'letter': letter,
                'original': normalized,
                'type': 'analysis'
            }
    
    # Если не удалось распарсить
    return {
        'surface': normalized,
        'number': None,
        'letter': None,
        'original': normalized,
        'type': 'unknown'
    }

def normalize_surface_name(surface: str) -> str:
    """
    Нормализует название поверхности нагрева для сопоставления
    """
    # Приводим к нижнему регистру и убираем лишние пробелы
    normalized = re.sub(r'\s+', ' ', surface.strip().lower())
    
    # Стандартизируем названия поверхностей
    surface_mappings = {
        r'кпп\s*вд': 'КПП ВД',
        r'кпп\s*нд-1': 'КПП НД-1',
        r'кпп\s*нд-2': 'КПП НД-2',
        r'кпп\s*нд': 'КПП НД',
        r'шпп': 'ШПП',
        r'эпк': 'ЭПК',
        r'пс\s*кш': 'ПС КШ',
    }
    
    for pattern, standard in surface_mappings.items():
        if re.search(pattern, normalized):
            return standard
    
    # Если не нашли стандартное название, возвращаем оригинал в нормализованном виде
    return normalized.title()

def parse_chemical_tables_improved(file_content: str) -> Dict[str, List[Dict]]:
    """
    Улучшенный парсинг химического анализа
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
        
        # Ищем строки с образцами
        if current_steel_grade and re.match(r'^\s*\d+\s+[^\d]', line):
            # Разделяем по множественным пробелам
            parts = re.split(r'\s{2,}', line)
            if len(parts) >= 2:
                sample_name = parts[1]
                measurements = parts[2:] if len(parts) > 2 else []
                
                # Парсим структурированную информацию
                parsed_info = parse_structured_name(sample_name)
                
                sample_data = {
                    'original_name': sample_name,
                    'measurements': measurements,
                    'parsed': parsed_info
                }
                current_table.append(sample_data)
    
    # Добавляем последнюю таблицу
    if current_steel_grade and current_table:
        tables[current_steel_grade] = current_table
    
    return tables

def find_best_match(analysis_sample: Dict, correct_samples: List[Dict]) -> Optional[Dict]:
    """
    Находит наилучшее соответствие для образца из анализа среди правильных образцов
    """
    analysis_parsed = analysis_sample['parsed']
    analysis_surface = normalize_surface_name(analysis_parsed['surface'])
    analysis_number = analysis_parsed['number']
    
    best_match = None
    best_score = 0
    
    for correct_sample in correct_samples:
        correct_parsed = correct_sample['parsed']
        correct_surface = normalize_surface_name(correct_parsed['surface'])
        correct_number = correct_parsed['number']
        correct_letter = correct_parsed['letter']
        
        score = 0
        
        # 1. Совпадение поверхности нагрева (самый важный критерий)
        if analysis_surface == correct_surface:
            score += 100
        elif analysis_surface in correct_surface or correct_surface in analysis_surface:
            score += 50  # Частичное совпадение
        
        # 2. Совпадение номера трубы
        if analysis_number and correct_number and analysis_number == correct_number:
            score += 50
        
        # 3. Если есть буква в анализе, проверяем совпадение
        if analysis_parsed['letter'] and correct_letter and analysis_parsed['letter'] == correct_letter:
            score += 25
        
        # 4. Дополнительные критерии для частичных совпадений
        if analysis_number and correct_number and analysis_number in correct_number:
            score += 10
        if correct_number and analysis_number and correct_number in analysis_number:
            score += 10
        
        # Обновляем лучший матч
        if score > best_score:
            best_score = score
            best_match = correct_sample
    
    # Возвращаем матч только если score достаточно высок
    return best_match if best_score >= 50 else None

def match_samples_improved(analysis_samples: List[Dict], correct_samples: List[Dict]) -> List[Dict]:
    """
    Улучшенное сопоставление образцов
    """
    matched = []
    used_correct_indices = set()
    
    # Первый проход: точные совпадения
    for i, analysis_sample in enumerate(analysis_samples):
        best_match = find_best_match(analysis_sample, correct_samples)
        
        if best_match and best_match['order'] not in used_correct_indices:
            matched.append({
                'correct_name': best_match['correct_name'],
                'original_name': analysis_sample['original_name'],
                'measurements': analysis_sample['measurements'],
                'order': best_match['order'],
                'match_quality': 'exact',
                'analysis_surface': analysis_sample['parsed']['surface'],
                'correct_surface': best_match['parsed']['surface'],
                'analysis_number': analysis_sample['parsed']['number'],
                'correct_number': best_match['parsed']['number']
            })
            used_correct_indices.add(best_match['order'])
    
    # Второй проход: частичные совпадения (для неподобранных образцов)
    for i, analysis_sample in enumerate(analysis_samples):
        # Пропускаем уже сопоставленные
        if any(m['original_name'] == analysis_sample['original_name'] for m in matched):
            continue
        
        # Ищем частичное совпадение
        analysis_parsed = analysis_sample['parsed']
        analysis_surface = normalize_surface_name(analysis_parsed['surface'])
        analysis_number = analysis_parsed['number']
        
        for correct_sample in correct_samples:
            if correct_sample['order'] in used_correct_indices:
                continue
                
            correct_parsed = correct_sample['parsed']
            correct_surface = normalize_surface_name(correct_parsed['surface'])
            correct_number = correct_parsed['number']
            
            # Частичное совпадение по поверхности
            surface_match = (analysis_surface in correct_surface or 
                           correct_surface in analysis_surface or
                           any(word in correct_surface for word in analysis_surface.split()) or
                           any(word in analysis_surface for word in correct_surface.split()))
            
            # Частичное совпадение по номеру
            number_match = (analysis_number and correct_number and 
                          (analysis_number in correct_number or correct_number in analysis_number))
            
            if surface_match and number_match:
                matched.append({
                    'correct_name': correct_sample['correct_name'],
                    'original_name': analysis_sample['original_name'],
                    'measurements': analysis_sample['measurements'],
                    'order': correct_sample['order'],
                    'match_quality': 'partial',
                    'analysis_surface': analysis_sample['parsed']['surface'],
                    'correct_surface': correct_sample['parsed']['surface'],
                    'analysis_number': analysis_sample['parsed']['number'],
                    'correct_number': correct_sample['parsed']['number']
                })
                used_correct_indices.add(correct_sample['order'])
                break
    
    # Сортируем по порядку из правильного списка
    matched.sort(key=lambda x: x['order'])
    
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
    .match-exact { background-color: #d4edda !important; }
    .match-partial { background-color: #fff3cd !important; }
</style>
""", unsafe_allow_html=True)

def main():
    st.markdown('<div class="main-header">🔬 Умная сортировка химического анализа</div>', unsafe_allow_html=True)
    
    with st.expander("📖 ИНСТРУКЦИЯ", expanded=False):
        st.markdown("""
        **Как работает программа:**
        1. **Анализ структуры названий**: Программа определяет шифр поверхности нагрева, номер трубы и корпус/нитку
        2. **Сопоставление по приоритетам**: 
           - Сначала точные совпадения по поверхности нагрева и номеру
           - Затем частичные совпадения
        3. **Визуализация результатов**: Показывает качество сопоставления
        
        **Структура названий:**
        - Шифры поверхностей: ЭПК, ПС КШ, КПП НД-1, КПП НД-2, КПП ВД, ШПП
        - Номер трубы: число
        - Корпус/нитка: А, Б, В, Г
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
            with st.spinner("🔍 Анализирую структуру названий..."):
                correct_samples = parse_correct_order(correct_order_content)
                chemical_tables = parse_chemical_tables_improved(chemical_analysis_content)
            
            if not correct_samples:
                st.error("❌ Не найдены образцы в файле порядка")
                return
                
            if not chemical_tables:
                st.error("❌ Не найдены таблицы анализа")
                return
            
            # Показываем информацию о распарсенных названиях
            st.subheader("🔍 СТРУКТУРИРОВАННЫЙ АНАЛИЗ НАЗВАНИЙ")
            
            # Правильные образцы
            st.write("**Правильные образцы (структурированные):**")
            correct_data = []
            for sample in correct_samples:
                parsed = sample['parsed']
                correct_data.append({
                    '№': sample['order'],
                    'Исходное название': sample['correct_name'],
                    'Поверхность': parsed['surface'],
                    'Номер': parsed['number'],
                    'Корпус': parsed['letter']
                })
            st.dataframe(pd.DataFrame(correct_data), use_container_width=True)
            
            # Образцы из анализа
            all_analysis_samples = []
            for steel_grade, samples in chemical_tables.items():
                for sample in samples:
                    all_analysis_samples.append(sample)
            
            st.write("**Образцы из химического анализа:**")
            analysis_data = []
            for i, sample in enumerate(all_analysis_samples):
                parsed = sample['parsed']
                analysis_data.append({
                    '№': i + 1,
                    'Исходное название': sample['original_name'],
                    'Поверхность': parsed['surface'],
                    'Номер': parsed['number'],
                    'Корпус': parsed['letter'],
                    'Тип': parsed['type']
                })
            st.dataframe(pd.DataFrame(analysis_data), use_container_width=True)
            
            # Сопоставление
            with st.spinner("🔄 Сопоставляю образцы..."):
                all_matched_samples = []
                final_tables = {}
                
                for steel_grade, samples in chemical_tables.items():
                    matched_samples = match_samples_improved(samples, correct_samples)
                    all_matched_samples.extend(matched_samples)
                    
                    if matched_samples:
                        # Создаем DataFrame для этой марки стали
                        if matched_samples and matched_samples[0]['measurements']:
                            num_measurements = len(matched_samples[0]['measurements'])
                            columns = ['№', 'Образец'] + [f'Измерение {i+1}' for i in range(num_measurements)]
                            
                            data = []
                            for i, sample in enumerate(matched_samples):
                                row = [i + 1, sample['correct_name']] + sample['measurements']
                                data.append(row)
                            
                            final_tables[steel_grade] = pd.DataFrame(data, columns=columns)
            
            # Статистика и результаты сопоставления
            total_correct = len(correct_samples)
            total_analysis = sum(len(samples) for samples in chemical_tables.values())
            total_matched = len(all_matched_samples)
            matching_rate = (total_matched / total_analysis) * 100 if total_analysis > 0 else 0
            
            st.subheader("📊 РЕЗУЛЬТАТЫ СОПОСТАВЛЕНИЯ")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Образцов в порядке", total_correct)
            with col2:
                st.metric("Образцов в анализе", total_analysis)
            with col3:
                st.metric("Сопоставлено", f"{total_matched} ({matching_rate:.1f}%)")
            
            # Детальная таблица сопоставления
            if all_matched_samples:
                st.subheader("🔍 ДЕТАЛЬНАЯ ТАБЛИЦА СОПОСТАВЛЕНИЯ")
                
                match_data = []
                for sample in all_matched_samples:
                    match_data.append({
                        '№ п/п': sample['order'],
                        'Правильное название': sample['correct_name'],
                        'Исходное название': sample['original_name'],
                        'Качество': sample['match_quality'],
                        'Поверхность (анализ)': sample['analysis_surface'],
                        'Поверхность (правильно)': sample['correct_surface'],
                        'Номер (анализ)': sample['analysis_number'],
                        'Номер (правильно)': sample['correct_number']
                    })
                
                match_df = pd.DataFrame(match_data)
                
                # Применяем стили для визуализации качества сопоставления
                def color_match_quality(val):
                    if val == 'exact':
                        return 'background-color: #d4edda'
                    elif val == 'partial':
                        return 'background-color: #fff3cd'
                    return ''
                
                styled_df = match_df.style.applymap(color_match_quality, subset=['Качество'])
                st.dataframe(styled_df, use_container_width=True)
                
                # Группируем по качеству сопоставления
                exact_matches = [m for m in all_matched_samples if m['match_quality'] == 'exact']
                partial_matches = [m for m in all_matched_samples if m['match_quality'] == 'partial']
                
                st.info(f"""
                **Качество сопоставления:**
                - ✅ Точные совпадения: {len(exact_matches)}
                - ⚠️ Частичные совпадения: {len(partial_matches)}
                """)
            
            # Отсортированные результаты
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
                ⚠️ **Не удалось сопоставить образцы**
                
                **Возможные причины:**
                - Названия в файлах сильно отличаются
                - Не распознается структура названий
                - Нужна дополнительная настройка алгоритма
                """)
                
        except Exception as e:
            st.error(f"❌ Ошибка: {str(e)}")
            import traceback
            st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
