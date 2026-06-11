import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime
from typing import Dict, List, Any, Optional
import copy
from io import BytesIO

# ---------------------------
# Конфигурация страницы
# ---------------------------
st.set_page_config(
    page_title="Заявка на кружки ГБОУ Школа №654",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------------
# Константы
# ---------------------------
DATA_FILE_PATH = 'data.xlsx'
WEEK_DAYS = ["ПН", "ВТ", "СР", "ЧТ", "ПТ", "СБ"]
WEEK_DAYS_RU = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"]

# ---------------------------
# Загрузка справочников
# ---------------------------
@st.cache_data
def load_reference_data():
    """Загрузка справочных данных из data.xlsx"""
    buildings_list = []
    activity_dict = {}
    
    try:
        if os.path.exists(DATA_FILE_PATH):
            df = pd.read_excel(DATA_FILE_PATH, header=None, dtype=str)
            
            # Столбец A (0) - Направления, Столбец B (1) - Виды деятельности
            if df.shape[1] >= 2:
                for idx, row in df.iterrows():
                    col_a = row.iloc[0]
                    col_b = row.iloc[1]
                    
                    if pd.isna(col_a) or pd.isna(col_b):
                        continue
                    
                    direction = str(col_a).strip()
                    activity = str(col_b).strip()
                    
                    if direction.lower() in ['направление', 'направленность', 'направления', 'direction', '']:
                        continue
                    if activity.lower() in ['вид деятельности', 'деятельность', 'activity', '']:
                        continue
                    
                    if direction and activity:
                        if direction not in activity_dict:
                            activity_dict[direction] = []
                        if activity not in activity_dict[direction]:
                            activity_dict[direction].append(activity)
            
            # Столбец C (2) - Учебные корпуса
            if df.shape[1] >= 3:
                col_c = df.iloc[:, 2].dropna()
                for val in col_c:
                    val_str = str(val).strip()
                    if val_str and val_str.lower() not in ['корпус', 'учебный корпус', 'building', '']:
                        if val_str not in buildings_list:
                            buildings_list.append(val_str)
            
            # Столбец D (3) - Дополнительные направленности
            if df.shape[1] >= 4:
                col_d = df.iloc[:, 3].dropna()
                for val in col_d:
                    val_str = str(val).strip()
                    if val_str and val_str.lower() not in ['направленность', 'направление', 'direction', '']:
                        if val_str not in activity_dict:
                            activity_dict[val_str] = []
            
            buildings_list.sort()
            for k in activity_dict:
                activity_dict[k].sort()
            activity_dict = dict(sorted(activity_dict.items()))
            
            return buildings_list, activity_dict
        else:
            return [], {}
    except Exception:
        return [], {}

# ---------------------------
# Вспомогательные функции
# ---------------------------
def safe_int(value) -> Optional[int]:
    try:
        if pd.isna(value) or str(value).strip() == '':
            return None
        return int(float(value))
    except (ValueError, TypeError):
        return None

def safe_str(value) -> str:
    if pd.isna(value) or value is None:
        return ''
    str_val = str(value).strip()
    if str_val.lower() == 'nan':
        return ''
    return str_val

def format_phone(phone: str) -> str:
    """Форматирование телефона в формат +7(XXX) XXX XX XX"""
    digits = re.sub(r'\D', '', phone)
    
    if len(digits) == 10:
        digits = '7' + digits
    elif len(digits) == 11:
        if digits.startswith('8'):
            digits = '7' + digits[1:]
        elif not digits.startswith('7'):
            digits = '7' + digits[1:]
    
    if len(digits) >= 11:
        return f"+7({digits[1:4]}) {digits[4:7]} {digits[7:9]} {digits[9:11]}"
    elif len(digits) >= 7:
        result = f"+7({digits[1:4]}) {digits[4:7]}"
        if len(digits) >= 9:
            result += f" {digits[7:9]}"
        if len(digits) >= 11:
            result += f" {digits[9:11]}"
        return result
    elif len(digits) >= 4:
        return f"+7({digits[1:4]})"
    elif len(digits) >= 2:
        return f"+7({digits[1:]})"
    elif len(digits) >= 1:
        return "+7("
    return ""

def format_time_auto(time_str: str) -> str:
    """Автоматическое форматирование времени с двоеточием"""
    if not time_str:
        return ""
    
    digits = re.sub(r'\D', '', time_str)
    
    if len(digits) == 0:
        return ""
    elif len(digits) <= 2:
        return digits
    elif len(digits) == 3:
        hours = digits[:2]
        minutes = digits[2] + "0"
        return f"{hours}:{minutes}"
    elif len(digits) >= 4:
        hours = digits[:2]
        minutes = digits[2:4]
        return f"{hours}:{minutes}"
    
    return time_str

def validate_phone(phone: str) -> bool:
    """Проверка формата телефона"""
    if not phone:
        return False
    digits = re.sub(r'\D', '', phone)
    return len(digits) == 11

def validate_email(email: str) -> bool:
    """Проверка email: должен быть вида *@ok654.ru"""
    if not email:
        return False
    pattern = r'^[a-zA-Z0-9._%+-]+@ok654\.ru$'
    return re.match(pattern, email) is not None

def generate_filename(fio: str) -> str:
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    if fio:
        parts = fio.split()
        surname = parts[0] if parts else "Педагог"
        initials = "".join([p[0] + "." for p in parts[1:]]) if len(parts) > 1 else ""
        base = f"{surname}_{initials}"
    else:
        base = "заявка"
    return f"{base}_{now}.xlsx"

def parse_schedule_from_str(time_str: str):
    if not time_str:
        return '', ''
    if ' - ' in time_str:
        parts = time_str.split(' - ')
        return parts[0].strip() if len(parts) > 0 else '', parts[1].strip() if len(parts) > 1 else ''
    return time_str.strip(), ''

# ---------------------------
# Инициализация session_state
# ---------------------------
def init_session_state():
    if 'initialized' not in st.session_state:
        buildings, activities = load_reference_data()
        
        st.session_state.school_buildings = buildings
        st.session_state.activity_types = activities
        
        st.session_state.teacher = {
            'fio': '', 'phone': '', 'email': '', 'consent': True
        }
        st.session_state.clubs = []
        st.session_state.current_section = 'teacher'
        st.session_state.data_changed = False
        st.session_state.initialized = True
        st.session_state.reload_counter = 0
        st.session_state.upload_success = False
        st.session_state.upload_message = ""
        st.session_state.phone_formatted = ""

init_session_state()

# ---------------------------
# Экспорт / Импорт Excel
# ---------------------------
def build_export_dataframe() -> pd.DataFrame:
    """Построение DataFrame для экспорта"""
    rows = []
    teacher = st.session_state.teacher
    for club in st.session_state.clubs:
        schedule = club.get('schedule', {})
        schedule_row = {}
        for day in WEEK_DAYS:
            d = schedule.get(day, {})
            start = d.get('start', '')
            end = d.get('end', '')
            if start and end:
                schedule_row[day] = f"{start} - {end}"
            elif start:
                schedule_row[day] = start
            elif end:
                schedule_row[day] = end
            else:
                schedule_row[day] = ''
        
        row = {
            'ФИО педагога': teacher.get('fio', '') or '',
            'Телефон': teacher.get('phone', '') or '',
            'Email': teacher.get('email', '') or '',
            'Согласие на обработку ПД': 'Да' if teacher.get('consent') else 'Нет',
            'Наименование кружка': club.get('name', '') or '',
            'Тип финансирования': club.get('funding', '') or '',
            'Направленность': club.get('direction', '') or '',
            'Вид деятельности': club.get('activity', '') or '',
            'Уровень программы': club.get('level', '') or '',
            'Гендерный состав': club.get('gender', '') or '',
            'Номер кабинета/зала': club.get('room', '') or '',
            'Классы обучения': ', '.join(club.get('classes', [])) or '',
            'Учебный корпус': '; '.join(club.get('buildings', [])) or '',
            **schedule_row,
            'Количество часов в неделю': club.get('hours') if club.get('hours') is not None else '',
            'Макс. человек в группе': club.get('max_group') if club.get('max_group') is not None else '',
            'Ориентация на конкурсы': club.get('competitions', '') or '',
            'Описание кружка': club.get('description', '') or '',
            'Срок реализации': club.get('duration', '') or '',
            'Количество часов в учебный период': club.get('program_hours') if club.get('program_hours') is not None else ''
        }
        rows.append(row)
    return pd.DataFrame(rows)

def import_data_from_dataframe(df: pd.DataFrame):
    """Импорт данных из DataFrame в session_state"""
    if df.empty:
        return False
    
    try:
        row0 = df.iloc[0]
        consent_str = safe_str(row0.get('Согласие на обработку ПД', '')).lower()
        imported_phone = safe_str(row0.get('Телефон', ''))
        
        consent_value = consent_str not in ['нет', 'no', 'false', '0']
        
        st.session_state.teacher = {
            'fio': safe_str(row0.get('ФИО педагога', '')),
            'phone': imported_phone,
            'email': safe_str(row0.get('Email', '')),
            'consent': consent_value
        }
        st.session_state.phone_formatted = imported_phone
        
        clubs = []
        for _, row in df.iterrows():
            classes_str = safe_str(row.get('Классы обучения', ''))
            classes = [c.strip() for c in classes_str.split(',') if c.strip()] if classes_str else []
            
            buildings_str = safe_str(row.get('Учебный корпус', ''))
            buildings = [b.strip() for b in buildings_str.split(';') if b.strip()] if buildings_str else []
            
            schedule = {}
            for day in WEEK_DAYS:
                start, end = parse_schedule_from_str(safe_str(row.get(day, '')))
                schedule[day] = {'start': start, 'end': end}
            
            club = {
                'name': safe_str(row.get('Наименование кружка', '')),
                'funding': safe_str(row.get('Тип финансирования', '')),
                'direction': safe_str(row.get('Направленность', '')),
                'activity': safe_str(row.get('Вид деятельности', '')),
                'level': safe_str(row.get('Уровень программы', '')),
                'gender': safe_str(row.get('Гендерный состав', '')),
                'room': safe_str(row.get('Номер кабинета/зала', '')),
                'hours': safe_int(row.get('Количество часов в неделю')),
                'classes': classes,
                'buildings': buildings,
                'schedule': schedule,
                'max_group': safe_int(row.get('Макс. человек в группе')),
                'competitions': safe_str(row.get('Ориентация на конкурсы', '')),
                'description': safe_str(row.get('Описание кружка', '')),
                'duration': safe_str(row.get('Срок реализации', '')),
                'program_hours': safe_int(row.get('Количество часов в учебный период'))
            }
            clubs.append(club)
        
        st.session_state.clubs = clubs
        st.session_state.data_changed = False
        st.session_state.current_section = 'teacher'
        st.session_state.reload_counter += 1
        
        return True
    except Exception as e:
        st.error(f"Ошибка при импорте данных: {e}")
        return False

# ---------------------------
# Callback для телефона
# ---------------------------
def on_phone_change():
    """Callback при изменении телефона"""
    phone_key = f"teacher_phone_{st.session_state.reload_counter}"
    if phone_key in st.session_state:
        raw_phone = st.session_state[phone_key]
        if raw_phone:
            formatted = format_phone(raw_phone)
            st.session_state.phone_formatted = formatted

# ---------------------------
# Callback для времени
# ---------------------------
def on_time_change(prefix, day, time_type):
    """Callback при изменении времени"""
    time_key = f"{prefix}_sch_{time_type}_{day}_{st.session_state.reload_counter}"
    if time_key in st.session_state:
        raw_time = st.session_state[time_key]
        if raw_time:
            formatted = format_time_auto(raw_time)
            if formatted != raw_time:
                st.session_state[f"{prefix}_sch_{time_type}_{day}_formatted"] = formatted

# ---------------------------
# Функции для форм
# ---------------------------
def get_club_form_fields(club: dict, prefix: str = "") -> dict:
    """Поля формы кружка"""
    school_buildings = st.session_state.get('school_buildings', [])
    activity_types = st.session_state.get('activity_types', {})
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Наименование детского объединения
        name = st.text_input(
            "Наименование детского объединения:", 
            value=club.get('name', '') or '',
            key=f"{prefix}_name_{st.session_state.reload_counter}",
            help="Не должно содержать слова: ЕГЭ, ОГЭ, экзамен, факультатив\n\nНапример: 'Баскетбол' или 'Весёлая кисточка'"
        )
        
        # Тип финансирования
        funding = st.selectbox(
            "Тип финансирования:", 
            ["", "платно", "бесплатно"],
            index=0 if not club.get('funding') else ["", "платно", "бесплатно"].index(club['funding']),
            key=f"{prefix}_funding_{st.session_state.reload_counter}",
            help="Выберите тип финансирования кружка:\n\n- платно (занятия на платной основе)\n\n- бесплатно (занятия на бюджетной основе)"
        )
        
        # Направленность кружка
        directions_list = [""] + list(activity_types.keys())
        current_direction = club.get('direction', '') or ''
        dir_index = directions_list.index(current_direction) if current_direction in directions_list else 0
        direction = st.selectbox(
            "Направленность кружка:", 
            directions_list, 
            index=dir_index, 
            key=f"{prefix}_direction_{st.session_state.reload_counter}",
            help="Выберите направленность кружка из списка"
        )
        
        # Вид деятельности
        activities = [""]
        if direction in activity_types:
            activities += activity_types[direction]
        current_activity = club.get('activity', '') or ''
        act_index = activities.index(current_activity) if current_activity in activities else 0
        activity = st.selectbox(
            "Вид деятельности:", 
            activities, 
            index=act_index, 
            key=f"{prefix}_activity_{st.session_state.reload_counter}",
            help="Выберите конкретный вид деятельности\n\n(список зависит от выбранной направленности)"
        )
        
        # Уровень программы
        level = st.selectbox(
            "Уровень программы:", 
            ["", "Ознакомительный", "Базовый", "Углубленный", "Вводный"],
            index=0 if not club.get('level') else ["", "Ознакомительный", "Базовый", "Углубленный", "Вводный"].index(club['level']),
            key=f"{prefix}_level_{st.session_state.reload_counter}",
            help="Выберите уровень образовательной программы:\n\n- Ознакомительный (5-18 лет, срок обучения > 3 месяцев, количество часов от 1 до 3 в неделю)\n\n- Базовый (8-18 лет, срок обучения > 1 года, количество часов от 3 до 5 в неделю)\n\n- Углубленный (12-18 лет, срок обучения от 2 лет, количество часов от 4 до 8 в неделю)\n\n- Вводный (5-18 лет, срок обучения > 10 часов, количество часов по Программе) (для каникулярных программ)"
        )
        
        # Гендерный состав
        gender = st.selectbox(
            "Гендерный состав:", 
            ["", "смешанный", "муж", "жен"],
            index=0 if not club.get('gender') else ["", "смешанный", "муж", "жен"].index(club['gender']),
            key=f"{prefix}_gender_{st.session_state.reload_counter}",
            help="Выберите гендерный состав группы:\n\n- смешанный (мальчики и девочки)\n\n- муж (только мальчики)\n\n- жен (только девочки)"
        )
        
        # Номер кабинета/зала
        room = st.text_input(
            "Номер кабинета/зала:", 
            value=club.get('room', '') or '', 
            key=f"{prefix}_room_{st.session_state.reload_counter}",
            help="Укажите номер кабинета или название зала\n\nНапример: 301, Спортивный зал, Актовый зал"
        )
    
    with col2:
        # Количество часов в неделю
        hours = st.number_input(
            "Количество часов в неделю:", 
            min_value=0, step=1, 
            value=club.get('hours') or 0, 
            key=f"{prefix}_hours_{st.session_state.reload_counter}",
            help="Укажите количество академических часов в неделю\n\n(только число, например: 2, 4, 6)"
        )
        
        # Максимальное количество человек в группе
        max_group = st.number_input(
            "Максимальное количество человек в одной группе:", 
            min_value=0, step=1, 
            value=club.get('max_group') or 0, 
            key=f"{prefix}_max_group_{st.session_state.reload_counter}",
            help="Эта цифра не может быть изменена в течение года!!!\n\n(только число, например: 15, 20, 25)"
        )
        
        # Срок реализации
        duration = st.text_input(
            "Срок реализации:", 
            value=club.get('duration', '') or '', 
            key=f"{prefix}_duration_{st.session_state.reload_counter}",
            help="Укажите срок реализации программы\n\nНапример: 6 месяцев, 1 год, 2 года"
        )
        
        # Количество часов в учебный период
        program_hours = st.number_input(
            "Количество часов:", 
            min_value=0, step=1, 
            value=club.get('program_hours') or 0, 
            key=f"{prefix}_program_hours_{st.session_state.reload_counter}",
            help="Общее количество часов за весь период обучения\n\n(только число)"
        )
        
        # Классы обучения
        st.write("Классы обучения:")
        st.caption("Для обучающихся каких классов данное объединение\n\nМожно выбрать несколько")
        cols = st.columns(4)
        classes_selected = club.get('classes', [])
        classes_check = {}
        all_classes = ["1","2","3","4","5","6","7","8","9","10","11","Дошкольники"]
        for i, cl in enumerate(all_classes):
            with cols[i % 4]:
                classes_check[cl] = st.checkbox(cl, value=cl in classes_selected, key=f"{prefix}_class_{cl}_{st.session_state.reload_counter}")
        
        # Учебные корпуса
        st.write("Выберите учебные корпуса:")
        st.caption("Учебный корпус, в котором будут проводиться занятия")
        buildings_selected = club.get('buildings', [])
        buildings_check = {}
        buildings_manual = ""
        if school_buildings:
            bcols = st.columns(2)
            for i, b in enumerate(school_buildings):
                with bcols[i % 2]:
                    buildings_check[b] = st.checkbox(b, value=b in buildings_selected, key=f"{prefix}_build_{b}_{st.session_state.reload_counter}")
        else:
            st.info("Справочник корпусов не загружен. Введите через запятую.")
            buildings_manual = st.text_input("Корпуса", value=", ".join(buildings_selected), key=f"{prefix}_buildings_manual_{st.session_state.reload_counter}")
    
    # Расписание
    st.write("**Расписание занятий:**")
    st.caption("Укажите время в формате ЧЧ:ММ (двоеточие добавляется автоматически)\n\nМежду занятиями должен быть перерыв 10 минут")
    
    schedule_data = club.get('schedule', {})
    new_schedule = {}
    
    # Заголовки таблицы
    header_cols = st.columns([2, 2, 2])
    with header_cols[0]:
        st.write("**День недели**")
    with header_cols[1]:
        st.write("**Время начала**")
    with header_cols[2]:
        st.write("**Время окончания**")
    
    for day, day_ru in zip(WEEK_DAYS, WEEK_DAYS_RU):
        cols = st.columns([2, 2, 2])
        with cols[0]:
            st.write(day_ru)
        with cols[1]:
            formatted_key = f"{prefix}_sch_start_{day}_formatted"
            start_val = st.session_state.get(formatted_key, '') or schedule_data.get(day, {}).get('start', '') or ''
            start = st.text_input(
                f"start_{day}", 
                value=start_val, 
                placeholder="15:00", 
                label_visibility="collapsed", 
                key=f"{prefix}_sch_start_{day}_{st.session_state.reload_counter}",
                on_change=on_time_change,
                args=(prefix, day, 'start')
            )
        with cols[2]:
            formatted_key = f"{prefix}_sch_end_{day}_formatted"
            end_val = st.session_state.get(formatted_key, '') or schedule_data.get(day, {}).get('end', '') or ''
            end = st.text_input(
                f"end_{day}", 
                value=end_val, 
                placeholder="16:30", 
                label_visibility="collapsed", 
                key=f"{prefix}_sch_end_{day}_{st.session_state.reload_counter}",
                on_change=on_time_change,
                args=(prefix, day, 'end')
            )
        new_schedule[day] = {'start': format_time_auto(start) if start else '', 'end': format_time_auto(end) if end else ''}
    
    # Ориентация на конкурсы
    competitions = st.text_area(
        "Участие в мероприятиях:", 
        value=club.get('competitions', '') or '', 
        height=100, max_chars=500, 
        key=f"{prefix}_competitions_{st.session_state.reload_counter}",
        help="На какие конкретные соревнования, конкурсы, фестивали ориентирована работа объединения\n\nГде дети смогут показать результат работы в объединении.\n\nМаксимум 500 символов"
    )
    
    # Описание кружка
    description = st.text_area(
        "Описание кружка:", 
        value=club.get('description', '') or '', 
        height=150, max_chars=1000, 
        key=f"{prefix}_description_{st.session_state.reload_counter}",
        help="Описание выкладывается на портал mos.ru!\n\nДля детей и родителей - подробно, понятное, без грамматических ошибок.\n\nМаксимум 1000 символов"
    )
    
    # Сбор данных
    selected_classes = [k for k, v in classes_check.items() if v]
    if school_buildings:
        selected_buildings = [k for k, v in buildings_check.items() if v]
    else:
        selected_buildings = [b.strip() for b in buildings_manual.split(',') if b.strip()] if buildings_manual else []
    
    return {
        'name': name,
        'funding': funding,
        'direction': direction,
        'activity': activity,
        'level': level,
        'gender': gender,
        'room': room,
        'hours': int(hours) if hours > 0 else None,
        'max_group': int(max_group) if max_group > 0 else None,
        'classes': selected_classes,
        'buildings': selected_buildings,
        'schedule': new_schedule,
        'competitions': competitions,
        'description': description,
        'duration': duration,
        'program_hours': int(program_hours) if program_hours > 0 else None
    }

def validate_club_data(data: dict) -> bool:
    forbidden = ['егэ', 'огэ', 'экзамен', 'факультатив']
    if any(word in data['name'].lower() for word in forbidden):
        st.error("Название кружка не должно содержать слова: ЕГЭ, ОГЭ, экзамен, факультатив")
        return False
    return True

# ---------------------------
# Функции отображения разделов
# ---------------------------
def show_teacher_section():
    """Отображение раздела контактов педагога"""
    st.header("Контакты педагога")
    
    teacher = st.session_state.teacher
    phone_value = st.session_state.get('phone_formatted', '') or teacher.get('phone', '') or ''
    
    col1, col2 = st.columns(2)
    with col1:
        # ФИО
        fio = st.text_input(
            "ФИО:", 
            value=teacher.get('fio', '') or '', 
            key=f"teacher_fio_{st.session_state.reload_counter}",
            help="Введите полное ФИО педагога\n\nНапример: Иванов Иван Иванович"
        )
        
        # Мобильный телефон
        phone = st.text_input(
            "Мобильный телефон:", 
            value=phone_value, 
            key=f"teacher_phone_{st.session_state.reload_counter}",
            help="Введите номер телефона для решения организационных вопросов\n\nНапример: +7 (999) 123-45-67",
            placeholder="+7(___) ___ __ __",
            on_change=on_phone_change
        )
        
        if phone:
            digits_count = len(re.sub(r'\D', '', phone))
            if digits_count > 0 and digits_count < 10:
                st.caption(f"Введено {digits_count}/10 цифр. Введите ещё {10 - digits_count}")
            elif digits_count >= 10:
                st.caption(f"Будет сохранено: **{phone_value}**")
    
    with col2:
        # Email
        email = st.text_input(
            "Email аккаунта @ok654.ru:", 
            value=teacher.get('email', '') or '', 
            key=f"teacher_email_{st.session_state.reload_counter}",
            help="Введите адрес электронной почты\n\nФормат: example@ok654.ru",
            placeholder="example@ok654.ru"
        )
        
        # Согласие на обработку ПД
        consent = st.checkbox(
            "Согласен на обработку персональных данных", 
            value=teacher.get('consent', True), 
            key=f"teacher_consent_{st.session_state.reload_counter}",
            help="Необходимо отметить для сохранения контактов"
        )
    
    # Валидация в реальном времени
    if phone:
        digits_count = len(re.sub(r'\D', '', phone))
        if digits_count > 0 and digits_count < 10:
            st.warning(f"⚠️ Введено {digits_count} цифр. Минимум 10 цифр.")
    
    if email and not email.endswith('@ok654.ru'):
        st.warning("⚠️ Email должен заканчиваться на @ok654.ru")
    
    if st.button("Сохранить", type="primary", key=f"save_teacher_btn_{st.session_state.reload_counter}"):
        phone_to_save = st.session_state.get('phone_formatted', '') or phone
        
        errors = []
        
        if not fio:
            errors.append("ФИО обязательно для заполнения")
        
        if not phone_to_save:
            errors.append("Мобильный телефон обязателен для заполнения")
        elif not validate_phone(phone_to_save):
            digits_count = len(re.sub(r'\D', '', phone_to_save))
            if digits_count < 10:
                errors.append(f"Телефон содержит только {digits_count} цифр. Введите минимум 10 цифр.")
            else:
                errors.append("Неверный формат телефона")
        
        if not email:
            errors.append("Email обязателен для заполнения")
        elif not email.endswith('@ok654.ru'):
            errors.append("Email должен заканчиваться на @ok654.ru")
        elif not validate_email(email):
            errors.append("Неверный формат Email. Используйте: example@ok654.ru")
        
        if not consent:
            errors.append("Для сохранения контактов необходимо согласие на обработку персональных данных")
        
        if errors:
            for error in errors:
                st.error(error)
        else:
            st.session_state.teacher = {
                'fio': fio,
                'phone': phone_to_save,
                'email': email,
                'consent': consent
            }
            st.session_state.phone_formatted = phone_to_save
            st.session_state.data_changed = True
            st.success("Контакты сохранены!")
            st.rerun()

def show_club_section(club_index: int):
    """Отображение раздела редактирования кружка"""
    if club_index >= len(st.session_state.clubs):
        st.warning("Кружок не найден")
        st.session_state.current_section = 'teacher'
        st.rerun()
        return
    
    club = st.session_state.clubs[club_index]
    st.header(f"Редактирование кружка: {club.get('name', 'Без названия')}")
    
    club_data = get_club_form_fields(club, prefix=f"edit_club_{club_index}")
    
    col1, col2, col3, col4 = st.columns([1, 1, 1, 3])
    with col1:
        if st.button("Сохранить", type="primary", key=f"save_club_{club_index}_{st.session_state.reload_counter}"):
            if validate_club_data(club_data):
                st.session_state.clubs[club_index] = club_data
                st.session_state.data_changed = True
                st.success("Изменения сохранены!")
                st.rerun()
    
    with col2:
        if st.button("Копировать", key=f"copy_club_{club_index}_{st.session_state.reload_counter}"):
            new_club = copy.deepcopy(st.session_state.clubs[club_index])
            new_club['name'] = new_club.get('name', '') + ' (копия)'
            st.session_state.clubs.append(new_club)
            st.session_state.data_changed = True
            new_index = len(st.session_state.clubs) - 1
            st.session_state.current_section = f'club_{new_index}'
            st.success("Кружок скопирован!")
            st.rerun()
    
    with col3:
        if st.button("Удалить", key=f"delete_club_{club_index}_{st.session_state.reload_counter}"):
            club_name = st.session_state.clubs[club_index].get('name', 'Без названия')
            st.session_state.clubs.pop(club_index)
            st.session_state.data_changed = True
            st.session_state.current_section = 'teacher'
            st.success(f"Кружок '{club_name}' удалён!")
            st.rerun()

def show_new_club_section():
    """Отображение раздела создания нового кружка"""
    st.header("Данные кружка")
    
    club_data = get_club_form_fields({}, prefix="new_club")
    
    if st.button("Сохранить", type="primary", key=f"create_club_btn_{st.session_state.reload_counter}"):
        if validate_club_data(club_data):
            st.session_state.clubs.append(club_data)
            st.session_state.data_changed = True
            new_index = len(st.session_state.clubs) - 1
            st.session_state.current_section = f'club_{new_index}'
            st.success("Кружок создан!")
            st.rerun()

# ---------------------------
# Боковая панель
# ---------------------------
with st.sidebar:
    st.title("Заявка на кружки")
    st.caption("ГБОУ ШКОЛА №654 ИМЕНИ А.Д. ФРИДМАНА")
    st.markdown("---")
    
    st.subheader("Работа с файлами:")
    
    def handle_file_upload():
        if st.session_state.get('file_uploader') is not None:
            try:
                df = pd.read_excel(st.session_state.file_uploader)
                if import_data_from_dataframe(df):
                    st.session_state.upload_success = True
                    st.session_state.upload_message = f"Загружено кружков: {len(st.session_state.clubs)}"
                else:
                    st.session_state.upload_success = False
                    st.session_state.upload_message = "Ошибка при импорте данных"
            except Exception as e:
                st.session_state.upload_success = False
                st.session_state.upload_message = f"Ошибка: {e}"
    
    st.file_uploader("Загрузить Excel", type=["xlsx", "xls"], 
                     key="file_uploader",
                     on_change=handle_file_upload)
    
    if st.session_state.get('upload_success'):
        st.success(st.session_state.upload_message)
        if st.button("Обновить интерфейс", key="refresh_btn", use_container_width=True):
            st.session_state.reload_counter += 1
            st.rerun()
    elif st.session_state.get('upload_message'):
        st.error(st.session_state.upload_message)
    
    if st.session_state.clubs:
        if st.session_state.data_changed:
            df_export = build_export_dataframe()
            df_export = df_export.fillna('')
            filename = generate_filename(st.session_state.teacher.get('fio', ''))
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Заявка на кружки')
            output.seek(0)
            
            st.download_button(
                label="Скачать Excel",
                data=output,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
                key=f"download_btn_{st.session_state.reload_counter}"
            )
            st.caption(f"Файл: {filename}")
        else:
            st.info("Нет новых изменений")
    else:
        st.info("Добавьте кружки для скачивания")
    
    st.subheader("Навигация")
    
    if st.button("Контакты педагога", use_container_width=True, 
                type="primary" if st.session_state.current_section == 'teacher' else "secondary",
                key="nav_teacher"):
        st.session_state.current_section = 'teacher'
        st.rerun()
    
    if st.session_state.clubs:
        st.markdown("**Кружки:**")
        for i, club in enumerate(st.session_state.clubs):
            name = club.get('name', 'Без названия')[:30]
            
            fields = [club.get('name'), club.get('funding'), club.get('direction'),
                     club.get('activity'), club.get('level'), club.get('gender'),
                     club.get('room'), club.get('hours'), club.get('classes'),
                     club.get('buildings'), club.get('max_group'),
                     club.get('competitions'), club.get('description'),
                     club.get('duration'), club.get('program_hours')]
            has_schedule = any(
                club.get('schedule', {}).get(d, {}).get('start') or 
                club.get('schedule', {}).get(d, {}).get('end') 
                for d in WEEK_DAYS
            )
            fields.append(has_schedule)
            filled = sum(1 for f in fields if f)
            pct = int(filled / len(fields) * 100) if fields else 0
            
            emoji = "🔴" if pct < 30 else ("🟡" if pct < 70 else "🟢")
            
            section_key = f'club_{i}'
            if st.button(f"{emoji} {i+1}. {name}", use_container_width=True,
                       type="primary" if st.session_state.current_section == section_key else "secondary",
                       key=f"nav_club_{i}"):
                st.session_state.current_section = section_key
                st.rerun()
    

    
    if st.button("Добавить кружок", use_container_width=True, type="primary", key="nav_add_club"):
        st.session_state.current_section = 'new_club'
        st.rerun()
    st.markdown("---")
# ---------------------------
# Основная область
# ---------------------------
if st.session_state.current_section == 'teacher':
    show_teacher_section()
elif st.session_state.current_section == 'new_club':
    show_new_club_section()
elif st.session_state.current_section.startswith('club_'):
    try:
        club_index = int(st.session_state.current_section.split('_')[1])
        show_club_section(club_index)
    except (ValueError, IndexError):
        st.session_state.current_section = 'teacher'
        st.rerun()
else:
    show_teacher_section()