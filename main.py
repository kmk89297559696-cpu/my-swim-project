import sqlite3
import pandas as pd
import os
import re

# ============= НАСТРОЙКИ =============
DB_PATH = "competition.db"
INPUT_EXCEL = "data/competition_data.xlsx"
OUTPUT_EXCEL = "output/protocol_sorevnovaniy.xlsx"

os.makedirs("data", exist_ok=True)
os.makedirs("output", exist_ok=True)

# Система начисления очков
POINTS = {1: 10, 2: 8, 3: 6, 4: 5, 5: 4, 6: 3, 7: 2, 8: 1}


# ============= КОНВЕРТАЦИЯ ВРЕМЕНИ =============

def parse_time_to_seconds(time_str):
    """Преобразует формат ММ,СС,МС в секунды"""
    if pd.isna(time_str):
        return None
    time_str = str(time_str).strip().replace(' ', '')
    # Паттерн: минуты,секунды,сотые
    match = re.search(r'(\d+)[,\.](\d+)[,\.](\d+)', time_str)
    if match:
        minutes = int(match.group(1))
        seconds = int(match.group(2))
        centiseconds = int(match.group(3))
        return minutes * 60 + seconds + centiseconds / 100
    # Альтернативный формат: секунды,сотые
    match2 = re.search(r'(\d+)[,\.](\d+)', time_str)
    if match2:
        seconds = int(match2.group(1))
        centiseconds = int(match2.group(2))
        return seconds + centiseconds / 100
    try:
        return float(time_str)
    except:
        print(f"⚠️ Не могу распознать время: {time_str}")
        return None


def format_time_from_seconds(seconds):
    """Преобразует секунды в формат ММ.СС.сс (минуты.секунды.сотые)"""
    if seconds is None or pd.isna(seconds):
        return "—"
    minutes = int(seconds // 60)
    secs = int(seconds % 60)
    centiseconds = int(round((seconds - int(seconds)) * 100))
    # Округление до сотых
    if centiseconds >= 100:
        centiseconds = 99
    return f"{minutes}.{secs:02d}.{centiseconds:02d}"


def normalize_gender(gender):
    if pd.isna(gender):
        return None
    g = str(gender).strip().upper()
    if g in ['M', 'М', 'MALE', 'МУЖ', 'МУЖСКОЙ']:
        return 'M'
    if g in ['F', 'Ж', 'FEMALE', 'ЖЕН', 'ЖЕНСКИЙ']:
        return 'F'
    return None


# ============= БАЗА ДАННЫХ =============

def get_connection():
    return sqlite3.connect(DB_PATH)


def create_database():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DROP TABLE IF EXISTS results")
    cursor.execute("DROP TABLE IF EXISTS heats")
    cursor.execute("DROP TABLE IF EXISTS participants")
    cursor.execute("DROP TABLE IF EXISTS events")
    cursor.execute("DROP TABLE IF EXISTS teams")

    cursor.execute("""
        CREATE TABLE teams (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE
        )
    """)
    cursor.execute("""
        CREATE TABLE participants (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            team_id INTEGER NOT NULL,
            last_name TEXT NOT NULL,
            first_name TEXT NOT NULL,
            gender TEXT,
            FOREIGN KEY (team_id) REFERENCES teams(id)
        )
    """)
    cursor.execute("""
        CREATE TABLE events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE
        )
    """)
    cursor.execute("""
        CREATE TABLE heats (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_id INTEGER NOT NULL,
            FOREIGN KEY (event_id) REFERENCES events(id)
        )
    """)
    cursor.execute("""
        CREATE TABLE results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            participant_id INTEGER NOT NULL,
            heat_id INTEGER NOT NULL,
            time_seconds REAL,
            place_in_heat INTEGER,
            gender TEXT,
            FOREIGN KEY (participant_id) REFERENCES participants(id),
            FOREIGN KEY (heat_id) REFERENCES heats(id)
        )
    """)
    conn.commit()
    conn.close()
    print("✅ База данных создана")


def add_team(conn, name):
    cursor = conn.cursor()
    cursor.execute("INSERT OR IGNORE INTO teams (name) VALUES (?)", (name,))
    conn.commit()
    cursor.execute("SELECT id FROM teams WHERE name = ?", (name,))
    row = cursor.fetchone()
    return row[0] if row else None


def add_participant(conn, team_id, last_name, first_name, gender=None):
    cursor = conn.cursor()
    cursor.execute("""
        INSERT OR IGNORE INTO participants (team_id, last_name, first_name, gender)
        VALUES (?, ?, ?, ?)
    """, (team_id, last_name, first_name, gender))
    conn.commit()
    cursor.execute("""
        SELECT id FROM participants 
        WHERE last_name = ? AND first_name = ? AND team_id = ?
    """, (last_name, first_name, team_id))
    row = cursor.fetchone()
    return row[0] if row else None


def add_event(conn, name):
    cursor = conn.cursor()
    cursor.execute("INSERT OR IGNORE INTO events (name) VALUES (?)", (name,))
    conn.commit()
    cursor.execute("SELECT id FROM events WHERE name = ?", (name,))
    row = cursor.fetchone()
    return row[0] if row else None


def add_heat(conn, event_id):
    cursor = conn.cursor()
    cursor.execute("INSERT INTO heats (event_id) VALUES (?)", (event_id,))
    conn.commit()
    return cursor.lastrowid


def add_result(conn, participant_id, heat_id, time_seconds, gender=None):
    cursor = conn.cursor()
    cursor.execute("""
        INSERT OR REPLACE INTO results (participant_id, heat_id, time_seconds, gender)
        VALUES (?, ?, ?, ?)
    """, (participant_id, heat_id, time_seconds, gender))
    conn.commit()


# ============= ЗАГРУЗКА ДАННЫХ =============

def load_all_data():
    print("\n📥 Загрузка данных из Excel...")
    if not os.path.exists(INPUT_EXCEL):
        print(f"❌ Файл {INPUT_EXCEL} не найден!")
        return False

    df = pd.read_excel(INPUT_EXCEL, sheet_name="Результаты")
    print(f"   Найдено строк: {len(df)}")

    conn = get_connection()

    # Команды
    teams = set()
    for _, row in df.iterrows():
        team = row.get('teams') or row.get('team') or row.get('команда')
        if pd.notna(team):
            teams.add(str(team))
    for t in teams:
        add_team(conn, t)
        print(f"   Добавлена команда: {t}")

    # Дисциплины
    events = set()
    for _, row in df.iterrows():
        ev = row.get('event_name') or row.get('дисциплина')
        if pd.notna(ev):
            events.add(str(ev))
    for e in events:
        add_event(conn, e)
        print(f"   Добавлена дисциплина: {e}")

    # Участники и результаты
    for idx, row in df.iterrows():
        last = row.get('last_name') or row.get('фамилия')
        first = row.get('first_name') or row.get('имя')
        event = row.get('event_name') or row.get('дисциплина')
        time_raw = row.get('time_seconds') or row.get('время')
        gender_raw = row.get('gender') or row.get('пол')
        team_name = row.get('teams') or row.get('team') or row.get('команда')

        if pd.isna(last) or pd.isna(first) or pd.isna(event):
            continue

        time_sec = parse_time_to_seconds(time_raw)
        if time_sec is None:
            print(f"   ⚠️ Пропущена строка {idx}: время {time_raw}")
            continue

        gender = normalize_gender(gender_raw)
        team_id = add_team(conn, str(team_name) if pd.notna(team_name) else "Без команды")
        participant_id = add_participant(conn, team_id, str(last), str(first), gender)
        event_id = add_event(conn, str(event))
        heat_id = add_heat(conn, event_id)
        add_result(conn, participant_id, heat_id, time_sec, gender)

        gender_disp = gender if gender else "—"
        print(f"   Добавлен: {last} {first} ({gender_disp}) - {event} - {time_sec} сек")

    conn.commit()
    conn.close()
    print("✅ Все данные загружены")
    return True


# ============= РАСЧЕТ МЕСТ (РАЗДЕЛЬНО ПО ПОЛУ) =============

def calculate_places():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM events")
    events = cursor.fetchall()

    for (event_id,) in events:
        # Мужчины
        cursor.execute("""
            SELECT r.id, r.time_seconds
            FROM results r JOIN heats h ON r.heat_id = h.id
            WHERE h.event_id = ? AND r.gender = 'M'
            ORDER BY r.time_seconds
        """, (event_id,))
        for place, (rid, _) in enumerate(cursor.fetchall(), 1):
            cursor.execute("UPDATE results SET place_in_heat = ? WHERE id = ?", (place, rid))

        # Женщины
        cursor.execute("""
            SELECT r.id, r.time_seconds
            FROM results r JOIN heats h ON r.heat_id = h.id
            WHERE h.event_id = ? AND r.gender = 'F'
            ORDER BY r.time_seconds
        """, (event_id,))
        for place, (rid, _) in enumerate(cursor.fetchall(), 1):
            cursor.execute("UPDATE results SET place_in_heat = ? WHERE id = ?", (place, rid))

    conn.commit()
    conn.close()
    print("✅ Места рассчитаны (раздельно по полу)")


# ============= ФОРМИРОВАНИЕ EXCEL ОТЧЕТА (БЕЗ ЛИСТА "ПРИЗЕРЫ") =============

def build_excel_report():
    conn = get_connection()

    with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
        # 1. Протокол (все)
        query_all = """
        SELECT 
            e.name as Дисциплина,
            CASE r.gender WHEN 'M' THEN 'Мужчины' WHEN 'F' THEN 'Женщины' ELSE '—' END as Категория,
            t.name as Команда,
            p.last_name as Фамилия,
            p.first_name as Имя,
            r.time_seconds as Время_сек,
            r.place_in_heat as Место
        FROM results r
        JOIN participants p ON r.participant_id = p.id
        JOIN teams t ON p.team_id = t.id
        JOIN heats h ON r.heat_id = h.id
        JOIN events e ON h.event_id = e.id
        WHERE r.time_seconds IS NOT NULL
        ORDER BY e.name, r.gender, r.place_in_heat
        """
        df_all = pd.read_sql_query(query_all, conn)
        if not df_all.empty:
            df_all['Время'] = df_all['Время_сек'].apply(format_time_from_seconds)
            df_all.drop(columns=['Время_сек'], inplace=True)
        df_all.to_excel(writer, sheet_name='Протокол (все)', index=False)

        # 2. Командный зачет (общий)
        query_team = """
        SELECT 
            t.name as Команда,
            COUNT(DISTINCT p.id) as Участников,
            SUM(CASE 
                WHEN r.place_in_heat = 1 THEN 10
                WHEN r.place_in_heat = 2 THEN 8
                WHEN r.place_in_heat = 3 THEN 6
                WHEN r.place_in_heat = 4 THEN 5
                WHEN r.place_in_heat = 5 THEN 4
                WHEN r.place_in_heat = 6 THEN 3
                ELSE 1
            END) as Очки
        FROM results r
        JOIN participants p ON r.participant_id = p.id
        JOIN teams t ON p.team_id = t.id
        WHERE r.place_in_heat IS NOT NULL
        GROUP BY t.name
        ORDER BY Очки DESC
        """
        df_team = pd.read_sql_query(query_team, conn)
        df_team.to_excel(writer, sheet_name='Командный зачет (общий)', index=False)

        # 3. Командный зачет (мужчины)
        query_team_m = """
        SELECT 
            t.name as Команда,
            COUNT(DISTINCT p.id) as Участников,
            SUM(CASE 
                WHEN r.place_in_heat = 1 THEN 10
                WHEN r.place_in_heat = 2 THEN 8
                WHEN r.place_in_heat = 3 THEN 6
                WHEN r.place_in_heat = 4 THEN 5
                WHEN r.place_in_heat = 5 THEN 4
                WHEN r.place_in_heat = 6 THEN 3
                ELSE 1
            END) as Очки
        FROM results r
        JOIN participants p ON r.participant_id = p.id
        JOIN teams t ON p.team_id = t.id
        WHERE r.place_in_heat IS NOT NULL AND r.gender = 'M'
        GROUP BY t.name
        ORDER BY Очки DESC
        """
        df_team_m = pd.read_sql_query(query_team_m, conn)
        df_team_m.to_excel(writer, sheet_name='Командный зачет (М)', index=False)

        # 4. Командный зачет (женщины)
        query_team_f = """
        SELECT 
            t.name as Команда,
            COUNT(DISTINCT p.id) as Участников,
            SUM(CASE 
                WHEN r.place_in_heat = 1 THEN 10
                WHEN r.place_in_heat = 2 THEN 8
                WHEN r.place_in_heat = 3 THEN 6
                WHEN r.place_in_heat = 4 THEN 5
                WHEN r.place_in_heat = 5 THEN 4
                WHEN r.place_in_heat = 6 THEN 3
                ELSE 1
            END) as Очки
        FROM results r
        JOIN participants p ON r.participant_id = p.id
        JOIN teams t ON p.team_id = t.id
        WHERE r.place_in_heat IS NOT NULL AND r.gender = 'F'
        GROUP BY t.name
        ORDER BY Очки DESC
        """
        df_team_f = pd.read_sql_query(query_team_f, conn)
        df_team_f.to_excel(writer, sheet_name='Командный зачет (Ж)', index=False)

        # 5. Отдельные листы по дисциплинам (с разделением по полу)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM events")
        for (event_name,) in cursor.fetchall():
            # Мужчины
            q_m = f"""
            SELECT 
                r.place_in_heat as Место,
                t.name as Команда,
                p.last_name as Фамилия,
                p.first_name as Имя,
                r.time_seconds as Время
            FROM results r
            JOIN participants p ON r.participant_id = p.id
            JOIN teams t ON p.team_id = t.id
            JOIN heats h ON r.heat_id = h.id
            JOIN events e ON h.event_id = e.id
            WHERE e.name = '{event_name}' AND r.gender = 'M' AND r.place_in_heat IS NOT NULL
            ORDER BY r.place_in_heat
            """
            df_m = pd.read_sql_query(q_m, conn)
            if not df_m.empty:
                df_m['Время'] = df_m['Время'].apply(format_time_from_seconds)
                df_m.to_excel(writer, sheet_name=f'{event_name} (Мужчины)'[:31], index=False)

            # Женщины
            q_f = f"""
            SELECT 
                r.place_in_heat as Место,
                t.name as Команда,
                p.last_name as Фамилия,
                p.first_name as Имя,
                r.time_seconds as Время
            FROM results r
            JOIN participants p ON r.participant_id = p.id
            JOIN teams t ON p.team_id = t.id
            JOIN heats h ON r.heat_id = h.id
            JOIN events e ON h.event_id = e.id
            WHERE e.name = '{event_name}' AND r.gender = 'F' AND r.place_in_heat IS NOT NULL
            ORDER BY r.place_in_heat
            """
            df_f = pd.read_sql_query(q_f, conn)
            if not df_f.empty:
                df_f['Время'] = df_f['Время'].apply(format_time_from_seconds)
                df_f.to_excel(writer, sheet_name=f'{event_name} (Женщины)'[:31], index=False)

    conn.close()
    print(f"✅ Протокол сохранен: {OUTPUT_EXCEL}")


# ============= ВЫВОД В КОНСОЛЬ (С ПРОБЕЛАМИ МЕЖДУ КАТЕГОРИЯМИ) =============

def print_results():
    conn = get_connection()
    cursor = conn.cursor()

    print("\n" + "=" * 70)
    print("🏆 ПРОТОКОЛ СОРЕВНОВАНИЙ 🏆")
    print("=" * 70)

    cursor.execute("SELECT id, name FROM events")
    events = cursor.fetchall()

    for idx, (event_id, event_name) in enumerate(events):
        if idx > 0:
            print()  # Пустая строка между дисциплинами

        print(f"\n📏 {event_name}")

        # ---- Мужчины ----
        print("\n   🚹 Мужчины:")
        print("   " + "-" * 58)
        print(f"   {'Место':<8} {'Команда':<18} {'Участник':<22} {'Время':<12}")
        print("   " + "-" * 58)

        cursor.execute("""
            SELECT r.place_in_heat, t.name, p.last_name, p.first_name, r.time_seconds
            FROM results r
            JOIN participants p ON r.participant_id = p.id
            JOIN teams t ON p.team_id = t.id
            JOIN heats h ON r.heat_id = h.id
            WHERE h.event_id = ? AND r.gender = 'M' AND r.place_in_heat IS NOT NULL
            ORDER BY r.place_in_heat
        """, (event_id,))

        male_rows = cursor.fetchall()
        if male_rows:
            for place, team, last, first, t_sec in male_rows:
                medal = "🥇" if place == 1 else "🥈" if place == 2 else "🥉" if place == 3 else "  "
                time_str = format_time_from_seconds(t_sec)
                print(f"   {medal} {place:<5} {team:<18} {last} {first:<15} {time_str:<12}")
        else:
            print("   Нет участников")

        # Пустая строка между мужчинами и женщинами
        print()

        # ---- Женщины ----
        print("   🚺 Женщины:")
        print("   " + "-" * 58)
        print(f"   {'Место':<8} {'Команда':<18} {'Участник':<22} {'Время':<12}")
        print("   " + "-" * 58)

        cursor.execute("""
            SELECT r.place_in_heat, t.name, p.last_name, p.first_name, r.time_seconds
            FROM results r
            JOIN participants p ON r.participant_id = p.id
            JOIN teams t ON p.team_id = t.id
            JOIN heats h ON r.heat_id = h.id
            WHERE h.event_id = ? AND r.gender = 'F' AND r.place_in_heat IS NOT NULL
            ORDER BY r.place_in_heat
        """, (event_id,))

        female_rows = cursor.fetchall()
        if female_rows:
            for place, team, last, first, t_sec in female_rows:
                medal = "🥇" if place == 1 else "🥈" if place == 2 else "🥉" if place == 3 else "  "
                time_str = format_time_from_seconds(t_sec)
                print(f"   {medal} {place:<5} {team:<18} {last} {first:<15} {time_str:<12}")
        else:
            print("   Нет участниц")

    conn.close()


def print_statistics():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM teams")
    teams = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM participants")
    parts = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM events")
    evs = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM results")
    res = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM results WHERE gender = 'M'")
    male = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM results WHERE gender = 'F'")
    female = cursor.fetchone()[0]
    conn.close()

    print("\n📊 СТАТИСТИКА:")
    print(f"   • Команд: {teams}")
    print(f"   • Участников: {parts}")
    print(f"   • Дисциплин: {evs}")
    print(f"   • Результатов: {res} (М: {male}, Ж: {female})")


# ============= ГЛАВНАЯ =============

def main():
    print("\n" + "=" * 70)
    print("🏊‍♂️ СИСТЕМА ОБРАБОТКИ РЕЗУЛЬТАТОВ СОРЕВНОВАНИЙ 🏃‍♀️")
    print("РАЗДЕЛЬНЫЕ ЗАЧЕТЫ | ФОРМАТ ВРЕМЕНИ ММ.СС.СС")
    print("=" * 70)

    create_database()
    if not load_all_data():
        return
    calculate_places()
    print_statistics()
    print_results()
    build_excel_report()

    print("\n" + "=" * 70)
    print(f"✅ ГОТОВО!")
    print(f"📁 Протокол: {OUTPUT_EXCEL}")
    print(f"💾 База данных: {DB_PATH}")
    print("=" * 70)


if __name__ == "__main__":
    main()
