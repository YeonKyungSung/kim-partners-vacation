# -*- coding: utf-8 -*-
import openpyxl, sys, re, json
from datetime import datetime
sys.stdout.reconfigure(encoding='utf-8')

CURRENT = ['손승웅','안희공','여현정','조윤미','안희진','이세림','윤승빈','한성종',
           '박하람','이진호','김태희','백운선','김태훈','유승아','정지인','김유진',
           '명지숙','구혜수','문희원','임동욱','김대영','유예람','김혜진','김예은',
           '남진영','성연경','문정하']

EMP_HIRE = {
    '손승웅':'2020-10-19','안희공':'2020-12-01','여현정':'2021-06-01',
    '조윤미':'2022-05-11','안희진':'2022-08-01','이세림':'2022-09-14',
    '윤승빈':'2022-09-22','한성종':'2022-12-06','박하람':'2023-03-20',
    '이진호':'2023-04-24','김태희':'2023-05-02','백운선':'2023-05-15',
    '김태훈':'2023-06-01','유승아':'2023-06-12','정지인':'2023-12-14',
    '김유진':'2024-01-01','명지숙':'2024-01-22','구혜수':'2024-06-17',
    '문희원':'2025-10-23','임동욱':'2025-11-03','김대영':'2025-11-03',
    '유예람':'2026-01-02','김혜진':'2026-01-02','김예은':'2025-05-07',
    '남진영':'2025-12-01','성연경':'2025-08-25','문정하':'2026-03-16',
}

HOLIDAYS = ['신정','설날','추석','삼일절','어린이날','현충일','광복절','개천절','한글날',
            '성탄절','크리스마스','대체공휴일','대체휴일','근로자의날','대통령 선거일',
            '선거일','대통령선거','부처님오신날']

wb = openpyxl.load_workbook(r'C:\Users\dexte\Downloads\2025년 서울사무소 휴가대장.xlsx', data_only=True)
vacation_data = {name: [] for name in CURRENT}

def parse_cell(cell_val, emp_names):
    results = []
    parts = re.split(r'[,\n]', cell_val)
    for part in parts:
        part = part.strip()
        if not part: continue
        slash_parts = re.split(r'\s*/\s*', part)
        for sp in slash_parts:
            sp = sp.strip()
            if not sp or len(sp) < 2: continue
            for name in emp_names:
                if name in sp:
                    vac_type = '연차'
                    if '오전반차' in sp or '오전 반차' in sp: vac_type = '오전반차'
                    elif '오후반차' in sp or '오후 반차' in sp: vac_type = '오후반차'
                    elif '생일' in sp: vac_type = '생일연차'
                    elif '반차' in sp: vac_type = '오전반차'
                    results.append((name, vac_type))
    return results

def parse_day_row(day_vals_dict, header_year, header_month, is_first_row=False):
    """Parse a day number row, handling month boundaries.
    is_first_row: True if this is the first day-number row right after a month header.
    Returns dict: col_idx -> (year, month, day)"""
    result = {}
    sorted_cols = sorted(day_vals_dict.keys())
    if not sorted_cols:
        return result

    days_list = [day_vals_dict[c] for c in sorted_cols]

    # Find boundary: where day number drops (e.g., 30, 31, 1, 2)
    boundary_idx = -1
    for ci in range(1, len(days_list)):
        if days_list[ci] < days_list[ci-1] and days_list[ci] <= 7:
            boundary_idx = ci
            break

    if boundary_idx == -1:
        # No boundary - all same month
        # Only treat as previous month if this is the FIRST row after a month header
        # AND all days are >= 20 (e.g., header=June, days=25,26,27,28,29,30,31 = May)
        if is_first_row and all(d >= 20 for d in days_list):
            pm = header_month - 1
            py = header_year
            if pm < 1: pm = 12; py -= 1
            for col in sorted_cols:
                result[col] = (py, pm, day_vals_dict[col])
        else:
            for col in sorted_cols:
                result[col] = (header_year, header_month, day_vals_dict[col])
    else:
        # Before boundary = previous month
        pm = header_month - 1
        py = header_year
        if pm < 1: pm = 12; py -= 1
        for ci, col in enumerate(sorted_cols):
            if ci < boundary_idx:
                result[col] = (py, pm, day_vals_dict[col])
            else:
                result[col] = (header_year, header_month, day_vals_dict[col])

    return result

for sheet_name in ['2022년','2023년','2024년','2025년','2026년']:
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    header_month = None
    header_year = None
    day_columns = {}
    i = 0
    while i < len(rows):
        row = rows[i]
        # Month header
        if row[1] and isinstance(row[1], datetime):
            header_year = row[1].year
            header_month = row[1].month
            i += 2  # skip dow
            if i >= len(rows): break
            day_row = rows[i]
            day_vals = {}
            for col_idx in range(1, min(len(day_row), 8)):
                if day_row[col_idx]:
                    try:
                        d = int(float(str(day_row[col_idx]).replace('.0','')))
                        if 1 <= d <= 31: day_vals[col_idx] = d
                    except: pass
            day_columns = parse_day_row(day_vals, header_year, header_month, is_first_row=True)
            i += 1
            continue
        elif row[1] and isinstance(row[1], str) and re.match(r'20\d{2}년\s*\d+월', str(row[1])):
            m = re.match(r'(20\d{2})년\s*(\d+)월', str(row[1]))
            if m:
                header_year = int(m.group(1))
                header_month = int(m.group(2))
            i += 2
            if i >= len(rows): break
            day_row = rows[i]
            day_vals = {}
            for col_idx in range(1, min(len(day_row), 8)):
                if day_row[col_idx]:
                    try:
                        d = int(float(str(day_row[col_idx]).replace('.0','')))
                        if 1 <= d <= 31: day_vals[col_idx] = d
                    except: pass
            day_columns = parse_day_row(day_vals, header_year, header_month, is_first_row=True)
            i += 1
            continue

        if not header_month:
            i += 1
            continue

        # Week day-number row
        day_vals = {}
        day_count = 0
        for col_idx in range(1, min(len(row), 8)):
            if row[col_idx]:
                try:
                    d = int(float(str(row[col_idx]).replace('.0','')))
                    if 1 <= d <= 31: day_vals[col_idx] = d; day_count += 1
                except: pass
        if day_count >= 2:
            day_columns = parse_day_row(day_vals, header_year, header_month)
            # Update header_month if we crossed into next month
            for col, (yr, mo, d) in day_columns.items():
                if mo == header_month + 1 or (header_month == 12 and mo == 1):
                    header_month = mo
                    header_year = yr
                    break
            i += 1
            continue

        # Data row
        if header_month and day_columns:
            for col_idx, (yr, mo, day) in day_columns.items():
                if col_idx < len(row) and row[col_idx]:
                    cell_val = str(row[col_idx]).strip()
                    if not cell_val or cell_val == 'None': continue
                    if any(h in cell_val for h in HOLIDAYS) and not any(name in cell_val for name in CURRENT):
                        continue
                    entries = parse_cell(cell_val, CURRENT)
                    for emp_name, vac_type in entries:
                        try:
                            date = datetime(yr, mo, day)
                            hire = datetime.strptime(EMP_HIRE.get(emp_name,'2020-01-01'), '%Y-%m-%d')
                            if date < hire: continue
                            date_str = date.strftime('%Y-%m-%d')
                            if not any(v['date']==date_str and v['type']==vac_type for v in vacation_data[emp_name]):
                                vacation_data[emp_name].append({'date': date_str, 'type': vac_type})
                        except Exception as e:
                            pass
        i += 1

# Sort
for name in CURRENT:
    vacation_data[name] = sorted(vacation_data[name], key=lambda x: x['date'])

# Compare
with open(r'C:\Users\dexte\Desktop\07_knp-claude-workspace\vacation-system\vacation_data.json', 'r', encoding='utf-8') as f:
    old_data = json.load(f)

print('=== 변경사항 ===')
total_changes = 0
for name in CURRENT:
    old_dates = set(v['date']+'|'+v['type'] for v in old_data.get(name,[]))
    new_dates = set(v['date']+'|'+v['type'] for v in vacation_data[name])
    added = new_dates - old_dates
    removed = old_dates - new_dates
    if added or removed:
        total_changes += len(added) + len(removed)
        print(f'\n{name}:')
        for a in sorted(added): print(f'  + {a}')
        for r in sorted(removed): print(f'  - {r}')

if total_changes == 0:
    print('변경사항 없음')
else:
    print(f'\n총 {total_changes}건 변경')

# Save
with open(r'C:\Users\dexte\Desktop\07_knp-claude-workspace\vacation-system\vacation_data.json', 'w', encoding='utf-8') as f:
    json.dump(vacation_data, f, ensure_ascii=False, indent=2)
print('vacation_data.json 업데이트 완료')
