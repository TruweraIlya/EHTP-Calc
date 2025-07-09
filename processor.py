import pandas as pd
def process_data(df, max_lines_per_box, breaker_nominal, max_pipe, max_limiter, rk_prefix, panel_prefix):
    lines = []
    for _, row in df.iterrows():
        lines.append({
            'Название линии': row.iloc[0],
            'Тип процесса': row.iloc[7],
            'Тип регулирования': row.iloc[16],
            'Рабочий ток': row.iloc[17],
            'Пусковой ток': row.iloc[18],
            'Нагрузка': row.iloc[19],
            'Номер линии': row.iloc[22],
        })

    # Этап 1: Компоновка линий в коробки
    boxes = []
    for process in set(l['Тип процесса'] for l in lines):
        process_lines = [l for l in lines if l['Тип процесса'] == process]
        manual_boxes = [[l] for l in process_lines if l['Пусковой ток'] > breaker_nominal]
        auto_lines = [l for l in process_lines if l['Пусковой ток'] <= breaker_nominal]

        auto_lines.sort(key=lambda x: x['Пусковой ток'], reverse=True)
        while auto_lines:
            box = [auto_lines.pop(0)]
            total = box[0]['Пусковой ток']
            i = 0
            while i < len(auto_lines):
                candidate = auto_lines[i]
                if total + candidate['Пусковой ток'] <= breaker_nominal and len(box) < max_lines_per_box:
                    box.append(candidate)
                    total += candidate['Пусковой ток']
                    auto_lines.pop(i)
                else:
                    i += 1
            boxes.append(box)
        boxes.extend(manual_boxes)

    # Этап 2: Компоновка коробок в шкафы
    panels = []
    panel_id = 1
    unassigned_boxes = boxes.copy()
    while unassigned_boxes:
        process_type = unassigned_boxes[0][0]['Тип процесса']
        panel = {'Шкаф': f"{panel_prefix} {panel_id}", 'Тип процесса': process_type, 'Pipe': 0, 'Limiter': 0, 'РК': []}
        remaining = []
        for box in unassigned_boxes:
            if box[0]['Тип процесса'] != process_type:
                remaining.append(box)
                continue
            rk_type = box[0]['Тип регулирования']
            if rk_type in ["Pipe Sensing", "Ambient Proportional Control"] and panel['Pipe'] < max_pipe:
                panel['Pipe'] += 1
                panel['РК'].append(box)
            elif rk_type == "Controlled Design (w/Limiter)" and panel['Limiter'] < max_limiter:
                panel['Limiter'] += 1
                panel['РК'].append(box)
            else:
                remaining.append(box)
        panels.append(panel)
        unassigned_boxes = remaining
        panel_id += 1

    # Подготовка DataFrame для Excel
    df_lines, df_boxes, df_panels = [], [], []
    for box_idx, box in enumerate(boxes, 1):
        rk_name = f"{rk_prefix} {box_idx}"
        for line in box:
            panel_name = next((p['Шкаф'] for p in panels if box in p['РК']), None)
            df_lines.append({
                'Название линии': line['Название линии'],
                'Номер линии': line['Номер линии'],
                'Тип процесса': line['Тип процесса'],
                'Тип регулирования': line['Тип регулирования'],
                'Распредкоробка': rk_name,
                'Шкаф': panel_name,
                'Ручной автомат': line['Пусковой ток'] > breaker_nominal
            })
        row = {
            'Распредкоробка': rk_name,
            'Тип процесса': box[0]['Тип процесса'],
            'Тип регулирования': box[0]['Тип регулирования'],
            'Суммарный ток': sum(l['Пусковой ток'] for l in box),
            'Количество линий': len(box)
        }
        for i, line in enumerate(box, 1):
            row[str(i)] = line['Название линии']
        df_boxes.append(row)
    for panel in panels:
        row = {
            'Шкаф': panel['Шкаф'],
            'Тип процесса': panel['Тип процесса'],
            'Количество РК': len(panel['РК']),
            'Количество РК PS и PASC': panel['Pipe'],
            'Количество РК w/ Limiter': panel['Limiter'],
            'Суммарный ток': sum(sum(l['Пусковой ток'] for l in box) for box in panel['РК']),
            'Суммарная мощность': sum(sum(l['Нагрузка'] for l in box) for box in panel['РК'])
        }
        for i, box in enumerate(panel['РК'], 1):
            rk_name = f"{rk_prefix} {boxes.index(box) + 1}"
            row[str(i)] = rk_name
        df_panels.append(row)
    return pd.DataFrame(df_lines), pd.DataFrame(df_boxes), pd.DataFrame(df_panels)
