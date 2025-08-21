import os, json, runpy, shutil
import gradio as gr

def _ensure_status2_map():
    for c in [
        "/content/repo/examples/status2_map.json",
        "/content/colab-merge-project/examples/status2_map.json",
        "examples/status2_map.json",
        "./examples/status2_map.json",
    ]:
        if os.path.exists(c):
            try:
                shutil.copy(c, "/content/status2_map.json")
            except Exception:
                pass
            break

def _run_nbutest(file1, file2, date_str):
    if file1 is None or file2 is None:
        raise gr.Error("Загрузите оба файла (CSV/XLSX).")
    if not date_str:
        raise gr.Error("Введите дату ДД.ММ.ГГГГ (например, 01.08.2025).")
    settings = {"files": [file1.name, file2.name], "поточнадата": date_str}
    with open("/content/app_settings.json", "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)
    _ensure_status2_map()
    # Try typical locations
    for p in ["/content/repo/scripts/nbutest.py",
              "/content/colab-merge-project/scripts/nbutest.py",
              "scripts/nbutest.py", "./scripts/nbutest.py"]:
        if os.path.exists(p):
            script_path = p
            break
    else:
        raise gr.Error("Не найден scripts/nbutest.py")
    try:
        runpy.run_path(script_path, run_name="__main__")
    except SystemExit:
        pass
    # Locate output
    outs = [f"/content/result_{date_str}.xlsx", f"result_{date_str}.xlsx"]
    for c in outs:
        if os.path.exists(c):
            return c
    latest = None
    for fn in os.listdir("/content"):
        p = os.path.join("/content", fn)
        if fn.lower().endswith(".xlsx") and fn.startswith("result_"):
            if latest is None or os.path.getmtime(p) > os.path.getmtime(latest):
                latest = p
    if latest:
        return latest
    raise gr.Error("Итоговый Excel не найден. Откройте логи (Runtime → View logs).")

def launch_app():
    with gr.Blocks() as demo:
        gr.Markdown("### 🧩 Обработка по `nbutest.py` (логика сохранена, код скрыт)")
        with gr.Row():
            f1 = gr.File(label="Файл 1 (Excel/CSV)")
            f2 = gr.File(label="Файл 2 (Excel/CSV)")
        date_input = gr.Textbox(label="Дата (ДД.ММ.ГГГГ)", placeholder="например, 01.08.2025")
        go = gr.Button("Запустить обработку")
        out_file = gr.File(label="Результат (XLSX)")
        go.click(_run_nbutest, inputs=[f1, f2, date_input], outputs=out_file)
    demo.launch()


# === Original UI fragments preserved below (commented) ===

# --- fragment 1 ---
# # @title 🖥️ UI: Выберите 2 файла и дату → запуск nbutest.py с логом
# !pip -q install gradio>=4.0 openpyxl pandas
# 
# import os, json, runpy, shutil, io, traceback, time
# import gradio as gr
# from contextlib import redirect_stdout, redirect_stderr
# 
# # убедимся, что status2_map.json доступен в рабочей директории
# EXAMPLES_DIR = '/content/examples'
# os.makedirs(EXAMPLES_DIR, exist_ok=True)
# # пробуем скопировать из структуры репозитория, если надо
# for candidate in [
#     '/content/repo/examples/status2_map.json',
#     '/content/colab-merge-project/examples/status2_map.json',
#     'examples/status2_map.json'
# ]:
#     if os.path.exists(candidate):
#         shutil.copy(candidate, '/content/status2_map.json')
#         break
# 
# def run_nbutest(file1, file2, date_str):
#     # подготовка
#     log_buf = io.StringIO()
#     result_path = None
# 
#     try:
#         if file1 is None or file2 is None:
#             raise ValueError('Загрузите оба файла (CSV/XLSX).')
#         if not date_str:
#             raise ValueError('Введите дату в формате ДД.ММ.ГГГГ (например, 01.08.2025).')
# 
#         # формируем app_settings.json как ожидает скрипт
#         settings = {
#             "files": [file1.name, file2.name],
#             "поточнадата": date_str
#         }
#         with open('/content/app_settings.json', 'w', encoding='utf-8') as f:
#             json.dump(settings, f, ensure_ascii=False, indent=2)
# 
#         # диагностическая информация
#         print("Рабочая папка:", os.getcwd(), file=log_buf)
#         print("Файл1:", file1.name, file=log_buf)
#         print("Файл2:", file2.name, file=log_buf)
#         print("Дата:", date_str, file=log_buf)
#         print("status2_map.json exists?:", os.path.exists('/content/status2_map.json'), file=log_buf)
# 
#         # запускаем скрипт и перехватываем stdout/stderr
#         with redirect_stdout(log_buf), redirect_stderr(log_buf):
#             try:
#                 # путь к скрипту (если вы клонировали репо в /content/repo)
#                 script_path = '/content/repo/scripts/nbutest.py'
#                 if not os.path.exists(script_path):
#                     # вариант, если ноутбук запущен прямо из архива
#                     script_path = '/content/colab-merge-project/scripts/nbutest.py'
#                 print("Запуск:", script_path)
#                 runpy.run_path(script_path, run_name='__main__')
#             except SystemExit:
#                 # некоторые скрипты вызывают sys.exit — это ок, файл уже мог создаться
#                 pass
# 
#         # ищем выходной .xlsx
#         # часто nbutest.py пишет result_ДД.ММ.ГГГГ.xlsx
#         candidates = [
#             f"/content/result_{date_str}.xlsx",
#             f"result_{date_str}.xlsx",
#         ]
#         for c in candidates:
#             if os.path.exists(c):
#                 result_path = c
#                 break
#         if result_path is None:
#             # fallback: самый свежий result_*.xlsx в /content
#             latest = None
#             for fname in os.listdir('/content'):
#                 if fname.lower().endswith('.xlsx') and fname.startswith('result_'):
#                     p = os.path.join('/content', fname)
#                     if latest is None or os.path.getmtime(p) > os.path.getmtime(latest):
#                         latest = p
#             result_path = latest
# 
#         if result_path is None:
#             raise FileNotFoundError("Выходной Excel не найден. Смотрите лог ниже.")
# 
#         return result_path, log_buf.getvalue()
# 
#     except Exception:
#         # полный трейсбек в лог
#         tb = traceback.format_exc()
#         return None, (log_buf.getvalue() + "\n--- TRACEBACK ---\n" + tb)
# 
# with gr.Blocks() as demo:
#     gr.Markdown(
#         "### 📘 Обработка по полной логике `scripts/nbutest.py`\n"
#         "Выберите **два файла** и введите **дату** (ДД.ММ.ГГГГ). "
#         "Интерфейс только формирует `app_settings.json` и запускает ваш скрипт. "
#         "Ниже выводится подробный **лог выполнения**."
#     )
#     with gr.Row():
#         f1 = gr.File(label='Файл 1 (Excel/CSV)')
#         f2 = gr.File(label='Файл 2 (Excel/CSV)')
#     date_input = gr.Textbox(label='Дата (ДД.ММ.ГГГГ)', placeholder='например, 01.08.2025')
#     go = gr.Button('Запустить обработку')
#     out_file = gr.File(label='Результат (XLSX)')
#     logs = gr.Textbox(label='Лог выполнения', lines=18)
#     go.click(run_nbutest, inputs=[f1, f2, date_input], outputs=[out_file, logs])
# 
# demo.launch()
# 
