
import os, json, runpy, shutil
import gradio as gr

def _ensure_status2_map():
    for c in [
        "/content/repo/examples/status2_map.json",
        "/content/tpok003_full/examples/status2_map.json",
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

    # Ищем скрипт
    for p in ["/content/repo/scripts/nbutest.py",
              "/content/tpok003_full/scripts/nbutest.py",
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

    # Находим выходной файл
    for c in [f"/content/result_{date_str}.xlsx", f"result_{date_str}.xlsx"]:
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
    raise gr.Error("Итоговый Excel не найден. Проверьте входные файлы и дату.")

def launch_app():
    with gr.Blocks() as demo:
        gr.Markdown("### 🧩 Обработка по `nbutest.py` — код скрыт")
        with gr.Row():
            f1 = gr.File(label="Файл 1 (Excel/CSV)")
            f2 = gr.File(label="Файл 2 (Excel/CSV)")
        date_input = gr.Textbox(label="Дата (ДД.ММ.ГГГГ)", placeholder="например, 01.08.2025")
        go = gr.Button("Запустить обработку")
        out_file = gr.File(label="Результат (XLSX)")
        go.click(_run_nbutest, inputs=[f1, f2, date_input], outputs=out_file)
    demo.launch()
