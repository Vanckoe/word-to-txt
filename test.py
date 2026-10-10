#!/usr/bin/env python3
"""Конвертер тестов из Word (.docx) в текстовый шаблон (TEXT / TOPIC / A: B: C: D: / CORRECT).

Запуск: двойной щелчок (выбор файлов) или перетаскивание .docx на программу.
Правильный ответ определяется по выделению: жирный, курсив, подчёркивание, цвет или маркер текста.
"""
import os, re, subprocess, sys
from docx import Document

LATIN = "ABCDEFGH"
OPT_RE = re.compile(r"^\s*([1-8]|[A-Za-zА-Яа-я])\s*\)\s*(.*)$", re.S)
Q_RE = re.compile(r"^\s*(?:ВОПРОС\s*№?\s*\d+\s*[.:]?|\d+\s*\.)\s*(.*)$", re.I | re.S)
VARIANT_RE = re.compile(r"^\s*Вариант", re.I)
# кириллица -> порядковый номер варианта (А Б В Г Д Е Ж З)
CYR = "АБВГДЕЖЗ"


def marked(par):
    """Есть ли в абзаце выделенный текст (признак правильного ответа)."""
    for r in par.runs:
        if not r.text.strip():
            continue
        f = r.font
        if r.bold or r.italic or r.underline or f.highlight_color or (f.color and f.color.type and str(f.color.rgb) not in ("000000", "None")):
            return True
    return False


def index_of(label):
    if label.isdigit():
        return int(label) - 1
    u = label.upper()
    if u in CYR:
        return CYR.index(u)
    if u in LATIN:
        return LATIN.index(u)
    return None


def convert(path):
    pars = [p for p in Document(path).paragraphs]
    nonempty = [p.text.strip() for p in pars if p.text.strip()]
    topic = nonempty[0] if nonempty else ""
    blocks, cur = [], None
    for p in pars:
        t = p.text.strip()
        if not t or VARIANT_RE.match(t) or t == topic:
            continue
        qm = Q_RE.match(t)
        if qm:
            cur = {"q": [qm.group(1).strip()] if qm.group(1).strip() else [], "opts": []}
            blocks.append(cur)
            continue
        if cur is None:
            continue
        m = OPT_RE.match(t)
        if m and (cur["q"] or not cur["opts"]) and index_of(m.group(1)) == len(cur["opts"]):
            cur["opts"].append((m.group(2).strip(), marked(p)))
        elif not cur["opts"]:
            cur["q"].append(t)
    out, warns = [], []
    for n, b in enumerate(blocks, 1):
        q = " ".join(b["q"])
        hits = [i for i, (_, mk) in enumerate(b["opts"]) if mk]
        if len(hits) != 1:
            warns.append(f"Вопрос {n}: найдено выделенных ответов: {len(hits)} - проверьте вручную")
        if len(b["opts"]) < 2:
            warns.append(f"Вопрос {n}: вариантов ответа: {len(b['opts'])}")
        lines = ["QUESTION:", f"TEXT: {q}", f"TOPIC: {topic}"]
        lines += [f"{LATIN[i]}: {txt}" for i, (txt, _) in enumerate(b["opts"])]
        lines.append("CORRECT: " + ("".join(LATIN[i] for i in hits) if hits else ""))
        out.append("\n".join(lines))
    return "\n\n".join(out) + "\n", len(blocks), warns


def run(files):
    """Конвертирует список файлов, результат кладёт рядом с исходником. Возвращает текст отчёта."""
    report = []
    for src in files:
        dst = re.sub(r"\.docx$", "", src, flags=re.I) + ".txt"
        try:
            text, n, warns = convert(src)
            with open(dst, "w", encoding="utf-8") as f:
                f.write(text)
            report.append(f"{os.path.basename(src)}: {n} вопросов -> {os.path.basename(dst)}")
            report += ["   ВНИМАНИЕ: " + w for w in warns]
        except Exception as e:
            report.append(f"{os.path.basename(src)}: ОШИБКА - {e}")
    return "\n".join(report)


def show(msg):
    if sys.platform == "darwin":
        script = '''on run argv
    display dialog (item 1 of argv) with title "Word → TXT" buttons {"OK"} default button "OK"
end run'''
        try:
            subprocess.run(["osascript", "-e", script, msg], check=True, capture_output=True, text=True)
            return
        except (OSError, subprocess.CalledProcessError):
            print(msg)
            return
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk(); root.withdraw()
        messagebox.showinfo("Word -> TXT", msg)
    except Exception:
        print(msg)


def choose_files():
    if sys.platform == "darwin":
        script = '''set selectedFiles to choose file with prompt "Выберите тесты Word" of type {"org.openxmlformats.wordprocessingml.document"} with multiple selections allowed
set paths to {}
repeat with selectedFile in selectedFiles
    set end of paths to POSIX path of selectedFile
end repeat
set AppleScript's text item delimiters to linefeed
return paths as text'''
        result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
        if result.returncode:
            if "(-128)" in result.stderr:  # пользователь нажал «Отмена»
                return []
            raise RuntimeError(result.stderr.strip() or "Не удалось открыть окно выбора файлов")
        return result.stdout.rstrip("\n").split("\n") if result.stdout else []
    from tkinter import Tk, filedialog
    root = Tk()
    root.withdraw()
    try:
        return list(filedialog.askopenfilenames(title="Выберите тесты Word", filetypes=[("Word", "*.docx")]))
    finally:
        root.destroy()


if __name__ == "__main__":
    files = sys.argv[1:]
    if not files:  # запуск двойным щелчком: открываем окно выбора файлов
        try:
            files = choose_files()
        except Exception as exc:
            show(f"Не удалось выбрать файлы: {exc}")
            sys.exit(1)
    if files:
        show(run(files))
