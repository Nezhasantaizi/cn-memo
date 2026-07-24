# -*- coding: utf-8 -*-
import sys
import os
import sqlite3
import csv
from datetime import datetime
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLabel, QTextEdit, QLineEdit, QListWidget,
                             QListWidgetItem, QDialog, QFormLayout, QMessageBox,
                             QFileDialog, QMenu, QComboBox)
from PyQt6.QtCore import Qt, QTranslator, QLocale
from PyQt6.QtGui import QAction
from fsrs import Scheduler, Card, Rating

# Office/PDF 解析依赖（修复odfpy新版导入错误，DrawPage来自odfdo）
import docx
import openpyxl
from pptx import Presentation
from odf.opendocument import OpenDocumentText, OpenDocumentSpreadsheet, OpenDocumentPresentation
from odf.text import P
from odf.table import Table, TableRow, TableCell
# 唯一修改行：odf.draw → odfdo.draw
from odfdo.draw_page import DrawPage
import pdfplumber

DB_PATH = "med_fsrs.db"
scheduler = Scheduler()

# ====================== 数据库初始化【新增笔记表】 ======================
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    # 开启外键约束
    cur.execute("PRAGMA foreign_keys = ON;")
    # 卡组表
    cur.execute('''
    CREATE TABLE IF NOT EXISTS decks (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        create_time TEXT
    )
    ''')
    # 卡片表
    cur.execute('''
    CREATE TABLE IF NOT EXISTS cards (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        deck_id INTEGER,
        question TEXT,
        options TEXT,
        answer TEXT,
        analysis TEXT,
        tags TEXT,
        create_time TEXT,
        FOREIGN KEY(deck_id) REFERENCES decks(id) ON DELETE CASCADE
    )
    ''')
    # 复习状态绑定卡片
    cur.execute('''
    CREATE TABLE IF NOT EXISTS review_state (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        card_id INTEGER UNIQUE,
        due TEXT,
        stability REAL,
        difficulty REAL,
        elapsed_days INTEGER,
        scheduled_days INTEGER,
        reps INTEGER,
        lapses INTEGER,
        last_review TEXT,
        FOREIGN KEY(card_id) REFERENCES cards(id) ON DELETE CASCADE
    )
    ''')
    # 【新增笔记表】绑定卡组，卡组删除同步删除笔记
    cur.execute('''
    CREATE TABLE IF NOT EXISTS notes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        deck_id INTEGER,
        title TEXT NOT NULL,
        content TEXT,
        create_time TEXT,
        update_time TEXT,
        FOREIGN KEY(deck_id) REFERENCES decks(id) ON DELETE CASCADE
    )
    ''')
    conn.commit()
    conn.close()

def db_query(sql, args=()):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("PRAGMA foreign_keys = ON;")
    cur.execute(sql, args)
    res = cur.fetchall()
    conn.close()
    return res

def db_exec(sql, args=()):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("PRAGMA foreign_keys = ON;")
    cur.execute(sql, args)
    last_id = cur.lastrowid
    conn.commit()
    conn.close()
    return last_id

# ====================== 通用弹窗组件 ======================
class InputDialog(QDialog):
    @staticmethod
    def getText(parent, title, label):
        dlg = InputDialog(parent)
        dlg.setWindowTitle(title)
        layout = QVBoxLayout(dlg)
        layout.addWidget(QLabel(label))
        edit = QLineEdit()
        layout.addWidget(edit)
        btn_row = QHBoxLayout()
        ok_btn = QPushButton(QApplication.translate("Dialog", "OK"))
        cancel_btn = QPushButton(QApplication.translate("Dialog", "Cancel"))
        btn_row.addWidget(ok_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

        def ok():
            dlg.result_val = edit.text()
            dlg.accept()

        ok_btn.clicked.connect(ok)
        cancel_btn.clicked.connect(dlg.reject)
        dlg.result_val = ""
        res = dlg.exec()
        return dlg.result_val, res == QDialog.DialogCode.Accepted

# ====================== 笔记编辑弹窗（新增） ======================
class NoteEditDialog(QDialog):
    def __init__(self, deck_id, note_data=None, deck_list=None):
        super().__init__()
        self.deck_id = deck_id
        self.note_data = note_data
        self.all_decks = deck_list
        self.setWindowTitle(QApplication.translate("Note", "Note Editor"))
        self.resize(550, 600)
        layout = QFormLayout(self)
        # 笔记标题
        self.title_edit = QLineEdit()
        layout.addRow(QApplication.translate("Note", "Note Title:"), self.title_edit)
        # 笔记正文
        self.content_edit = QTextEdit()
        layout.addRow(QApplication.translate("Note", "Note Content:"), self.content_edit)
        # 移动卡组下拉框
        self.deck_combo = QComboBox()
        for did, dname in self.all_decks:
            self.deck_combo.addItem(dname, did)
        layout.addRow(QApplication.translate("Note", "Belong Deck:"), self.deck_combo)
        # 按钮
        btn_layout = QHBoxLayout()
        save_btn = QPushButton(QApplication.translate("Note", "Save Note"))
        cancel_btn = QPushButton(QApplication.translate("Note", "Cancel"))
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addRow(btn_layout)
        save_btn.clicked.connect(self.save_note)
        cancel_btn.clicked.connect(self.reject)
        # 编辑模式填充数据
        if self.note_data:
            nid, d_id, title, content, _, _ = self.note_data
            self.title_edit.setText(title)
            self.content_edit.setPlainText(content)
            # 选中当前卡组
            idx = self.deck_combo.findData(d_id)
            if idx != -1:
                self.deck_combo.setCurrentIndex(idx)

    def save_note(self):
        title = self.title_edit.text().strip()
        content = self.content_edit.toPlainText().strip()
        target_deck_id = self.deck_combo.currentData()
        if not title:
            QMessageBox.warning(self, QApplication.translate("Msg", "Warning"),
                                QApplication.translate("Msg", "Note title cannot be empty"))
            return
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if self.note_data:
            # 更新笔记/移动卡组
            nid = self.note_data[0]
            db_exec("UPDATE notes SET deck_id=?,title=?,content=?,update_time=? WHERE id=?",
                    (target_deck_id, title, content, now, nid))
        else:
            # 新建笔记
            db_exec('''INSERT INTO notes(deck_id,title,content,create_time,update_time)
            VALUES (?,?,?,?,?)''', (target_deck_id, title, content, now, now))
        self.accept()

# ====================== 卡片编辑弹窗 ======================
class CardEditDialog(QDialog):
    def __init__(self, deck_id, card_data=None):
        super().__init__()
        self.deck_id = deck_id
        self.card_data = card_data
        self.setWindowTitle(QApplication.translate("CardEdit", "Edit Question Card"))
        self.resize(520, 560)
        layout = QFormLayout(self)
        self.question_edit = QTextEdit()
        self.option_edit = QTextEdit()
        self.answer_edit = QLineEdit()
        self.analysis_edit = QTextEdit()
        self.tag_edit = QLineEdit()
        layout.addRow(QApplication.translate("CardEdit", "Question:"), self.question_edit)
        layout.addRow(QApplication.translate("CardEdit", "Options:"), self.option_edit)
        layout.addRow(QApplication.translate("CardEdit", "Answer:"), self.answer_edit)
        layout.addRow(QApplication.translate("CardEdit", "Analysis:"), self.analysis_edit)
        layout.addRow(QApplication.translate("CardEdit", "Tags (comma separated):"), self.tag_edit)
        btn_layout = QHBoxLayout()
        save_btn = QPushButton(QApplication.translate("CardEdit", "Save"))
        cancel_btn = QPushButton(QApplication.translate("CardEdit", "Cancel"))
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addRow(btn_layout)
        save_btn.clicked.connect(self.save_data)
        cancel_btn.clicked.connect(self.reject)
        if card_data:
            _, q, opt, a, ana, t = card_data
            self.question_edit.setPlainText(q)
            self.option_edit.setPlainText(opt)
            self.answer_edit.setText(a)
            self.analysis_edit.setPlainText(ana)
            self.tag_edit.setText(t)

    def save_data(self):
        q = self.question_edit.toPlainText().strip()
        opt = self.option_edit.toPlainText().strip()
        a = self.answer_edit.text().strip()
        ana = self.analysis_edit.toPlainText().strip()
        tag = self.tag_edit.text().strip()
        if not q:
            QMessageBox.warning(self, QApplication.translate("Msg", "Warning"),
                                QApplication.translate("Msg", "Question cannot be empty"))
            return
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if self.card_data:
            cid = self.card_data[0]
            db_exec("UPDATE cards SET question=?,options=?,answer=?,analysis=?,tags=? WHERE id=?",
                    (q, opt, a, ana, tag, cid))
        else:
            cid = db_exec(
                "INSERT INTO cards(deck_id,question,options,answer,analysis,tags,create_time) VALUES (?,?,?,?,?,?,?)",
                (self.deck_id, q, opt, a, ana, tag, now))
            card = Card()
            db_exec('''INSERT INTO review_state(card_id,due,stability,difficulty,elapsed_days,scheduled_days,reps,lapses,last_review)
            VALUES (?,?,?,?,?,?,?,?,?)''',
                    (cid, card.due.isoformat(), card.stability, card.difficulty,
                     card.elapsed_days, card.scheduled_days, card.reps, card.lapses, None))
        self.accept()

# ====================== 复习弹窗 ======================
class ReviewDialog(QDialog):
    def __init__(self, card_id):
        super().__init__()
        self.card_id = card_id
        self.setWindowTitle(QApplication.translate("Review", "Card Review"))
        self.resize(640, 500)
        layout = QVBoxLayout(self)
        self.q_label = QLabel()
        self.q_label.setWordWrap(True)
        self.q_label.setStyleSheet("font-size:15px;")
        self.opt_label = QLabel()
        self.opt_label.setWordWrap(True)
        self.opt_label.setStyleSheet("font-size:14px;")
        self.ans_label = QLabel("???")
        self.ans_label.setWordWrap(True)
        self.ans_label.setStyleSheet("font-size:14px;color:#227722;")
        self.ana_label = QLabel()
        self.ana_label.setWordWrap(True)
        show_btn = QPushButton(QApplication.translate("Review", "Show Answer & Analysis"))
        rating_layout = QHBoxLayout()
        again_btn = QPushButton(QApplication.translate("Rating", "Again"))
        hard_btn = QPushButton(QApplication.translate("Rating", "Hard"))
        good_btn = QPushButton(QApplication.translate("Rating", "Good"))
        easy_btn = QPushButton(QApplication.translate("Rating", "Easy"))
        rating_layout.addWidget(again_btn)
        rating_layout.addWidget(hard_btn)
        rating_layout.addWidget(good_btn)
        rating_layout.addWidget(easy_btn)
        layout.addWidget(QLabel(f"<b>{QApplication.translate('Review','Question')}</b>"))
        layout.addWidget(self.q_label)
        layout.addWidget(QLabel(f"<b>{QApplication.translate('Review','Options')}</b>"))
        layout.addWidget(self.opt_label)
        layout.addWidget(show_btn)
        layout.addWidget(QLabel(f"<b>{QApplication.translate('Review','Answer & Analysis')}</b>"))
        layout.addWidget(self.ans_label)
        layout.addWidget(self.ana_label)
        layout.addLayout(rating_layout)
        self.raw_ans = ""
        self.raw_ana = ""
        self.card = None
        self.load_card()
        show_btn.clicked.connect(self.show_content)
        again_btn.clicked.connect(lambda: self.do_review(Rating.Again))
        hard_btn.clicked.connect(lambda: self.do_review(Rating.Hard))
        good_btn.clicked.connect(lambda: self.do_review(Rating.Good))
        easy_btn.clicked.connect(lambda: self.do_review(Rating.Easy))

    def load_card(self):
        q, opt, a, ana = db_query("SELECT question,options,answer,analysis FROM cards WHERE id=?", (self.card_id,))[0]
        self.q_label.setText(q)
        self.opt_label.setText(opt)
        self.raw_ans = f"{QApplication.translate('Review','Correct Answer')}：{a}"
        self.raw_ana = f"{QApplication.translate('Review','Analysis')}：{ana}"
        row = db_query('''SELECT stability,difficulty,elapsed_days,scheduled_days,reps,lapses,last_review,due FROM review_state WHERE card_id=?''',
                       (self.card_id,))[0]
        self.card = Card(
            stability=row[0],
            difficulty=row[1],
            elapsed_days=row[2],
            scheduled_days=row[3],
            reps=row[4],
            lapses=row[5]
        )

    def show_content(self):
        self.ans_label.setText(self.raw_ans)
        self.ana_label.setText(self.raw_ana)

    def do_review(self, rating):
        now = datetime.now()
        new_card, log = scheduler.review_card(self.card, rating, now)
        db_exec('''UPDATE review_state SET due=?,stability=?,difficulty=?,elapsed_days=?,scheduled_days=?,reps=?,lapses=?,last_review=?
        WHERE card_id=?''',
                (new_card.due.isoformat(), new_card.stability, new_card.difficulty,
                 new_card.elapsed_days, new_card.scheduled_days, new_card.reps,
                 new_card.lapses, now.isoformat(), self.card_id))
        self.accept()

# ====================== 文件导入/导出工具类（全格式双向支持） ======================
class FileIOHandler:
    @staticmethod
    def extract_text(file_path):
        """导入：统一解析所有支持格式，返回纯文本块列表"""
        ext = os.path.splitext(file_path)[1].lower()
        blocks = []
        try:
            if ext == ".txt":
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    text = f.read()
                blocks = [p.strip() for p in text.split("\n\n") if p.strip()]
            elif ext in (".csv", ".tsv"):
                delimiter = "\t" if ext == ".tsv" else ","
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    reader = csv.reader(f, delimiter=delimiter)
                    for row in reader:
                        blocks.append(" | ".join([str(x) for x in row]))
            elif ext == ".docx":
                doc = docx.Document(file_path)
                for para in doc.paragraphs:
                    if para.text.strip():
                        blocks.append(para.text.strip())
            elif ext == ".xlsx":
                wb = openpyxl.load_workbook(file_path, read_only=True)
                for sheet in wb.worksheets:
                    for row in sheet.iter_rows(values_only=True):
                        blocks.append(" | ".join([str(x) for x in row if x is not None]))
            elif ext == ".pptx":
                prs = Presentation(file_path)
                for slide in prs.slides:
                    slide_text = []
                    for shape in slide.shapes:
                        if hasattr(shape, "text") and shape.text.strip():
                            slide_text.append(shape.text.strip())
                    blocks.append("\n".join(slide_text))
            elif ext == ".odt":
                doc = OpenDocumentText(file_path)
                for para in doc.getElementsByType(P):
                    t = para.text.strip()
                    if t: blocks.append(t)
            elif ext == ".ods":
                doc = OpenDocumentSpreadsheet(file_path)
                for table in doc.getElementsByType(Table):
                    for row in table.getElementsByType(TableRow):
                        cells = []
                        for cell in row.getElementsByType(TableCell):
                            cells.append(cell.text.strip())
                        blocks.append(" | ".join(cells))
            elif ext == ".odp":
                doc = OpenDocumentPresentation(file_path)
                for slide in doc.getElementsByType(DrawPage):
                    slide_text = []
                    for para in slide.getElementsByType(P):
                        t = para.text.strip()
                        if t: slide_text.append(t)
                    blocks.append("\n".join(slide_text))
            elif ext == ".pdf":
                with pdfplumber.open(file_path) as pdf:
                    for page in pdf.pages:
                        txt = page.extract_text()
                        if txt:
                            blocks.extend([p.strip() for p in txt.split("\n\n") if p.strip()])
        except Exception as e:
            return [], str(e)
        return blocks, ""

    @staticmethod
    def export_data(file_path, export_rows, export_type="card"):
        """导出：全格式同步导入支持的格式，卡片/笔记/复习记录通用"""
        ext = os.path.splitext(file_path)[1].lower()
        try:
            if ext == ".txt":
                with open(file_path, "w", encoding="utf-8") as f:
                    for row in export_rows:
                        line = "\n".join([str(item) for item in row])
                        f.write(line + "\n" + "-"*40 + "\n")
            elif ext == ".csv":
                with open(file_path, "w", encoding="utf-8", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerows(export_rows)
            elif ext == ".docx":
                doc = docx.Document()
                for row in export_rows:
                    p = doc.add_paragraph()
                    p.add_run("\n".join([str(x) for x in row]))
                    doc.add_paragraph("-"*40)
                doc.save(file_path)
            elif ext == ".xlsx":
                wb = openpyxl.Workbook()
                sheet = wb.active
                for r_idx, row in enumerate(export_rows, start=1):
                    for c_idx, val in enumerate(row, start=1):
                        sheet.cell(row=r_idx, column=c_idx, value=str(val))
                wb.save(file_path)
            elif ext == ".odt":
                doc = OpenDocumentText()
                body = doc.text
                for row in export_rows:
                    p = P()
                    p.addText("\n".join([str(x) for x in row]))
                    body.addElement(p)
                    sep = P()
                    sep.addText("-"*40)
                    body.addElement(sep)
                doc.save(file_path)
            elif ext == ".ods":
                doc = OpenDocumentSpreadsheet()
                table = Table()
                doc.spreadsheet.addElement(table)
                for row_data in export_rows:
                    tr = TableRow()
                    table.addElement(tr)
                    for cell_val in row_data:
                        tc = TableCell()
                        tc.addText(str(cell_val))
                        tr.addElement(tc)
                doc.save(file_path)
            else:
                return False, QApplication.translate("Export", "Unsupported export format")
        except Exception as e:
            return False, str(e)
        return True, ""

# ====================== 主窗口【扩展笔记面板、导出按钮】 ======================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        init_db()
        self.setWindowTitle(QApplication.translate("Main", "Medical Memory Card System"))
        self.resize(1100, 700)
        self.current_deck = None
        self.current_note_id = None
        center = QWidget()
        self.setCentralWidget(center)
        main_layout = QHBoxLayout(center)

        # ========== 左侧总面板（卡组 + 笔记双列表） ==========
        left_total_widget = QWidget()
        left_total_layout = QVBoxLayout(left_total_widget)
        # 卡组区域
        deck_group = QWidget()
        deck_layout = QVBoxLayout(deck_group)
        self.deck_list = QListWidget()
        self.deck_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.deck_list.customContextMenuRequested.connect(self.show_deck_menu)
        new_deck_btn = QPushButton(QApplication.translate("Main", "New Deck"))
        add_card_btn = QPushButton(QApplication.translate("Main", "Add Question Card"))
        import_btn = QPushButton(QApplication.translate("Main", "Batch Import Files"))
        export_btn = QPushButton(QApplication.translate("Main", "Export Data"))
        review_btn = QPushButton(QApplication.translate("Main", "Start Today's Review"))
        deck_layout.addWidget(QLabel(f"<h3>{QApplication.translate('Main','Deck List')}</h3>"))
        deck_layout.addWidget(self.deck_list)
        deck_layout.addWidget(new_deck_btn)
        deck_layout.addWidget(add_card_btn)
        deck_layout.addWidget(import_btn)
        deck_layout.addWidget(export_btn)
        deck_layout.addWidget(review_btn)
        # 笔记区域
        note_group = QWidget()
        note_layout = QVBoxLayout(note_group)
        self.note_list = QListWidget()
        self.note_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.note_list.customContextMenuRequested.connect(self.show_note_menu)
        new_note_btn = QPushButton(QApplication.translate("Note", "Create New Note"))
        del_note_btn = QPushButton(QApplication.translate("Note", "Delete Selected Note"))
        note_layout.addWidget(QLabel(f"<h3>{QApplication.translate('Note','Note List')}</h3>"))
        note_layout.addWidget(self.note_list)
        note_layout.addWidget(new_note_btn)
        note_layout.addWidget(del_note_btn)
        # 拼接左侧两部分
        left_total_layout.addWidget(deck_group, stretch=3)
        left_total_layout.addWidget(note_group, stretch=2)

        # 右侧信息预览面板
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        self.info_label = QLabel(QApplication.translate("Main", "Select deck/note to view details"))
        self.info_label.setWordWrap(True)
        self.note_preview = QTextEdit()
        self.note_preview.setReadOnly(True)
        right_layout.addWidget(self.info_label)
        right_layout.addWidget(QLabel(f"<b>{QApplication.translate('Note','Note Preview')}</b>"))
        right_layout.addWidget(self.note_preview)

        main_layout.addWidget(left_total_widget, 1)
        main_layout.addWidget(right_widget, 2)

        # 绑定卡组信号
        new_deck_btn.clicked.connect(self.create_deck)
        add_card_btn.clicked.connect(self.add_card)
        import_btn.clicked.connect(self.batch_import)
        export_btn.clicked.connect(self.batch_export)
        review_btn.clicked.connect(self.start_review)
        self.deck_list.itemClicked.connect(self.select_deck)
        # 绑定笔记信号
        new_note_btn.clicked.connect(self.create_note)
        del_note_btn.clicked.connect(self.delete_note)
        self.note_list.itemClicked.connect(self.select_note)

        self.refresh_deck()
        self.refresh_note()

    # ====================== 卡组右键菜单 ======================
    def show_deck_menu(self, pos):
        item = self.deck_list.itemAt(pos)
        if not item:
            return
        menu = QMenu(self)
        del_act = QAction(QApplication.translate("Menu", "Delete Deck"), self)
        menu.addAction(del_act)
        act = menu.exec(self.deck_list.mapToGlobal(pos))
        if act == del_act:
            self.delete_deck(item)

    def delete_deck(self, item):
        deck_id = item.data(Qt.ItemDataRole.UserRole)
        deck_name = item.text()
        reply = QMessageBox.question(
            self,
            QApplication.translate("Msg", "Confirm Delete"),
            QApplication.translate("Msg", f"Are you sure to delete deck [{deck_name}]? All cards, notes & review data will be erased!"),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            db_exec("DELETE FROM decks WHERE id=?", (deck_id,))
            if self.current_deck == deck_id:
                self.current_deck = None
                self.info_label.setText(QApplication.translate("Main", "Select a deck on the left to operate"))
            self.refresh_deck()
            self.refresh_note()

    def refresh_deck(self):
        self.deck_list.clear()
        decks = db_query("SELECT id,name FROM decks")
        for did, name in decks:
            item = QListWidgetItem(name)
            item.setData(Qt.ItemDataRole.UserRole, did)
            self.deck_list.addItem(item)

    def create_deck(self):
        name, ok = InputDialog.getText(self, QApplication.translate("Main", "New Deck"),
                                       QApplication.translate("Main", "Input deck name:"))
        if ok and name.strip():
            try:
                db_exec("INSERT INTO decks(name,create_time) VALUES (?,?)", (name.strip(), datetime.now().isoformat()))
            except sqlite3.IntegrityError:
                QMessageBox.warning(self, QApplication.translate("Msg", "Warning"),
                                    QApplication.translate("Msg", "Deck name already exists"))
                return
            self.refresh_deck()

    def select_deck(self, item):
        self.current_deck = item.data(Qt.ItemDataRole.UserRole)
        self.current_note_id = None
        self.note_preview.clear()
        today = datetime.now().isoformat()
        cnt = db_query('''SELECT COUNT(c.id) FROM cards c JOIN review_state rs ON c.id=rs.card_id WHERE c.deck_id=? AND rs.due <= ?''',
                        (self.current_deck, today))[0][0]
        tip = QApplication.translate("Main", "Current Deck ID: {did}\nCards due today: {cnt}")
        self.info_label.setText(tip.format(did=self.current_deck, cnt=cnt))
        self.refresh_note()

    # ====================== 笔记右键菜单（编辑/移动） ======================
    def show_note_menu(self, pos):
        item = self.note_list.itemAt(pos)
        if not item:
            return
        menu = QMenu(self)
        edit_act = QAction(QApplication.translate("Menu", "Edit / Move Note"), self)
        del_act = QAction(QApplication.translate("Menu", "Delete Note"), self)
        menu.addAction(edit_act)
        menu.addAction(del_act)
        act = menu.exec(self.note_list.mapToGlobal(pos))
        nid = item.data(Qt.ItemDataRole.UserRole)
        if act == edit_act:
            self.edit_note(nid)
        elif act == del_act:
            self.delete_note_by_id(nid)

    def refresh_note(self):
        self.note_list.clear()
        if not self.current_deck:
            return
        notes = db_query("SELECT id,title FROM notes WHERE deck_id=?", (self.current_deck,))
        for nid, title in notes:
            item = QListWidgetItem(title)
            item.setData(Qt.ItemDataRole.UserRole, nid)
            self.note_list.addItem(item)

    def select_note(self, item):
        self.current_note_id = item.data(Qt.ItemDataRole.UserRole)
        nid = self.current_note_id
        _, _, title, content, _, _ = db_query("SELECT * FROM notes WHERE id=?", (nid,))[0]
        self.note_preview.setPlainText(f"# {title}\n\n{content}")

    # 新建笔记
    def create_note(self):
        if not self.current_deck:
            QMessageBox.information(self, QApplication.translate("Msg", "Tip"),
                                    QApplication.translate("Msg", "Please select a deck first"))
            return
        all_decks = db_query("SELECT id,name FROM decks")
        dlg = NoteEditDialog(self.current_deck, note_data=None, deck_list=all_decks)
        if dlg.exec():
            self.refresh_note()

    # 编辑/移动笔记
    def edit_note(self, note_id):
        note_data = db_query("SELECT * FROM notes WHERE id=?", (note_id,))[0]
        all_decks = db_query("SELECT id,name FROM decks")
        dlg = NoteEditDialog(self.current_deck, note_data=note_data, deck_list=all_decks)
        if dlg.exec():
            self.refresh_note()

    # 删除笔记（按钮/右键）
    def delete_note(self):
        item = self.note_list.currentItem()
        if not item:
            QMessageBox.information(self, QApplication.translate("Msg", "Tip"),
                                    QApplication.translate("Msg", "Select a note first"))
            return
        nid = item.data(Qt.ItemDataRole.UserRole)
        self.delete_note_by_id(nid)

    def delete_note_by_id(self, nid):
        title = db_query("SELECT title FROM notes WHERE id=?", (nid,))[0][0]
        reply = QMessageBox.question(self, QApplication.translate("Msg", "Confirm Delete"),
                                     QApplication.translate("Msg", f"Delete note [{title}]?"),
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            db_exec("DELETE FROM notes WHERE id=?", (nid,))
            self.current_note_id = None
            self.note_preview.clear()
            self.refresh_note()

    # ====================== 卡片操作 ======================
    def add_card(self):
        if not self.current_deck:
            QMessageBox.information(self, QApplication.translate("Msg", "Tip"),
                                    QApplication.translate("Msg", "Please select a deck first"))
            return
        dlg = CardEditDialog(self.current_deck)
        dlg.exec()
        self.select_deck(self.deck_list.currentItem())

    # 批量导入文件（原有逻辑不变）
    def batch_import(self):
        if not self.current_deck:
            QMessageBox.information(self, QApplication.translate("Msg", "Tip"),
                                    QApplication.translate("Msg", "Please select a deck first"))
            return
        filter_str = (
            "All Supported Files (*.txt *.csv *.tsv *.docx *.xlsx *.pptx *.odt *.ods *.odp *.pdf);;"
            "Text File (*.txt);;CSV/TSV (*.csv *.tsv);;"
            "Word (*.docx *.odt);;Excel (*.xlsx *.ods);;PPT (*.pptx *.odp);;PDF (*.pdf)"
        )
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            QApplication.translate("Import", "Select files to import"),
            "",
            filter_str
        )
        if not file_paths:
            return
        total_add = 0
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for fp in file_paths:
            blocks, err = FileIOHandler.extract_text(fp)
            if err:
                QMessageBox.warning(self, QApplication.translate("Msg", "File Error"),
                                    f"{os.path.basename(fp)}: {err}")
                continue
            # 每个文本块自动生成简易卡片（文本=题干，其余留空）
            for text in blocks:
                if len(text.strip()) < 3:
                    continue
                cid = db_exec('''INSERT INTO cards(deck_id,question,options,answer,analysis,tags,create_time)
                VALUES (?,?,?,?,?,?,?)''', (self.current_deck, text, "", "", "", "import", now))
                card = Card()
                db_exec('''INSERT INTO review_state(card_id,due,stability,difficulty,elapsed_days,scheduled_days,reps,lapses,last_review)
                VALUES (?,?,?,?,?,?,?,?,?)''',
                        (cid, card.due.isoformat(), card.stability, card.difficulty,
                         card.elapsed_days, card.reps, card.lapses, card.lapses, None))
                total_add += 1
        QMessageBox.information(self, QApplication.translate("Import", "Import Complete"),
                                QApplication.translate("Import", f"Successfully imported {total_add} cards!"))
        self.select_deck(self.deck_list.currentItem())

    # 【新增全格式导出】同步导入所有支持格式，导出卡片/笔记/复习记录
    def batch_export(self):
        if not self.current_deck:
            QMessageBox.information(self, QApplication.translate("Msg", "Tip"),
                                    QApplication.translate("Msg", "Please select a deck first"))
            return
        # 选择导出内容类型
        type_choice, ok = InputDialog.getText(self, QApplication.translate("Export", "Export Type"),
                                              QApplication.translate("Export", "Input export type: card / note / review"))
        if not ok or type_choice.strip() not in ["card", "note", "review"]:
            QMessageBox.warning(self, QApplication.translate("Msg", "Warning"),
                                QApplication.translate("Msg", "Please input valid type: card / note / review"))
            return
        # 导出文件筛选（和导入格式完全同步，剔除仅导入的pdf/pptx）
        filter_str = (
            "Export Format (*.txt *.csv *.docx *.xlsx *.odt *.ods);;"
            "Text File (*.txt);;CSV Table (*.csv);;Word Docx (*.docx);;Excel Xlsx (*.xlsx);;ODT Writer (*.odt);;ODS Calc (*.ods)"
        )
        save_path, _ = QFileDialog.getSaveFileName(self, QApplication.translate("Export", "Save Export File"),
                                                   f"deck_{self.current_deck}_{type_choice}_{datetime.now().strftime('%Y%m%d')}", filter_str)
        if not save_path:
            return
        # 读取导出数据
        export_rows = []
        if type_choice == "card":
            export_rows.append(["CardID", "Question", "Options", "Answer", "Analysis", "Tags", "CreateTime"])
            card_data = db_query("SELECT id,question,options,answer,analysis,tags,create_time FROM cards WHERE deck_id=?", (self.current_deck,))
            export_rows.extend(card_data)
        elif type_choice == "note":
            export_rows.append(["NoteID", "Title", "Content", "CreateTime", "UpdateTime"])
            note_data = db_query("SELECT id,title,content,create_time,update_time FROM notes WHERE deck_id=?", (self.current_deck,))
            export_rows.extend(note_data)
        elif type_choice == "review":
            export_rows.append(["CardID", "DueDate", "Stability", "Difficulty", "ElapsedDays", "ScheduledDays", "Reps", "Lapses", "LastReview"])
            review_data = db_query('''SELECT c.id,rs.due,rs.stability,rs.difficulty,rs.elapsed_days,rs.scheduled_days,rs.reps,rs.lapses,rs.last_review
            FROM cards c JOIN review_state rs ON c.id=rs.card_id WHERE c.deck_id=?''', (self.current_deck,))
            export_rows.extend(review_data)
        # 执行导出
        success, err = FileIOHandler.export_data(save_path, export_rows, export_type=type_choice)
        if success:
            QMessageBox.information(self, QApplication.translate("Export", "Export Success"),
                                    QApplication.translate("Export", f"Data saved to {save_path}"))
        else:
            QMessageBox.critical(self, QApplication.translate("Export", "Export Failed"), err)

    def start_review(self):
        if not self.current_deck:
            QMessageBox.information(self, QApplication.translate("Msg", "Tip"),
                                    QApplication.translate("Msg", "Please select a deck first"))
            return
        today = datetime.now().isoformat()
        cids = [row[0] for row in db_query('''SELECT c.id FROM cards c JOIN review_state rs ON c.id=rs.card_id WHERE c.deck_id=? AND rs.due <= ?''',
                                           (self.current_deck, today))]
        if not cids:
            QMessageBox.information(self, QApplication.translate("Msg", "Tip"),
                                    QApplication.translate("Msg", "No cards to review today"))
            return
        for cid in cids:
            dlg = ReviewDialog(cid)
            if dlg.exec() != QDialog.DialogCode.Accepted:
                break

# ====================== 多语言加载入口 ======================
def load_translation(app):
    locale = QLocale.system().name()
    trans = QTranslator()
    return trans

if __name__ == "__main__":
    app = QApplication(sys.argv)
    load_translation(app)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
