# -*- coding: utf-8 -*-
import sys
import sqlite3
from datetime import datetime
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLabel, QTextEdit, QLineEdit, QListWidget,
                             QListWidgetItem, QDialog, QFormLayout, QMessageBox)
from PyQt6.QtCore import Qt
from fsrs import Scheduler, Card, Rating

DB_PATH = "med_fsrs.db"
scheduler = Scheduler()


def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute('''
    CREATE TABLE IF NOT EXISTS decks (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        create_time TEXT
    )
    ''')
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
        FOREIGN KEY(deck_id) REFERENCES decks(id)
    )
    ''')
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
        FOREIGN KEY(card_id) REFERENCES cards(id)
    )
    ''')
    conn.commit()
    conn.close()


def db_query(sql, args=()):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(sql, args)
    res = cur.fetchall()
    conn.close()
    return res


def db_exec(sql, args=()):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(sql, args)
    last_id = cur.lastrowid
    conn.commit()
    conn.close()
    return last_id


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
        ok_btn = QPushButton("确定")
        cancel_btn = QPushButton("取消")
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


class CardEditDialog(QDialog):
    def __init__(self, deck_id, card_data=None):
        super().__init__()
        self.deck_id = deck_id
        self.card_data = card_data
        self.setWindowTitle("编辑题目卡片")
        self.resize(520, 560)
        layout = QFormLayout(self)
        self.question_edit = QTextEdit()
        self.option_edit = QTextEdit()
        self.answer_edit = QLineEdit()
        self.analysis_edit = QTextEdit()
        self.tag_edit = QLineEdit()
        layout.addRow("题干：", self.question_edit)
        layout.addRow("选项：", self.option_edit)
        layout.addRow("答案：", self.answer_edit)
        layout.addRow("解析：", self.analysis_edit)
        layout.addRow("标签(逗号分隔)：", self.tag_edit)
        btn_layout = QHBoxLayout()
        save_btn = QPushButton("保存")
        cancel_btn = QPushButton("取消")
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
            QMessageBox.warning(self, "提示", "题干不能为空")
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


class ReviewDialog(QDialog):
    def __init__(self, card_id):
        super().__init__()
        self.card_id = card_id
        self.setWindowTitle("卡片复习")
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
        show_btn = QPushButton("显示答案和解析")
        rating_layout = QHBoxLayout()
        again_btn = QPushButton("重来")
        hard_btn = QPushButton("困难")
        good_btn = QPushButton("良好")
        easy_btn = QPushButton("简单")
        rating_layout.addWidget(again_btn)
        rating_layout.addWidget(hard_btn)
        rating_layout.addWidget(good_btn)
        rating_layout.addWidget(easy_btn)
        layout.addWidget(QLabel("<b>题干</b>"))
        layout.addWidget(self.q_label)
        layout.addWidget(QLabel("<b>选项</b>"))
        layout.addWidget(self.opt_label)
        layout.addWidget(show_btn)
        layout.addWidget(QLabel("<b>答案、解析</b>"))
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
        self.raw_ans = f"正确答案：{a}"
        self.raw_ana = f"解析：{ana}"
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


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        init_db()
        self.setWindowTitle("医学记忆卡片系统")
        self.resize(900, 650)
        self.current_deck = None
        center = QWidget()
        self.setCentralWidget(center)
        main_layout = QHBoxLayout(center)
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        self.deck_list = QListWidget()
        new_deck_btn = QPushButton("新建卡组")
        add_card_btn = QPushButton("新增题目")
        review_btn = QPushButton("开始今日复习")
        left_layout.addWidget(QLabel("<h3>卡组列表</h3>"))
        left_layout.addWidget(self.deck_list)
        left_layout.addWidget(new_deck_btn)
        left_layout.addWidget(add_card_btn)
        left_layout.addWidget(review_btn)
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        self.info_label = QLabel("请在左侧选中一个卡组进行操作")
        self.info_label.setWordWrap(True)
        right_layout.addWidget(self.info_label)
        main_layout.addWidget(left_widget, 1)
        main_layout.addWidget(right_widget, 2)
        new_deck_btn.clicked.connect(self.create_deck)
        add_card_btn.clicked.connect(self.add_card)
        review_btn.clicked.connect(self.start_review)
        self.deck_list.itemClicked.connect(self.select_deck)
        self.refresh_deck()

    def refresh_deck(self):
        self.deck_list.clear()
        decks = db_query("SELECT id,name FROM decks")
        for did, name in decks:
            item = QListWidgetItem(name)
            item.setData(Qt.ItemDataRole.UserRole, did)
            self.deck_list.addItem(item)

    def create_deck(self):
        name, ok = InputDialog.getText(self, "新建卡组", "输入卡组名称：")
        if ok and name.strip():
            try:
                db_exec("INSERT INTO decks(name,create_time) VALUES (?,?)", (name.strip(), datetime.now().isoformat()))
            except sqlite3.IntegrityError:
                QMessageBox.warning(self, "提示", "卡组名称已存在")
                return
            self.refresh_deck()

    def select_deck(self, item):
        self.current_deck = item.data(Qt.ItemDataRole.UserRole)
        today = datetime.now().isoformat()
        cnt = db_query('''SELECT COUNT(c.id) FROM cards c JOIN review_state rs ON c.id=rs.card_id WHERE c.deck_id=? AND rs.due <= ?''',
                        (self.current_deck, today))[0][0]
        self.info_label.setText(f"当前卡组ID：{self.current_deck}\n今日待复习卡片数量：{cnt}")

    def add_card(self):
        if not self.current_deck:
            QMessageBox.information(self, "提示", "请先选择卡组")
            return
        dlg = CardEditDialog(self.current_deck)
        dlg.exec()
        self.select_deck(self.deck_list.currentItem())

    def start_review(self):
        if not self.current_deck:
            QMessageBox.information(self, "提示", "请先选择卡组")
            return
        today = datetime.now().isoformat()
        cids = [row[0] for row in db_query('''SELECT c.id FROM cards c JOIN review_state rs ON c.id=rs.card_id WHERE c.deck_id=? AND rs.due <= ?''',
                                           (self.current_deck, today))]
        if not cids:
            QMessageBox.information(self, "提示", "今天没有需要复习的卡片")
            return
        for cid in cids:
            dlg = ReviewDialog(cid)
            if dlg.exec() != QDialog.DialogCode.Accepted:
                break


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())