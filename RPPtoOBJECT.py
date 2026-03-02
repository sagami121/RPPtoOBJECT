import os
import sys
import math
import json
import pretty_midi
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QGridLayout, QLabel, QLineEdit, 
                             QPushButton, QCheckBox, QComboBox, QTextEdit, 
                             QFileDialog, QMessageBox, QScrollArea, QFrame, QInputDialog,
                             QSplitter, QTreeWidget, QTreeWidgetItem) 
from PyQt6.QtCore import Qt, QPointF, QRectF
from PyQt6.QtGui import QPalette, QColor, QPainter, QPen, QBrush, QPainterPath

EffDict = {
    "座標": [["X", 0.0], ["Y", 0.0], ["Z", 0.0]],
    "拡大率": [["拡大率", 100.00], ["X", 100.00], ["Y", 100.00]],
}

XDict = {
    "移動無し": "",
    "直線移動": "直線移動",
    "直線移動(時間制御)": "直線移動(時間制御)",
    "補間移動": "補間移動",
    "補間移動(時間制御)": "補間移動(時間制御)",
}

PLAY_SPEED_STEPS = (
    float(100.0 * 0.5),
    float(math.floor(100.0 * (2.0 / 3.0))),  
    float(100.0 * 1.0),
    float(100.0 * 2.0),
    float(100.0 * 4.0),
    float(100.0 * 8.0),
)

FRAME_SPEED_OVERRIDES = {
    72: 66.0,  
    96: 100.0, 
}

def _parse_midi(path):
    midi = pretty_midi.PrettyMIDI(path)

    items = []
    track_names = []

    for idx, inst in enumerate(midi.instruments, 1):
        name = inst.name.strip() if inst.name else f"Track {idx}"
        track_names.append(name)
        seen = set()
        for note in inst.notes:
            if note.end <= note.start:
                continue
            s_tick = int(round(midi.time_to_tick(note.start)))
            e_tick = int(round(midi.time_to_tick(note.end)))
            if e_tick <= s_tick:
                e_tick = s_tick + 1
            key = (s_tick, e_tick)
            if key in seen:
                continue
            seen.add(key)
            s = float(midi.tick_to_time(s_tick))
            e = float(midi.tick_to_time(e_tick))
            items.append({
                "pos": s,
                "length": e - s,
                "track": idx
            })

    if not track_names:
        track_names = ["Track 1"]

    return items, track_names

class BezierCanvas(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(280, 280)
        self.p1 = QPointF(0.0, 0.0)
        self.p2 = QPointF(1.0, 1.0)
        self.active_point = None 
        self.update_bezier_str()

    def update_bezier_str(self):
        x1 = int(round(self.p1.x()))
        y2 = int(round(self.p2.y()))
        self.bezier_str = f"0|0,{x1},{y2},0"

    def get_draw_rect(self):
        side = min(self.width(), self.height()) - 60
        left = (self.width() - side) / 2
        top = (self.height() - side) / 2
        return QRectF(left, top, side, side)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = self.get_draw_rect()
        m_x, m_y, dw, dh = rect.x(), rect.y(), rect.width(), rect.height()
        
        painter.fillRect(self.rect(), QColor(35, 35, 35))
        painter.setPen(QPen(QColor(80, 80, 80), 1))
        painter.drawRect(rect)

        def to_s(p): return QPointF(m_x + p.x() * dw, m_y + (1.0 - p.y()) * dh)
        sp, ep = to_s(QPointF(0, 0)), to_s(QPointF(1, 1))
        cp1, cp2 = to_s(self.p1), to_s(self.p2)

        painter.setPen(QPen(QColor(100, 100, 100, 150), 1, Qt.PenStyle.DashLine))
        painter.drawLine(sp, cp1)
        painter.drawLine(ep, cp2)

        path = QPainterPath()
        path.moveTo(sp)
        path.cubicTo(cp1, cp2, ep)
        painter.setPen(QPen(QColor(100, 255, 150), 3))
        painter.drawPath(path)

        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor(255, 120, 120)) 
        painter.drawEllipse(cp1, 8, 8)
        painter.setBrush(QColor(120, 180, 255)) 
        painter.drawEllipse(cp2, 8, 8)
        
        painter.setBrush(Qt.GlobalColor.white)
        painter.drawEllipse(sp, 4, 4)
        painter.drawEllipse(ep, 4, 4)

    def mousePressEvent(self, event):
        rect = self.get_draw_rect()
        def to_s(p): return QPointF(rect.x() + p.x() * rect.width(), rect.y() + (1.0 - p.y()) * rect.height())
        
        p1_s, p2_s = to_s(self.p1), to_s(self.p2)
        if (event.position() - p1_s).manhattanLength() < 20:
            self.active_point = "p1"
        elif (event.position() - p2_s).manhattanLength() < 20:
            self.active_point = "p2"
        else:
            self.active_point = None

    def mouseMoveEvent(self, event):
        if not self.active_point: return
        
        rect = self.get_draw_rect()
        nx = max(0.0, min(1.0, (event.position().x() - rect.x()) / rect.width()))
        ny = max(0.0, min(1.0, 1.0 - ((event.position().y() - rect.y()) / rect.height())))
        
        if self.active_point == "p1":
            self.p1 = QPointF(nx, ny) 
        else:
            self.p2 = QPointF(nx, ny)
        
        self.update()
        self.update_bezier_str()

    def mouseReleaseEvent(self, event):
        self.active_point = None

class RPPtoObjectApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.lang_dir = os.path.join(os.path.dirname(__file__), "language")
        self.settings_path = os.path.join(os.path.dirname(__file__), "settings.json")
        self.lang_code = "ja"
        self.play_speed_steps = tuple(PLAY_SPEED_STEPS)
        self.i18n = {}
        self.obj_terms = {}
        self.obj_tokens = {}
        self.settings = self._load_settings()
        self._load_language(self.lang_code)
        self.setWindowTitle(self._tr("window_title"))
        self.setMinimumSize(1100, 800)
        self.file_rows = []
        self.form_rows = []
        self.added_effects_data = []
        self.midi_cache = {"path": None, "mtime": None, "items": None, "track_names": None}
        self.init_ui()
        self._apply_settings_to_ui()

    def _load_settings(self):
        defaults = {
            "version": "1.0beta3",
            "language": "ja",
            "fps": "60",
            "scene_no": "1",
            "base_len_sec": "1.0",
            "play_speed_steps": list(PLAY_SPEED_STEPS)
        }
        loaded = {}
        try:
            with open(self.settings_path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
        except Exception:
            loaded = {}

        settings = dict(defaults)
        if isinstance(loaded, dict):
            settings.update(loaded)

        lang = settings.get("language", "ja")
        self.lang_code = lang if lang in ("ja", "en") else "ja"

        raw_steps = settings.get("play_speed_steps", PLAY_SPEED_STEPS)
        parsed_steps = []
        if isinstance(raw_steps, list):
            for v in raw_steps:
                try:
                    fv = float(v)
                    if fv > 0:
                        parsed_steps.append(fv)
                except Exception:
                    continue
        self.play_speed_steps = tuple(parsed_steps) if parsed_steps else tuple(PLAY_SPEED_STEPS)
        settings["play_speed_steps"] = list(self.play_speed_steps)
        settings["language"] = self.lang_code
        return settings

    def _apply_settings_to_ui(self):
        self.fps_in.setText(str(self.settings.get("fps", "60")))
        self.scene_in.setText(str(self.settings.get("scene_no", "1")))
        self.base_len.setText(str(self.settings.get("base_len_sec", "1.0")))
        idx = self.lang_combo.findData(self.lang_code)
        if idx >= 0 and self.lang_combo.currentIndex() != idx:
            self.lang_combo.setCurrentIndex(idx)

    def _save_settings(self):
        settings = {
            "version": "1.0beta3",
            "language": self.lang_combo.currentData() or self.lang_code,
            "fps": self.fps_in.text().strip() or "60",
            "scene_no": self.scene_in.text().strip() or "1",
            "base_len_sec": self.base_len.text().strip() or "1.0",
            "play_speed_steps": [float(v) for v in self.play_speed_steps]
        }
        try:
            with open(self.settings_path, "w", encoding="utf-8") as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _load_language(self, code):
        base = {}
        ja_path = os.path.join(self.lang_dir, "ja.json")
        lang_path = os.path.join(self.lang_dir, f"{code}.json")
        try:
            with open(ja_path, "r", encoding="utf-8") as f:
                base = json.load(f)
        except Exception:
            base = {}
        if code != "ja":
            try:
                with open(lang_path, "r", encoding="utf-8") as f:
                    base.update(json.load(f))
            except Exception:
                pass
        self.lang_code = code
        self.i18n = base
        self._load_object_language("ja")

    def _load_object_language(self, code):
        self.obj_terms = {}
        self.obj_tokens = {}
        obj_path = os.path.join(self.lang_dir, "object.json")
        try:
            with open(obj_path, "r", encoding="utf-8") as f:
                obj_i18n = json.load(f)
        except Exception:
            obj_i18n = {}

        ja_pack = obj_i18n.get("ja", {})
        self.obj_terms = dict(ja_pack.get("terms", {}))
        self.obj_tokens = dict(ja_pack.get("tokens", {}))

        if code != "ja":
            lang_pack = obj_i18n.get(code, {})
            self.obj_terms.update(lang_pack.get("terms", {}))
            self.obj_tokens.update(lang_pack.get("tokens", {}))

    def _tr(self, key, **kwargs):
        text = self.i18n.get(key, key)
        if kwargs:
            try:
                return text.format(**kwargs)
            except Exception:
                return text
        return text

    def _obj(self, key):
        return self.obj_terms.get(key, key)

    def _obj_text(self, text):
        return self.obj_tokens.get(text, text)

    def _ui_token(self, text):
        mapping = {
            "座標": "ui_eff_coordinate",
            "拡大率": "ui_eff_scale",
            "X": "ui_param_x",
            "Y": "ui_param_y",
            "Z": "ui_param_z",
            "移動無し": "ui_motion_none",
            "直線移動": "ui_motion_linear",
            "直線移動(時間制御)": "ui_motion_linear_time",
            "補間移動": "ui_motion_interp",
            "補間移動(時間制御)": "ui_motion_interp_time"
        }
        key = mapping.get(text)
        return self._tr(key) if key else text

    def _translate_motion_name(self, method_name):
        mapping = {
            "": "motion.none",
            "直線移動": "motion.linear",
            "直線移動(時間制御)": "motion.linear_time",
            "補間移動": "motion.interpolate",
            "補間移動(時間制御)": "motion.interpolate_time"
        }
        return self._obj(mapping.get(method_name, "motion.none"))

    def _fill_motion_combo(self, combo):
        current_key = combo.currentData()
        combo.clear()
        for key in XDict.keys():
            combo.addItem(self._ui_token(key), key)
        if current_key is None:
            current_key = "移動無し"
        idx = combo.findData(current_key)
        if idx >= 0:
            combo.setCurrentIndex(idx)

    def refresh_effect_combo(self):
        current_key = self.eff_combo.currentData()
        self.eff_combo.blockSignals(True)
        self.eff_combo.clear()
        for key in EffDict.keys():
            self.eff_combo.addItem(self._ui_token(key), key)
        if current_key is not None:
            idx = self.eff_combo.findData(current_key)
            if idx >= 0:
                self.eff_combo.setCurrentIndex(idx)
        self.eff_combo.blockSignals(False)

    def init_ui(self):
        cw = QWidget()
        self.setCentralWidget(cw)
        main_layout = QVBoxLayout(cw)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        left_widget = QWidget()
        left = QVBoxLayout(left_widget)

        lang_row = QHBoxLayout()
        self.lang_label = QLabel(self._tr("language"))
        self.lang_combo = QComboBox()
        self.lang_combo.addItem(self._tr("lang_japanese"), "ja")
        self.lang_combo.addItem(self._tr("lang_english"), "en")
        self.lang_combo.setCurrentIndex(0 if self.lang_code == "ja" else 1)
        self.lang_combo.currentIndexChanged.connect(self.on_language_changed)
        lang_row.addWidget(self.lang_label)
        lang_row.addWidget(self.lang_combo)
        lang_row.addStretch()
        left.addLayout(lang_row)
        
        self.rpp_path = self.add_file_row(left, self._tr("file_rpp_midi"), self.select_rpp)
        self.exo_path = self.add_file_row(left, self._tr("file_object"), self.save_exo)
        self.src_path = self.add_file_row(left, self._tr("file_source"), self.select_src)
        
        self.lbl_basic = QLabel(f"<b>{self._tr('section_basic')}</b>")
        left.addWidget(self.lbl_basic)
        opt_grid = QGridLayout()
        self.cb_flip_h = QCheckBox(self._tr("flip_h"))
        self.cb_flip_v = QCheckBox(self._tr("flip_v"))
        self.cb_loop = QCheckBox(self._tr("loop_play"))
        self.cb_no_gap = QCheckBox(self._tr("no_gap"))
        self.cb_as_scene = QCheckBox(self._tr("as_scene"))
        self.cb_auto_speed = QCheckBox(self._tr("auto_speed"))
        self.cb_redzone = QCheckBox(self._tr("redzone_mode"))
        self.cb_redzone.setStyleSheet("color: #FF6666; font-weight: bold;")

        self.base_len = QLineEdit("1.0")
        self.base_len.setFixedWidth(40)

        opt_grid.addWidget(self.cb_flip_h, 0, 0)
        opt_grid.addWidget(self.cb_flip_v, 0, 1)
        opt_grid.addWidget(self.cb_loop, 1, 0)
        opt_grid.addWidget(self.cb_no_gap, 1, 1)
        opt_grid.addWidget(self.cb_as_scene, 2, 0)
        opt_grid.addWidget(self.cb_redzone, 3, 0)
        
        speed_h = QHBoxLayout()
        speed_h.addWidget(self.cb_auto_speed)
        self.lbl_base = QLabel(self._tr("base_label"))
        speed_h.addWidget(self.lbl_base)
        speed_h.addWidget(self.base_len)
        self.lbl_seconds = QLabel(self._tr("seconds"))
        speed_h.addWidget(self.lbl_seconds)
        speed_h.addStretch()
        opt_grid.addLayout(speed_h, 2, 1)
        left.addLayout(opt_grid)

        tc_group = QFrame()
        tc_group.setStyleSheet("QFrame { background-color: #333; border: 1px solid #555; border-radius: 6px; }")
        tc_lay = QVBoxLayout(tc_group)
        
        tc_top_h = QHBoxLayout()
        self.cb_time_ctrl = QCheckBox(self._tr("time_control"))
        self.cb_apply_easing = QCheckBox(self._tr("apply_easing"))
        self.cb_apply_easing.setChecked(False) 
        tc_top_h.addWidget(self.cb_time_ctrl)
        tc_top_h.addWidget(self.cb_apply_easing)
        tc_top_h.addStretch()

        step_h = QHBoxLayout()
        step_h.setContentsMargins(5, 0, 5, 5)
        self.lbl_step = QLabel(self._tr("step_frame"))
        step_h.addWidget(self.lbl_step)
        self.tc_step = QLineEdit("1")
        self.tc_step.setFixedWidth(45)
        step_h.addWidget(self.tc_step)
        step_h.addStretch()
        
        tc_lay.addLayout(tc_top_h)
        tc_lay.addLayout(step_h)
        left.addWidget(tc_group)

        form_grid = QGridLayout()
        self.fps_in = self.add_form_row(form_grid, self._tr("fps"), "60", 0)
        self.scene_in = self.add_form_row(form_grid, self._tr("scene_no"), "1", 1)
        left.addLayout(form_grid)

        self.lbl_track = QLabel(self._tr("track_select"))
        left.addWidget(self.lbl_track)
        self.track_tree = QTreeWidget()
        self.track_tree.setHeaderHidden(True)
        self.track_tree.setFixedHeight(120)
        self.track_tree.setStyleSheet("background-color: #2a2a2a; color: white;")

        self.root_item = QTreeWidgetItem(self.track_tree, [self._tr("all_tracks")])
        self.root_item.setCheckState(0, Qt.CheckState.Checked)
        self.track_tree.itemChanged.connect(self.on_tree_changed)
        left.addWidget(self.track_tree)

        self.lbl_script = QLabel(self._tr("script_control"))
        left.addWidget(self.lbl_script)
        self.script_edit = QTextEdit()
        left.addWidget(self.script_edit)
        
        self.run_btn = QPushButton(self._tr("run_output"))
        self.run_btn.setStyleSheet("height: 50px; background-color: #0D47A1; color: white; font-weight: bold; border-radius: 4px;")
        self.run_btn.clicked.connect(self.run_process)
        left.addWidget(self.run_btn)

        right_widget = QWidget()
        right = QVBoxLayout(right_widget)
        
        self.lbl_easing = QLabel(f"<b>{self._tr('section_easing')}</b>")
        right.addWidget(self.lbl_easing)
        self.bezier_ui = BezierCanvas() 
        right.addWidget(self.bezier_ui)
        
        self.lbl_effect = QLabel(f"<b>{self._tr('section_effect')}</b>")
        right.addWidget(self.lbl_effect)
        h_eff = QHBoxLayout()
        self.eff_combo = QComboBox()
        self.refresh_effect_combo()
        self.add_btn = QPushButton(self._tr("add_effect"))
        self.add_btn.setFixedWidth(80)
        self.add_btn.setStyleSheet("background-color: #2E7D32; color: white; font-weight: bold;")
        self.add_btn.clicked.connect(self.add_eff_ui)
        h_eff.addWidget(self.eff_combo)
        h_eff.addWidget(self.add_btn)
        right.addLayout(h_eff)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setStyleSheet("QScrollArea { border: none; background-color: #252525; }")
        self.eff_cont = QWidget()
        self.eff_cont.setStyleSheet("background-color: #252525;")
        self.eff_list_layout = QVBoxLayout(self.eff_cont)
        self.eff_list_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.scroll.setWidget(self.eff_cont)
        right.addWidget(self.scroll)

        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 3) 
        splitter.setStretchFactor(1, 2) 
        splitter.setHandleWidth(4)
        splitter.setStyleSheet("QSplitter::handle { background-color: #555; }") 
        main_layout.addWidget(splitter)

    def on_tree_changed(self, item, column):
        if item == self.root_item:
            state = item.checkState(0)
            for i in range(self.root_item.childCount()):
                self.root_item.child(i).setCheckState(0, state)

    def on_language_changed(self, _index):
        code = self.lang_combo.currentData()
        if code not in ("ja", "en"):
            return
        self._load_language(code)
        self.apply_language()

    def apply_language(self):
        self.setWindowTitle(self._tr("window_title"))
        self.lang_label.setText(self._tr("language"))

        file_labels = [self._tr("file_rpp_midi"), self._tr("file_object"), self._tr("file_source")]
        for i, row in enumerate(self.file_rows):
            row["label"].setText(file_labels[i])
            row["button"].setText(self._tr("select_btn"))

        form_labels = [self._tr("fps"), self._tr("scene_no")]
        for i, row in enumerate(self.form_rows):
            row["label"].setText(form_labels[i])

        self.lbl_basic.setText(f"<b>{self._tr('section_basic')}</b>")
        self.cb_flip_h.setText(self._tr("flip_h"))
        self.cb_flip_v.setText(self._tr("flip_v"))
        self.cb_loop.setText(self._tr("loop_play"))
        self.cb_no_gap.setText(self._tr("no_gap"))
        self.cb_as_scene.setText(self._tr("as_scene"))
        self.cb_auto_speed.setText(self._tr("auto_speed"))
        self.cb_redzone.setText(self._tr("redzone_mode"))
        self.lbl_base.setText(self._tr("base_label"))
        self.lbl_seconds.setText(self._tr("seconds"))
        self.cb_time_ctrl.setText(self._tr("time_control"))
        self.cb_apply_easing.setText(self._tr("apply_easing"))
        self.lbl_step.setText(self._tr("step_frame"))
        self.lbl_track.setText(self._tr("track_select"))
        self.root_item.setText(0, self._tr("all_tracks"))
        self.lbl_script.setText(self._tr("script_control"))
        self.run_btn.setText(self._tr("run_output"))
        self.lbl_easing.setText(f"<b>{self._tr('section_easing')}</b>")
        self.lbl_effect.setText(f"<b>{self._tr('section_effect')}</b>")
        self.add_btn.setText(self._tr("add_effect"))
        self.refresh_effect_combo()

        self.lang_combo.blockSignals(True)
        self.lang_combo.setItemText(0, self._tr("lang_japanese"))
        self.lang_combo.setItemText(1, self._tr("lang_english"))
        self.lang_combo.blockSignals(False)

    def add_file_row(self, layout, label, func):
        h = QHBoxLayout()
        lbl = QLabel(label)
        e = QLineEdit()
        b = QPushButton(self._tr("select_btn"))
        b.setFixedWidth(60)
        b.clicked.connect(func)
        h.addWidget(lbl)
        h.addWidget(e)
        h.addWidget(b)
        layout.addLayout(h)
        self.file_rows.append({"label": lbl, "button": b})
        return e

    def add_form_row(self, grid, label, default, row):
        lbl = QLabel(label)
        e = QLineEdit(default)
        grid.addWidget(lbl, row, 0)
        grid.addWidget(e, row, 1)
        self.form_rows.append({"label": lbl})
        return e

    def select_rpp(self):
        p, _ = QFileDialog.getOpenFileName(self, self._tr("dialog_select_rpp_midi"), "", "RPP/MIDI (*.rpp *.mid *.midi)")
        if p: self.rpp_path.setText(p); self.load_tracks(p)

    def save_exo(self):
        p, _ = QFileDialog.getSaveFileName(self, self._tr("dialog_save"), "", "*.object")
        if p: self.exo_path.setText(p)

    def select_src(self):
        p, _ = QFileDialog.getOpenFileName(self, self._tr("dialog_select_source"))
        if p: self.src_path.setText(p)

    def load_tracks(self, path):
        self.root_item.takeChildren()
        self.midi_cache = {"path": None, "mtime": None, "items": None, "track_names": None}

        ext = os.path.splitext(path)[1].lower()
        if ext in (".mid", ".midi"):
            try:
                items, track_names = _parse_midi(path)
                mtime = os.path.getmtime(path)
                self.midi_cache = {"path": path, "mtime": mtime, "items": items, "track_names": track_names}
                for idx, name in enumerate(track_names, 1):
                    child = QTreeWidgetItem(self.root_item, [f"{idx:02} MIDI \"{name}\""])
                    child.setData(0, Qt.ItemDataRole.UserRole, idx)
                    child.setCheckState(0, Qt.CheckState.Checked)
                self.root_item.setExpanded(True)
            except Exception as e:
                QMessageBox.critical(self, self._tr("midi_load_error_en"), str(e))
            return

        try:
            with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                idx = 0
                for line in f:
                    if "<TRACK" in line:
                        idx += 1; l2 = next(f)
                        name = l2.replace("NAME ", "").strip() if "NAME" in l2 else f"Track {idx}"
                        symbol = "┣" if idx > 0 else "┗"
                        child = QTreeWidgetItem(self.root_item, [f"{idx:02} {symbol} \"{name}\""])
                        child.setData(0, Qt.ItemDataRole.UserRole, idx) 
                        child.setCheckState(0, Qt.CheckState.Checked)
            self.root_item.setExpanded(True)
        except:
            pass

    def add_eff_ui(self):
        name = self.eff_combo.currentData() or self.eff_combo.currentText()
        frame = QFrame(); frame.setObjectName("EffectCard")
        frame.setStyleSheet("""
            QFrame#EffectCard { background-color: #383838; border: 1px solid #555; border-radius: 8px; margin-bottom: 8px; }
            QLabel { color: #DDD; border: none; }
            QPushButton#DelBtn { background-color: #C62828; color: white; border-radius: 11px; font-weight: bold; border: none; }
        """)
        outer_lay = QVBoxLayout(frame); outer_lay.setContentsMargins(0, 0, 0, 0); outer_lay.setSpacing(0)
        header = QWidget(); header.setFixedHeight(30); header.setStyleSheet("background-color: #4A4A4A; border-top-left-radius: 7px; border-top-right-radius: 7px;")
        header_lay = QHBoxLayout(header); header_lay.setContentsMargins(10, 0, 10, 0)
        lbl = QLabel(f"<b>{self._ui_token(name)}</b>"); del_btn = QPushButton("×"); del_btn.setObjectName("DelBtn"); del_btn.setFixedSize(22, 22)
        header_lay.addWidget(lbl); header_lay.addStretch(); header_lay.addWidget(del_btn); outer_lay.addWidget(header)
        
        content = QWidget(); lay = QGridLayout(content); lay.setContentsMargins(10, 8, 10, 10); lay.setSpacing(8)
        info = {"name": name, "params": [], "frame": frame}
        del_btn.clicked.connect(lambda: self.remove_eff(info))
        
        row = 0
        for pdef in EffDict[name]:
            pname = pdef[0]
            if len(pdef) > 2 and pdef[2] == -1:
                cb = QCheckBox(self._ui_token(pname)); cb.setChecked(bool(pdef[1])); lay.addWidget(cb, row, 0, 1, 5)
                info["params"].append({"type": "cb", "name": pname, "obj": cb})
            else:
                lay.addWidget(QLabel(self._ui_token(pname)), row, 0)
                s_ed = QLineEdit(str(pdef[1])); s_ed.setFixedWidth(50); lay.addWidget(s_ed, row, 2)
                m_cb = QComboBox(); self._fill_motion_combo(m_cb); m_cb.setMinimumWidth(100); lay.addWidget(m_cb, row, 3)
                e_ed = QLineEdit(str(pdef[1])); e_ed.setFixedWidth(50); lay.addWidget(e_ed, row, 4)
                info["params"].append({"type": "motion", "name": pname, "start": s_ed, "end": e_ed, "method": m_cb})
            row += 1
        outer_lay.addWidget(content); self.eff_list_layout.addWidget(frame); self.added_effects_data.append(info)

    def remove_eff(self, info):
        self.added_effects_data.remove(info); info["frame"].deleteLater()

    def run_process(self):
        try:
            rpp_p = self.rpp_path.text().strip()
            exo_p = self.exo_path.text().strip()
            src_p = self.src_path.text().strip()

            if not rpp_p:
                QMessageBox.warning(self, self._tr("error_input"), self._tr("msg_need_rpp_midi"))
                return
            if not os.path.exists(rpp_p):
                QMessageBox.critical(self, self._tr("error_file"), self._tr("msg_rpp_midi_not_found"))
                return
            if not exo_p:
                QMessageBox.warning(self, self._tr("error_input"), self._tr("msg_need_output"))
                return

            try:
                fps_text = self.fps_in.text().strip() or "60"
                fps = float(fps_text)
                if fps <= 0:
                    raise ValueError
            except:
                QMessageBox.critical(self, self._tr("error_number"), self._tr("msg_fps_gt0"))
                return

            try:
                base_s = float(self.base_len.text() or 1.0)
                if base_s <= 0:
                    raise ValueError
            except:
                QMessageBox.critical(self, self._tr("error_number"), self._tr("msg_base_gt0"))
                return

            bez_str = self.bezier_ui.bezier_str
            is_redzone = self.cb_redzone.isChecked()
            flip_h = self.cb_flip_h.isChecked()
            flip_v = self.cb_flip_v.isChecked()
            effect_scene = self._obj("effect.scene")
            effect_video = self._obj("effect.video_file")
            effect_standard_draw = self._obj("effect.standard_draw")
            effect_flip = self._obj("effect.flip")
            effect_clip = self._obj("effect.clip")
            effect_time_ctrl_obj = "時間制御(オブジェクト)"
            param_play_pos = self._obj("param.play_pos")
            param_play_speed = self._obj("param.play_speed")
            param_scene = self._obj("param.scene")
            param_file = self._obj("param.file")
            param_audio_on = self._obj("param.audio_on")
            param_x = self._obj("param.x")
            param_y = self._obj("param.y")
            param_z = self._obj("param.z")
            param_opacity = self._obj("param.opacity")
            param_flip_ud = self._obj("param.flip_ud")
            param_flip_lr = self._obj("param.flip_lr")
            param_flip_luma = self._obj("param.flip_luma")
            param_flip_hue = self._obj("param.flip_hue")
            param_flip_alpha = self._obj("param.flip_alpha")
            param_clip_top = self._obj("param.clip_top")
            param_clip_bottom = self._obj("param.clip_bottom")
            param_clip_left = self._obj("param.clip_left")
            param_clip_right = self._obj("param.clip_right")
            param_position = "位置"
            param_frame_step = "コマ落ち"
            param_target_layer_count = "対象レイヤー数"
            motion_linear_time = "直線移動(時間制御)"

            active_tracks = []
            for i in range(self.root_item.childCount()):
                child = self.root_item.child(i)
                if child.checkState(0) == Qt.CheckState.Checked:
                    active_tracks.append(child.data(0, Qt.ItemDataRole.UserRole))

            if not active_tracks:
                QMessageBox.warning(self, self._tr("error_track_unselected"), self._tr("msg_need_track"))
                return

            ext = os.path.splitext(rpp_p)[1].lower()
            is_midi_input = ext in (".mid", ".midi")
            all_objs = []

            if is_midi_input:
                try:
                    mtime = os.path.getmtime(rpp_p)
                    if (self.midi_cache["path"] == rpp_p and
                            self.midi_cache["mtime"] == mtime and
                            self.midi_cache["items"] is not None):
                        all_objs = list(self.midi_cache["items"])
                    else:
                        all_objs, track_names = _parse_midi(rpp_p)
                        self.midi_cache = {"path": rpp_p, "mtime": mtime, "items": all_objs, "track_names": track_names}
                except Exception as e:
                    QMessageBox.critical(self, self._tr("error_midi_read"), str(e))
                    return

                all_objs = [o for o in all_objs if o["track"] in active_tracks]
            else:
                try:
                    with open(rpp_p, 'r', encoding='utf-8', errors='ignore') as f:
                        content = f.read()
                except Exception as e:
                    QMessageBox.critical(self, self._tr("error_read"), str(e))
                    return

                items_raw = content.split("<ITEM")
                if len(items_raw) <= 1:
                    QMessageBox.warning(self, self._tr("error_parse"), self._tr("msg_item_not_found"))
                    return

                track_count = 0

                for i in range(1, len(items_raw)):
                    track_count += items_raw[i-1].count("<TRACK")
                    if track_count not in active_tracks:
                        continue

                    block = items_raw[i]
                    pos, length = None, None

                    for line in block.split("\n"):
                        ls = line.strip()
                        if ls.startswith("POSITION"):
                            try:
                                pos = float(ls.split()[1])
                            except:
                                pass
                        if ls.startswith("LENGTH"):
                            try:
                                length = float(ls.split()[1])
                            except:
                                pass

                    if pos is None or length is None:
                        continue
                    if length <= 0:
                        continue

                    all_objs.append({
                        "pos": pos,
                        "length": length,
                        "track": track_count
                    })

            if not all_objs:
                QMessageBox.warning(self, self._tr("error_no_result"), self._tr("msg_valid_item_not_found"))
                return

            objs_to_process = sorted(all_objs, key=lambda x: (x["track"], x["pos"]))
            output = []
            last_end_frames = {}
            track_item_counts = {}
            emitted_frame_keys = set()
            total_obj_idx = 0

            for o in objs_to_process:
                try:
                    t_idx = o["track"]

                    if t_idx not in track_item_counts:
                        track_item_counts[t_idx] = 0
                    track_item_counts[t_idx] += 1
                    curr_cnt = track_item_counts[t_idx]

                    start_f = round(o["pos"] * fps)

                    if self.cb_no_gap.isChecked() and t_idx in last_end_frames:
                        if abs(start_f - (last_end_frames[t_idx] + 1)) < 5:
                            start_f = last_end_frames[t_idx] + 1

                    end_f = start_f + round(o["length"] * fps) - 1
                    if end_f < start_f:
                        continue
                    if is_midi_input:
                        frame_key = (t_idx, start_f, end_f)
                        if frame_key in emitted_frame_keys:
                            continue
                        emitted_frame_keys.add(frame_key)

                    last_end_frames[t_idx] = end_f

                    speed = 100.0
                    if self.cb_auto_speed.isChecked() and o["length"] > 0:
                        speed = base_s / o["length"] * 100.0
                    duration_frames = end_f - start_f + 1
                    if duration_frames in FRAME_SPEED_OVERRIDES:
                        speed = FRAME_SPEED_OVERRIDES[duration_frames]
                    else:
                        speed = min(self.play_speed_steps, key=lambda v: abs(v - speed))

                    main_layer = t_idx * 2 if self.cb_time_ctrl.isChecked() else t_idx

                    obj_x = 0.0
                    if is_redzone:
                        if t_idx == 1:
                            obj_x = -480.0
                        elif t_idx == 2:
                            obj_x = 480.0

                    output.append(f"[{total_obj_idx}]\nframe={start_f},{end_f}\nlayer={main_layer}\n")

                    if self.cb_as_scene.isChecked():
                        output.append(
                            f"[{total_obj_idx}.0]\n"
                            f"effect.name={effect_scene}\n"
                            f"{param_play_pos}=0.000\n"
                            f"{param_play_speed}={speed:.2f}\n"
                            f"{param_scene}={self.scene_in.text()}\n"
                        )
                    else:
                        output.append(
                            f"[{total_obj_idx}.0]\n"
                            f"effect.name={effect_video}\n"
                            f"{param_file}={src_p}\n"
                            f"{param_play_pos}=0.000\n"
                            f"{param_play_speed}={speed:.2f}\n"
                            f"{param_audio_on}=1\n"
                        )

                    output.append(
                        f"[{total_obj_idx}.1]\n"
                        f"effect.name={effect_standard_draw}\n"
                        f"{param_x}={obj_x:.2f}\n{param_y}=0.00\n{param_z}=0.00\n{param_opacity}=0.00\n"
                    )

                    p_idx = 2

                    if flip_h or flip_v:
                        ud_val = 0
                        lr_val = 0
                        if flip_h and flip_v:
                            phase = (curr_cnt - 1) % 4
                            if phase == 1:
                                lr_val = 1
                            elif phase == 2:
                                ud_val = 1
                            elif phase == 3:
                                ud_val, lr_val = 1, 1
                        else:
                            flip_on = (curr_cnt % 2 == 0)
                            ud_val = 1 if (flip_v and flip_on) else 0
                            lr_val = 1 if (flip_h and flip_on) else 0

                        if ud_val or lr_val:
                            output.append(
                                f"[{total_obj_idx}.{p_idx}]\n"
                                f"effect.name={effect_flip}\n"
                                f"{param_flip_ud}={ud_val}\n"
                                f"{param_flip_lr}={lr_val}\n"
                                f"{param_flip_luma}=0\n{param_flip_hue}=0\n{param_flip_alpha}=0\n"
                            )
                            p_idx += 1

                    if is_redzone and (t_idx == 1 or t_idx == 2):
                        output.append(
                            f"[{total_obj_idx}.{p_idx}]\n"
                            f"effect.name={effect_clip}\n"
                            f"{param_clip_top}=0\n{param_clip_bottom}=0\n{param_clip_left}=480\n{param_clip_right}=480\n"
                        )
                        p_idx += 1

                    for eff in self.added_effects_data:
                        output.append(f"[{total_obj_idx}.{p_idx}]\neffect.name={self._obj_text(eff['name'])}\n")
                        for p in eff["params"]:
                            out_pname = self._obj_text(p["name"])
                            if p["type"] == "cb":
                                val = int(p["obj"].isChecked())
                            else:
                                try:
                                    float(p["start"].text())
                                    float(p["end"].text())
                                except:
                                    QMessageBox.warning(self, self._tr("error_number"), self._tr("msg_invalid_value", name=p["name"]))
                                    return

                                motion_name = self._translate_motion_name(XDict[p["method"].currentText()])
                                val = f"{p['start'].text()},{p['end'].text()},{motion_name},{bez_str}"

                            output.append(f"{out_pname}={val}\n")
                        p_idx += 1

                    output.append("\n")
                    total_obj_idx += 1

                    if self.cb_time_ctrl.isChecked():
                        if is_redzone:
                            if curr_cnt % 2 == 0:
                                s_val, e_val = "100.000", "0.000"
                            else:
                                s_val, e_val = "0.000", "100.000"
                        else:
                            s_val, e_val = "0.000", "100.000"

                        t_easing = bez_str if self.cb_apply_easing.isChecked() else "0"

                        output.append(
                            f"[{total_obj_idx}]\n"
                            f"frame={start_f},{end_f}\n"
                            f"layer={main_layer-1}\n"
                        )

                        output.append(
                            f"[{total_obj_idx}.0]\n"
                            f"effect.name={effect_time_ctrl_obj}\n"
                            f"{param_position}={s_val},{e_val},{motion_linear_time},{t_easing}\n"
                            f"{param_frame_step}={self.tc_step.text()}\n"
                            f"{param_target_layer_count}=1\n\n"
                        )

                        total_obj_idx += 1

                except Exception as inner_e:
                    print("オブジェクト処理エラー:", inner_e)
                    continue

            try:
                with open(exo_p, 'w', encoding='utf-8') as f:
                    f.write("".join(output))
            except Exception as e:
                QMessageBox.critical(self, self._tr("error_save"), str(e))
                return

            QMessageBox.information(
                self,
                self._tr("done"),
                self._tr("msg_done", prefix=(self._tr("redzone_prefix") if is_redzone else ""))
            )

        except Exception as e:
            QMessageBox.critical(self, self._tr("error_fatal"), str(e))

    def closeEvent(self, event):
        self._save_settings()
        super().closeEvent(event)

if __name__ == "__main__":
    app = QApplication(sys.argv); app.setStyle("Fusion")
    dark_p = QPalette()
    dark_p.setColor(QPalette.ColorRole.Window, QColor(45, 45, 45))
    dark_p.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
    dark_p.setColor(QPalette.ColorRole.Base, QColor(30, 30, 30))
    dark_p.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
    dark_p.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
    dark_p.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
    app.setPalette(dark_p); win = RPPtoObjectApp(); win.show(); sys.exit(app.exec())
