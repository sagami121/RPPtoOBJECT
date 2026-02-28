import os
import sys
import math
import json
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
        self.setWindowTitle("RPPtoOBJECT v1.0 Beta2")
        self.setMinimumSize(1100, 800)
        self.added_effects_data = []
        self.init_ui()

    def init_ui(self):
        cw = QWidget()
        self.setCentralWidget(cw)
        main_layout = QVBoxLayout(cw)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        left_widget = QWidget()
        left = QVBoxLayout(left_widget)
        
        self.rpp_path = self.add_file_row(left, "RPP:", self.select_rpp)
        self.exo_path = self.add_file_row(left, ".object:", self.save_exo)
        self.src_path = self.add_file_row(left, "素材:", self.select_src)
        
        left.addWidget(QLabel("<b>基本設定</b>"))
        opt_grid = QGridLayout()
        self.cb_flip_h = QCheckBox("左右反転")
        self.cb_flip_v = QCheckBox("上下反転")
        self.cb_loop = QCheckBox("ループ再生")
        self.cb_no_gap = QCheckBox("隙間なく配置")
        self.cb_as_scene = QCheckBox("シーンとして配置")
        self.cb_auto_speed = QCheckBox("長さで速度変更")
        self.cb_redzone = QCheckBox("REDZONEモード")
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
        speed_h.addWidget(QLabel(" 基準:"))
        speed_h.addWidget(self.base_len)
        speed_h.addWidget(QLabel("秒"))
        speed_h.addStretch()
        opt_grid.addLayout(speed_h, 2, 1)
        left.addLayout(opt_grid)

        tc_group = QFrame()
        tc_group.setStyleSheet("QFrame { background-color: #333; border: 1px solid #555; border-radius: 6px; }")
        tc_lay = QVBoxLayout(tc_group)
        
        tc_top_h = QHBoxLayout()
        self.cb_time_ctrl = QCheckBox("時間制御")
        self.cb_apply_easing = QCheckBox("イージングを適用")
        self.cb_apply_easing.setChecked(False) 
        tc_top_h.addWidget(self.cb_time_ctrl)
        tc_top_h.addWidget(self.cb_apply_easing)
        tc_top_h.addStretch()

        step_h = QHBoxLayout()
        step_h.setContentsMargins(5, 0, 5, 5)
        step_h.addWidget(QLabel("コマ送り:"))
        self.tc_step = QLineEdit("1")
        self.tc_step.setFixedWidth(45)
        step_h.addWidget(self.tc_step)
        step_h.addStretch()
        
        tc_lay.addLayout(tc_top_h)
        tc_lay.addLayout(step_h)
        left.addWidget(tc_group)

        form_grid = QGridLayout()
        self.fps_in = self.add_form_row(form_grid, "FPS:", "60", 0)
        self.scene_in = self.add_form_row(form_grid, "シーン番号:", "1", 1)
        left.addLayout(form_grid)

        left.addWidget(QLabel("トラック選択:"))
        self.track_tree = QTreeWidget()
        self.track_tree.setHeaderHidden(True)
        self.track_tree.setFixedHeight(120)
        self.track_tree.setStyleSheet("background-color: #2a2a2a; color: white;")

        self.root_item = QTreeWidgetItem(self.track_tree, ["* 全トラック"])
        self.root_item.setCheckState(0, Qt.CheckState.Checked)
        self.track_tree.itemChanged.connect(self.on_tree_changed)
        left.addWidget(self.track_tree)

        left.addWidget(QLabel("スクリプト制御:"))
        self.script_edit = QTextEdit()
        left.addWidget(self.script_edit)
        
        run_btn = QPushButton("オブジェクトを出力")
        run_btn.setStyleSheet("height: 50px; background-color: #0D47A1; color: white; font-weight: bold; border-radius: 4px;")
        run_btn.clicked.connect(self.run_process)
        left.addWidget(run_btn)

        right_widget = QWidget()
        right = QVBoxLayout(right_widget)
        
        right.addWidget(QLabel("<b>イージング設定</b>"))
        self.bezier_ui = BezierCanvas() 
        right.addWidget(self.bezier_ui)
        
        right.addWidget(QLabel("<b>追加エフェクト</b>"))
        h_eff = QHBoxLayout()
        self.eff_combo = QComboBox()
        self.eff_combo.addItems(list(EffDict.keys()))
        add_btn = QPushButton("＋追加")
        add_btn.setFixedWidth(80)
        add_btn.setStyleSheet("background-color: #2E7D32; color: white; font-weight: bold;")
        add_btn.clicked.connect(self.add_eff_ui)
        h_eff.addWidget(self.eff_combo)
        h_eff.addWidget(add_btn)
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

    def add_file_row(self, layout, label, func):
        h = QHBoxLayout(); h.addWidget(QLabel(label)); e = QLineEdit()
        b = QPushButton("選択"); b.setFixedWidth(60); b.clicked.connect(func)
        h.addWidget(e); h.addWidget(b); layout.addLayout(h); return e

    def add_form_row(self, grid, label, default, row):
        grid.addWidget(QLabel(label), row, 0); e = QLineEdit(default); grid.addWidget(e, row, 1); return e

    def select_rpp(self):
        p, _ = QFileDialog.getOpenFileName(self, "RPP選択", "", "*.rpp")
        if p: self.rpp_path.setText(p); self.load_tracks(p)

    def save_exo(self):
        p, _ = QFileDialog.getSaveFileName(self, "保存", "", "*.object")
        if p: self.exo_path.setText(p)

    def select_src(self):
        p, _ = QFileDialog.getOpenFileName(self, "素材選択")
        if p: self.src_path.setText(p)

    def load_tracks(self, path):
        self.root_item.takeChildren()
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
        except: pass

    def add_eff_ui(self):
        name = self.eff_combo.currentText()
        frame = QFrame(); frame.setObjectName("EffectCard")
        frame.setStyleSheet("""
            QFrame#EffectCard { background-color: #383838; border: 1px solid #555; border-radius: 8px; margin-bottom: 8px; }
            QLabel { color: #DDD; border: none; }
            QPushButton#DelBtn { background-color: #C62828; color: white; border-radius: 11px; font-weight: bold; border: none; }
        """)
        outer_lay = QVBoxLayout(frame); outer_lay.setContentsMargins(0, 0, 0, 0); outer_lay.setSpacing(0)
        header = QWidget(); header.setFixedHeight(30); header.setStyleSheet("background-color: #4A4A4A; border-top-left-radius: 7px; border-top-right-radius: 7px;")
        header_lay = QHBoxLayout(header); header_lay.setContentsMargins(10, 0, 10, 0)
        lbl = QLabel(f"<b>{name}</b>"); del_btn = QPushButton("×"); del_btn.setObjectName("DelBtn"); del_btn.setFixedSize(22, 22)
        header_lay.addWidget(lbl); header_lay.addStretch(); header_lay.addWidget(del_btn); outer_lay.addWidget(header)
        
        content = QWidget(); lay = QGridLayout(content); lay.setContentsMargins(10, 8, 10, 10); lay.setSpacing(8)
        info = {"name": name, "params": [], "frame": frame}
        del_btn.clicked.connect(lambda: self.remove_eff(info))
        
        row = 0
        for pdef in EffDict[name]:
            pname = pdef[0]
            if len(pdef) > 2 and pdef[2] == -1:
                cb = QCheckBox(pname); cb.setChecked(bool(pdef[1])); lay.addWidget(cb, row, 0, 1, 5)
                info["params"].append({"type": "cb", "name": pname, "obj": cb})
            else:
                lay.addWidget(QLabel(pname), row, 0)
                s_ed = QLineEdit(str(pdef[1])); s_ed.setFixedWidth(50); lay.addWidget(s_ed, row, 2)
                m_cb = QComboBox(); m_cb.addItems(list(XDict.keys())); m_cb.setMinimumWidth(100); lay.addWidget(m_cb, row, 3)
                e_ed = QLineEdit(str(pdef[1])); e_ed.setFixedWidth(50); lay.addWidget(e_ed, row, 4)
                info["params"].append({"type": "motion", "name": pname, "start": s_ed, "end": e_ed, "method": m_cb})
            row += 1
        outer_lay.addWidget(content); self.eff_list_layout.addWidget(frame); self.added_effects_data.append(info)

    def remove_eff(self, info):
        self.added_effects_data.remove(info); info["frame"].deleteLater()

    def run_process(self):
        try:
            rpp_p, exo_p, src_p = self.rpp_path.text(), self.exo_path.text(), self.src_path.text()
            if not rpp_p or not exo_p: return
            fps = float(self.fps_in.text() or 60); base_s = float(self.base_len.text() or 1.0)
            bez_str = self.bezier_ui.bezier_str
            is_redzone = self.cb_redzone.isChecked()

            active_tracks = []
            for i in range(self.root_item.childCount()):
                child = self.root_item.child(i)
                if child.checkState(0) == Qt.CheckState.Checked:
                    active_tracks.append(child.data(0, Qt.ItemDataRole.UserRole))

            with open(rpp_p, 'r', encoding='utf-8', errors='ignore') as f: content = f.read()
            items_raw = content.split("<ITEM"); all_objs = []; track_count = 0
            for i in range(1, len(items_raw)):
                track_count += items_raw[i-1].count("<TRACK")
                if track_count not in active_tracks: continue 

                block = items_raw[i]; pos, length = 0.0, 0.0
                for line in block.split("\n"):
                    ls = line.strip()
                    if ls.startswith("POSITION"): pos = float(ls.split()[1])
                    if ls.startswith("LENGTH"): length = float(ls.split()[1])
                all_objs.append({"pos": pos, "length": length, "track": track_count})

            objs_to_process = sorted(all_objs, key=lambda x: (x["track"], x["pos"]))
            output = []; last_end_frames = {}; track_item_counts = {}; total_obj_idx = 0

            for o in objs_to_process:
                t_idx = o["track"]
                if t_idx not in track_item_counts: track_item_counts[t_idx] = 0
                track_item_counts[t_idx] += 1
                curr_cnt = track_item_counts[t_idx]

                start_f = round(o["pos"] * fps)
                if self.cb_no_gap.isChecked() and t_idx in last_end_frames:
                    if abs(start_f - (last_end_frames[t_idx] + 1)) < 5: start_f = last_end_frames[t_idx] + 1
                end_f = start_f + round(o["length"] * fps) - 1
                last_end_frames[t_idx] = end_f

                speed = (base_s / o["length"] * 100.0) if self.cb_auto_speed.isChecked() else 100.0
                main_layer = t_idx * 2 if self.cb_time_ctrl.isChecked() else t_idx
                
                obj_x = 0.0
                if is_redzone:
                    if t_idx == 1: obj_x = -480.0
                    elif t_idx == 2: obj_x = 480.0

                output.append(f"[{total_obj_idx}]\nframe={start_f},{end_f}\nlayer={main_layer}\n")
                if self.cb_as_scene.isChecked():
                    output.append(f"[{total_obj_idx}.0]\neffect.name=シーン\n再生位置=0.000\n再生速度={speed:.2f}\nシーン={self.scene_in.text()}\n")
                else:
                    output.append(f"[{total_obj_idx}.0]\neffect.name=動画ファイル\nファイル={src_p}\n再生位置=0.000\n再生速度={speed:.2f}\n音声付き=1\n")
                
                output.append(f"[{total_obj_idx}.1]\neffect.name=標準描画\nX={obj_x:.2f}\nY=0.00\nZ=0.00\n透明度=0.00\n")
                p_idx = 2

                if is_redzone and (t_idx == 1 or t_idx == 2):
                    output.append(f"[{total_obj_idx}.{p_idx}]\neffect.name=クリッピング\n上=0\n下=0\n左=480\n右=480\n")
                    p_idx += 1

                h_f, v_f = 0, 0
                if self.cb_flip_h.isChecked() or self.cb_flip_v.isChecked():
                    if curr_cnt % 2 == 0:
                        h_f = 1 if self.cb_flip_h.isChecked() else 0
                        v_f = 1 if self.cb_flip_v.isChecked() else 0

                if h_f == 1 or v_f == 1:
                    output.append(f"[{total_obj_idx}.{p_idx}]\neffect.name=反転\n上下反転={v_f}\n左右反転={h_f}\n")
                    p_idx += 1

                for eff in self.added_effects_data:
                    output.append(f"[{total_obj_idx}.{p_idx}]\neffect.name={eff['name']}\n")
                    for p in eff["params"]:
                        val = int(p["obj"].isChecked()) if p["type"] == "cb" else f"{p['start'].text()},{p['end'].text()},{XDict[p['method'].currentText()]},{bez_str}"
                        output.append(f"{p['name']}={val}\n")
                    p_idx += 1

                output.append("\n"); total_obj_idx += 1

                if self.cb_time_ctrl.isChecked():
                    s_val, e_val = (("100.000", "0.000") if curr_cnt % 2 == 0 else ("0.000", "100.000")) if is_redzone else ("0.000", "100.000")
                    t_easing = bez_str if self.cb_apply_easing.isChecked() else "0"
                    output.append(f"[{total_obj_idx}]\nframe={start_f},{end_f}\nlayer={main_layer-1}\n")
                    output.append(f"[{total_obj_idx}.0]\neffect.name=時間制御(オブジェクト)\n"
                                f"位置={s_val},{e_val},直線移動(時間制御),{t_easing}\n"
                                f"コマ落ち={self.tc_step.text()}\n"
                                f"対象レイヤー数=1\n\n")
                    total_obj_idx += 1

            with open(exo_p, 'w', encoding='utf-8') as f: f.write("".join(output))
            QMessageBox.information(self, "完了", f"{'REDZONEモードの' if is_redzone else ''}出力が完了しました。")
        except Exception as e: QMessageBox.critical(self, "エラー", str(e))

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