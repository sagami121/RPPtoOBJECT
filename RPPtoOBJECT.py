import os
import sys
import math
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QGridLayout, QLabel, QLineEdit, 
                             QPushButton, QCheckBox, QComboBox, QTextEdit, 
                             QFileDialog, QMessageBox, QScrollArea, QFrame)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPalette, QColor

EffDict = {
    "座標": [["X", 0.0], ["Y", 0.0], ["Z", 0.0]],
    "拡大率": [["拡大率", 100.00], ["X", 100.00], ["Y", 100.00]],
    "透明度": [["透明度", 0.0]],
    "回転": [["X", 0.0], ["Y", 0.0], ["Z", 0.0]],
    "アニメーション効果": [["track0", 0.00], ["track1", 0.00], ["check0", 0, -1], ["type", 0], ["name", ""]]
}

XDict = {
    "移動なし": "", "直線移動": 1, "加減速移動": 103, "曲線移動": 2, "瞬間移動": 3,
    "イージング": "15@イージング",
}

class RPPtoObjectApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("RPPtoOBJECT v1.0")
        self.setMinimumSize(1000, 750)
        self.added_effects_data = []
        self.init_ui()

    def init_ui(self):
        cw = QWidget()
        self.setCentralWidget(cw)
        main_layout = QHBoxLayout(cw)

        left = QVBoxLayout()
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
        self.base_len = QLineEdit("1.0")
        self.base_len.setFixedWidth(40)

        opt_grid.addWidget(self.cb_flip_h, 0, 0)
        opt_grid.addWidget(self.cb_flip_v, 0, 1)
        opt_grid.addWidget(self.cb_loop, 1, 0)
        opt_grid.addWidget(self.cb_no_gap, 1, 1)
        opt_grid.addWidget(self.cb_as_scene, 2, 0)
        
        speed_h = QHBoxLayout()
        speed_h.addWidget(self.cb_auto_speed)
        speed_h.addWidget(QLabel(" 基準:"))
        speed_h.addWidget(self.base_len)
        speed_h.addWidget(QLabel("秒"))
        speed_h.addStretch()
        opt_grid.addLayout(speed_h, 2, 1)
        left.addLayout(opt_grid)

        tc_group = QFrame()
        tc_group.setFrameShape(QFrame.Shape.StyledPanel)
        tc_group.setStyleSheet("QFrame { background-color: #333333; border: 1px solid #555555; border-radius: 4px; }")
        tc_lay = QGridLayout(tc_group)
        self.cb_time_ctrl = QCheckBox("時間制御")
        self.tc_step = QLineEdit("1")
        self.tc_step.setFixedWidth(40)
        
        tc_lay.addWidget(self.cb_time_ctrl, 0, 0, 1, 3)
        tc_lay.addWidget(QLabel("コマ送り:"), 1, 0)
        tc_lay.addWidget(self.tc_step, 1, 1)
        left.addWidget(tc_group)

        form_grid = QGridLayout()
        self.fps_in = self.add_form_row(form_grid, "FPS:", "60", 0)
        self.scene_in = self.add_form_row(form_grid, "シーン番号:", "1", 1)
        self.track_combo = QComboBox()
        self.track_combo.addItem("全トラック")
        form_grid.addWidget(QLabel("トラック選択:"), 2, 0)
        form_grid.addWidget(self.track_combo, 2, 1)
        left.addLayout(form_grid)

        left.addWidget(QLabel("スクリプト制御:"))
        self.script_edit = QTextEdit()
        left.addWidget(self.script_edit)

        run_btn = QPushButton("オブジェクトを出力")
        run_btn.setStyleSheet("height: 50px; background-color: #0D47A1; color: white; font-weight: bold; border-radius: 4px;")
        run_btn.clicked.connect(self.run_process)
        left.addWidget(run_btn)
        main_layout.addLayout(left, 3)

        right = QVBoxLayout()
        right.addWidget(QLabel("<b>追加エフェクト</b>"))
        h_eff = QHBoxLayout()
        self.eff_combo = QComboBox()
        self.eff_combo.addItems(list(EffDict.keys()))
        add_btn = QPushButton("＋追加")
        add_btn.clicked.connect(self.add_eff_ui)
        h_eff.addWidget(self.eff_combo); h_eff.addWidget(add_btn)
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
        main_layout.addLayout(right, 2)

    def add_file_row(self, layout, label, func):
        h = QHBoxLayout(); h.addWidget(QLabel(label))
        e = QLineEdit(); b = QPushButton("選択"); b.setFixedWidth(60)
        b.clicked.connect(func); h.addWidget(e); h.addWidget(b)
        layout.addLayout(h); return e

    def add_form_row(self, grid, label, default, row):
        grid.addWidget(QLabel(label), row, 0)
        e = QLineEdit(default); grid.addWidget(e, row, 1)
        return e

    def select_rpp(self):
        p, _ = QFileDialog.getOpenFileName(self, "RPP選択", "", "*.rpp")
        if p: 
            self.rpp_path.setText(p)
            self.load_tracks(p)

    def save_exo(self):
        p, _ = QFileDialog.getSaveFileName(self, "保存", "", "*.object")
        if p: self.exo_path.setText(p)

    def select_src(self):
        p, _ = QFileDialog.getOpenFileName(self, "素材選択")
        if p: self.src_path.setText(p)

    def load_tracks(self, path):
        self.track_combo.clear()
        self.track_combo.addItem("全トラック")
        try:
            with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                idx = 0
                for line in f:
                    if "<TRACK" in line:
                        idx += 1
                        l2 = next(f)
                        name = l2.replace("NAME ", "").strip() if "NAME" in l2 else f"Track {idx}"
                        self.track_combo.addItem(f"{idx}: {name}")
        except: pass

    def add_eff_ui(self):
        name = self.eff_combo.currentText()
        frame = QFrame()
        frame.setFrameShape(QFrame.Shape.StyledPanel)
        frame.setStyleSheet("QFrame { background-color: #333333; border: 1px solid #555555; margin-bottom: 5px; border-radius: 4px; }")
        lay = QGridLayout(frame)
        info = {"name": name, "params": [], "frame": frame}
        
        title_h = QHBoxLayout()
        title_h.addWidget(QLabel(f"<b>{name}</b>"))
        del_btn = QPushButton("×"); del_btn.setFixedWidth(30); del_btn.clicked.connect(lambda: self.remove_eff(info))
        title_h.addStretch(); title_h.addWidget(del_btn)
        lay.addLayout(title_h, 0, 0, 1, 6)

        row = 1
        for pdef in EffDict[name]:
            pname = pdef[0]
            if len(pdef) > 2 and pdef[2] == -1:
                cb = QCheckBox(pname); cb.setChecked(bool(pdef[1]))
                lay.addWidget(cb, row, 0, 1, 6)
                info["params"].append({"type": "cb", "name": pname, "obj": cb})
            else:
                lay.addWidget(QLabel(pname), row, 0)
                s_ed = QLineEdit(str(pdef[1])); s_ed.setFixedWidth(60); lay.addWidget(s_ed, row, 2)
                m_cb = QComboBox(); m_cb.addItems(list(XDict.keys())); m_cb.setFixedWidth(100); lay.addWidget(m_cb, row, 3)
                e_ed = QLineEdit(""); e_ed.setFixedWidth(60); lay.addWidget(e_ed, row, 4)
                info["params"].append({"type": "motion", "name": pname, "start": s_ed, "end": e_ed, "method": m_cb})
            row += 1
        self.eff_list_layout.addWidget(frame)
        self.added_effects_data.append(info)

    def remove_eff(self, info):
        self.added_effects_data.remove(info)
        info["frame"].setParent(None); info["frame"].deleteLater()

    def run_process(self):
        try:
            rpp_p, exo_p, src_p = self.rpp_path.text(), self.exo_path.text(), self.src_path.text()
            if not rpp_p or not exo_p: return QMessageBox.warning(self, "エラー", "パスを指定してください")
            fps = float(self.fps_in.text() or 60)
            base_s = float(self.base_len.text() or 1.0)
            
            with open(rpp_p, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            
            items_raw = content.split("<ITEM")
            all_objs = []
            track_count = 0
            for i in range(1, len(items_raw)):
                track_count += items_raw[i-1].count("<TRACK")
                block = items_raw[i]
                pos, length = 0.0, 0.0
                for line in block.split("\n"):
                    ls = line.strip()
                    if ls.startswith("POSITION"): pos = float(ls.split()[1])
                    if ls.startswith("LENGTH"): length = float(ls.split()[1])
                all_objs.append({"pos": pos, "length": length, "track": track_count})

            sel_track = self.track_combo.currentIndex()
            objs_to_process = [o for o in all_objs if sel_track == 0 or o["track"] == sel_track]
            objs_to_process.sort(key=lambda x: x["pos"])

            output = []
            last_end_frame = -1
            total_obj_idx = 0

            for item_num, o in enumerate(objs_to_process):
                start_frame = round(o["pos"] * fps)
                duration_frames = round(o["length"] * fps)
                if self.cb_no_gap.isChecked() and last_end_frame != -1:
                    if abs(start_frame - (last_end_frame + 1)) < 5: 
                        start_frame = last_end_frame + 1
                end_frame = start_frame + duration_frames - 1
                last_end_frame = end_frame

                speed = 100.0
                if self.cb_auto_speed.isChecked():
                    raw_speed = (base_s / o["length"] * 100.0)
                    speed = 100.0 * (2 ** max(0, round(math.log2(raw_speed / 100.0)))) if raw_speed > 150 else 100.0

                output.append(f"[{total_obj_idx}]\nframe={start_frame},{end_frame}\nlayer=2\n")
                if self.cb_as_scene.isChecked():
                    output.append(f"[{total_obj_idx}.0]\neffect.name=シーン\n再生位置=1.000\n再生速度={speed:.2f}\nシーン={self.scene_in.text()}\n")
                else:
                    output.append(f"[{total_obj_idx}.0]\neffect.name=動画ファイル\nファイル={src_p}\n再生位置=0.000\n再生速度={speed:.2f}\n音声付き=1\n")
                
                output.append(f"[{total_obj_idx}.1]\neffect.name=標準描画\nX=0.00\nY=0.00\nZ=0.00\n透明度=0.00\n")
                
                p_idx = 2
                if (self.cb_flip_h.isChecked() or self.cb_flip_v.isChecked()) and (item_num % 2 != 0):
                    output.append(f"[{total_obj_idx}.{p_idx}]\neffect.name=反転\n上下反転={int(self.cb_flip_v.isChecked())}\n左右反転={int(self.cb_flip_h.isChecked())}\n")
                    p_idx += 1
                for eff in self.added_effects_data:
                    output.append(f"[{total_obj_idx}.{p_idx}]\neffect.name={eff['name']}\n")
                    for p in eff["params"]:
                        if p["type"] == "cb": output.append(f"{p['name']}={int(p['obj'].isChecked())}\n")
                        else:
                            meth = XDict.get(p["method"].currentText(), "")
                            s, e = p["start"].text(), p["end"].text()
                            output.append(f"{p['name']}={(f'{s},{e},{meth}' if meth!='' else s)}\n")
                    p_idx += 1
                output.append("\n")
                total_obj_idx += 1

                if self.cb_time_ctrl.isChecked():
                    time_val = "100.000,0.000,直線移動,0" if item_num % 2 != 0 else "0.000,100.000,直線移動,0"
                    output.append(f"[{total_obj_idx}]\nframe={start_frame},{end_frame}\nlayer=1\n")
                    output.append(f"[{total_obj_idx}.0]\neffect.name=時間制御(オブジェクト)\n位置={time_val}\nコマ落ち={self.tc_step.text()}\n対象レイヤー数=1\n\n")
                    total_obj_idx += 1

            with open(exo_p, 'w', encoding='utf-8') as f:
                f.write("".join(output))

            QMessageBox.information(self, "完了", f"オブジェクト出力が完了しました。\n{exo_p}")

        except Exception as e:
            QMessageBox.critical(self, "エラー", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv); app.setStyle("Fusion")
    dark_p = QPalette()
    dark_p.setColor(QPalette.ColorRole.Window, QColor(45, 45, 45))
    dark_p.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
    dark_p.setColor(QPalette.ColorRole.Base, QColor(30, 30, 30))
    dark_p.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
    dark_p.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
    dark_p.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
    app.setPalette(dark_p)
    win = RPPtoObjectApp(); win.show(); sys.exit(app.exec())