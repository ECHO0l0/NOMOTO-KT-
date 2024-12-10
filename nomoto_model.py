import sys
import numpy as np
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QHBoxLayout, QFileDialog, QMessageBox, QTextEdit
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPalette, QColor

# 计算船舶参数的函数
def calculate_ships_parameters(L=105, B=18, T=5.4, Cb=0.5595, J=5735.5, Ad=11.8, V_knots=16.7, Xg=-0.51, rho=1.025):
    try:
        V = V_knots * 1852 / 3600  # 航速转换为米/秒

        m = J * rho  # 计算船舶质量 (kg)
        m_nondim = m / (0.5 * rho * L**3)  # 质量无量纲化
        Xg_nondim = Xg / L  # 重心无量纲化
        Izz = (m * L**2) / 16  # 船舶惯性矩 (kg·m^2)
        Izz_nondim = Izz / (0.5 * rho * L**5)  # 惯性矩无量纲化
        gama = 0.30  # 舵效增益因子

        Yv1 = -(1 + 0.16 * Cb * B / T - 5.1 * (B / L)**2) * np.pi * (T / L)**2
        Yr1 = -(0.67 * B / L - 0.0033 * (B / T)**2) * np.pi * (T / L)**2
        Nv1 = -(1.1 * B / L - 0.041 * B / T) * np.pi * (T / L)**2
        Nr1 = -(1 / 12 + 0.017 * Cb * B / T - 0.33 * B / L) * np.pi * (T / L)**2

        Yv = -(1 + 0.4 * Cb * B / T) * np.pi * (T / L)**2
        Yr = -(-1 / 2 + 2.2 * B / L - 0.080 * B / T) * np.pi * (T / L)**2
        Nv = -(1 / 2 + 2.4 * T / L) * np.pi * (T / L)**2
        Nr = -(1 / 4 + 0.039 * B / T - 0.56 * B / L) * np.pi * (T / L)**2

        Yd = (3.0 * Ad) / L**2
        Nd = -(1 / 2) * Yd
        Yvc = -gama * Yd
        Yrc = -1 / 2 * Yvc
        Nvc = -1 / 2 * Yvc
        Nrc = 1 / 4 * Yvc

        Yv += Yvc
        Yr += Yrc
        Nv += Nvc
        Nr += Nrc

        # 计算 K0 和 T0
        I2 = np.array([[m_nondim - Yv1, L * (m_nondim * Xg_nondim - Yr1)], 
                        [m_nondim * Xg_nondim - Nv1, L * (Izz_nondim - Nr1)]])
        P2 = np.array([[V * Yv / L, V * (Yr - m_nondim)], 
                        [V * Nv / L, V * (Nr - m_nondim * Xg_nondim)]])
        Q2 = np.array([(V**2) * Yd / L, (V**2) * Nd / L])

        A2 = np.linalg.inv(I2) @ P2
        B2 = np.linalg.inv(I2) @ Q2

        a11, a12 = A2[0, 0], A2[0, 1]
        a21, a22 = A2[1, 0], A2[1, 1]
        b11, b21 = -B2[0], -B2[1]

        K0 = (b11 * a21 - b21 * a11) / (a11 * a22 - a12 * a21)
        T0 = -(a11 + a22) / (a11 * a22 - a12 * a21) - b21 / (b11 * a21 - b21 * a11)

        return K0, T0, Yr1, Nv1  # 返回 K0, T0 和 Yr1, Nv1 作为参数
    except Exception as e:
        return None, None, f"计算出错: {str(e)}"
# 创建窗口
class ShipControlApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('船舶控制系统KT参数计算')
        self.setGeometry(100, 100, 600, 450)

        layout = QVBoxLayout()

        self.input_labels = {
            "L": "船长 (m)", 
            "B": "船宽 (m)", 
            "T": "设计吃水 (m)", 
            "Cb": "方形系数", 
            "J": "排水体积 (m³)", 
            "Ad": "舵叶面积 (m²)", 
            "V": "航速 (节)", 
            "Xg": "重心距中心距 (m)", 
            "rho": "海水密度 (kg/m³)"
        }

        self.default_values = {
            "L": 105.0, 
            "B": 18.0, 
            "T": 5.4, 
            "Cb": 0.5595, 
            "J": 5735.5, 
            "Ad": 11.8, 
            "V": 16.7, 
            "Xg": -0.51, 
            "rho": 1.025
        }

        self.inputs = {}
        for param, label in self.input_labels.items():
            h_layout = QHBoxLayout()
            h_layout.addWidget(QLabel(label))
            input_field = QLineEdit(self)
            input_field.setText(str(self.default_values[param]))  # 默认值
            input_field.setPlaceholderText(f"请输入 {label}")
            self.inputs[param] = input_field
            h_layout.addWidget(input_field)
            layout.addLayout(h_layout)

        self.result_text = QTextEdit(self)
        self.result_text.setReadOnly(True)  # 只读
        self.result_text.setPlaceholderText("计算结果显示在这里")
        layout.addWidget(self.result_text)

        calc_button = QPushButton("计算", self)
        calc_button.clicked.connect(self.calculate)
        layout.addWidget(calc_button)



        export_button = QPushButton("导出结果为Excel", self)
        export_button.clicked.connect(self.export_to_excel)
        layout.addWidget(export_button)

        note_label = QLabel("姓名：施志强 学号：1120241060", self)
        note_label.setAlignment(Qt.AlignCenter)
        note_label.setStyleSheet("font-size: 14px; color: #555;")
        layout.addWidget(note_label)

        self.setLayout(layout)

        self.set_mac_style()

        self.K0_old = None
        self.T0_old = None
        self.Yr1 = None
        self.Nv1 = None

    def set_mac_style(self):
        self.setStyleSheet(""" 
            QWidget { background-color: #f5f5f5; } 
            QPushButton { background-color: #007aff; color: white; border-radius: 12px; padding: 10px 20px; font-size: 14px; } 
            QPushButton:hover { background-color: #0051a8; } 
            QLineEdit, QTextEdit { background-color: #ffffff; border: 1px solid #ccc; border-radius: 8px; padding: 8px; font-size: 14px; } 
            QLabel { font-size: 14px; font-weight: bold; } 
        """)

    def calculate(self):
        try:
            params = {param: float(self.inputs[param].text()) for param in self.input_labels}
        except ValueError:
            self.result_text.setText("请输入有效的数字！")
            return
        
        self.K0_old, self.T0_old, self.Yr1, self.Nv1 = calculate_ships_parameters(
            params["L"], params["B"], params["T"], params["Cb"], params["J"],
            params["Ad"], params["V"], params["Xg"], params["rho"]
        )

        if self.K0_old is not None and self.T0_old is not None:
            result = f"K0 = {self.K0_old:.4f}\nT0 = {self.T0_old:.4f}"
            self.result_text.setText(result)
        else:
            self.result_text.setText("计算出错，请检查输入。")



    def export_to_excel(self):
        try:
            params = {param: float(self.inputs[param].text()) for param in self.input_labels}
            K0, T0 = self.K0_old, self.T0_old
            data = {
                "参数": list(params.keys()) + ['K0', 'T0'],
                "值": list(params.values()) + [K0, T0]
            }

            df = pd.DataFrame(data)

            file_dialog = QFileDialog(self)
            file_path, _ = file_dialog.getSaveFileName(self, "保存为Excel", "船舶参数计算结果.xlsx", "Excel Files (*.xlsx)")
            if not file_path:
                return  # 取消保存

            df.to_excel(file_path, index=False)

            # 弹出保存成功提示框
            QMessageBox.information(self, "导出成功", "结果已成功导出为Excel文件！")
        except Exception as e:
            QMessageBox.warning(self, "导出失败", f"导出时发生错误：{str(e)}")


# 启动应用
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ShipControlApp()
    window.show()
    sys.exit(app.exec_())
