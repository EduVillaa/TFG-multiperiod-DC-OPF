from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QFileDialog,
    QVBoxLayout, QHBoxLayout, QMessageBox, QComboBox,
    QSpinBox, QDoubleSpinBox, QDateEdit, QFormLayout,
    QProgressBar, QTimeEdit
)
from PySide6.QtCore import QThread, Signal, QDate, QTime

from GridReader import run_program
import traceback
import pandas as pd

def getBatteryOptimizationMode(filename):
    df = pd.read_excel(
        filename,
        sheet_name="StorageUnit",
        header=2
    ).iloc[:, 1:]
    print("Optimización de batería?")

    return bool(
        df["Optimization mode"].isin(
            ["Optimize both", "Optimize MW", "Optimize MWh"]
        ).any()
    )


class Worker(QThread):
    finished = Signal()
    error = Signal(str)
    progress = Signal(int, str)

    def __init__(self, input_path=None, system_parameters=None):
        super().__init__()
        self.input_path = input_path
        self.system_parameters = system_parameters

    def run(self):
        try:
            run_program(
                self.input_path,
                self.system_parameters,
                progress_callback=self.progress.emit
            )
            self.finished.emit()
        except Exception:
            error_text = traceback.format_exc()  
            print(error_text)                   
            self.error.emit(error_text)          

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("DC OPF")
        self.resize(650, 430)

        self.input_path = None
        self.worker = None

        self.file_label = QLabel("Excel archive: GridInputs.xlsx (default)")
        self.btn_select = QPushButton("Select excel")
        self.btn_run = QPushButton("Execute DC OPF")
        self.status = QLabel("State: Ready")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")

        self.voll_input = QDoubleSpinBox()
        self.voll_input.setRange(0, 1_000_000)
        self.voll_input.setValue(10000)
        self.voll_input.setDecimals(0)
        self.voll_input.setSuffix(" €/MWh")

        self.horizon_combo = QComboBox()
        self.horizon_combo.addItems(["Static", "Multiperiod"])
        self.horizon_combo.setCurrentText("Multiperiod")

        self.min_start_date = QDate(2015, 1, 5)
        self.max_end_date = QDate(2024, 12, 31)

        self.start_date_edit = QDateEdit()

        # Campos para simulación estática
        self.static_date_edit = QDateEdit()
        self.static_date_edit.setCalendarPopup(True)
        self.static_date_edit.setDisplayFormat("dd/MM/yyyy")
        self.static_date_edit.setMinimumDate(self.min_start_date)
        self.static_date_edit.setMaximumDate(self.max_end_date)
        self.static_date_edit.setDate(QDate(2022, 1, 1))

        self.static_hour_edit = QTimeEdit()
        self.static_hour_edit.setDisplayFormat("HH:mm")
        self.static_hour_edit.setTime(QTime(12, 0))

        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDisplayFormat("dd/MM/yyyy")
        self.start_date_edit.setMinimumDate(self.min_start_date)
        self.start_date_edit.setMaximumDate(self.max_end_date)
        self.start_date_edit.setDate(QDate(2022, 1, 1))

        self.duration_input = QSpinBox()
        self.duration_input.setRange(1, 3650)
        self.duration_input.setValue(1)
        self.duration_input.setSuffix(" days")

        self.resolution_combo = QComboBox()
        self.resolution_combo.addItems(["Auto", "Hourly", "Daily", "Weekly"])
        self.resolution_combo.setCurrentText("Auto")

        self.end_date_label = QLabel("End date: -")

        # Campos para optimización de batería
        self.discount_rate_input = QDoubleSpinBox()
        self.discount_rate_input.setRange(0.0, 100.0)
        self.discount_rate_input.setDecimals(2)
        self.discount_rate_input.setValue(7.0)
        self.discount_rate_input.setSuffix(" %")

        self.default_battery_lifetime_input = QSpinBox()
        self.default_battery_lifetime_input.setRange(1, 100)
        self.default_battery_lifetime_input.setValue(15)
        self.default_battery_lifetime_input.setSuffix(" years")

        # Solver options
        self.mip_rel_gap_input = QDoubleSpinBox()
        self.mip_rel_gap_input.setRange(0.0, 100.0)
        self.mip_rel_gap_input.setDecimals(3)
        self.mip_rel_gap_input.setSingleStep(0.1)
        self.mip_rel_gap_input.setValue(0.1)
        self.mip_rel_gap_input.setSuffix(" %")
        self.mip_rel_gap_input.setToolTip(
        "Relative MIP gap in percent. Example: 0.1% = 0.001 for the solver."
        )

        self.time_limit_input = QSpinBox()
        self.time_limit_input.setRange(0, 1_000_000)
        self.time_limit_input.setValue(3600)
        self.time_limit_input.setSuffix(" s")
        self.time_limit_input.setToolTip("Solver time limit in seconds. 3600 = 1 hour")

        form_layout = QFormLayout()

        self.voll_row_label = QLabel("VOLL")
        self.horizon_row_label = QLabel("Static / Multiperiod")
        self.start_date_row_label = QLabel("Start date")
        self.duration_row_label = QLabel("Simulation duration")
        self.static_date_row_label = QLabel("Static snapshot date")
        self.static_hour_row_label = QLabel("Static snapshot hour")
        self.end_date_row_label = QLabel("End date")
        self.resolution_row_label = QLabel("Graph resolution")
        self.discount_rate_row_label = QLabel("Discount rate")
        self.default_battery_lifetime_row_label = QLabel("Default battery lifetime")
        self.mip_rel_gap_row_label = QLabel("MIP relative gap (%)")
        self.time_limit_row_label = QLabel("Solver time limit")

        form_layout.addRow(self.voll_row_label, self.voll_input)
        form_layout.addRow(self.horizon_row_label, self.horizon_combo)
        form_layout.addRow(self.start_date_row_label, self.start_date_edit)
        form_layout.addRow(self.static_date_row_label, self.static_date_edit)
        form_layout.addRow(self.static_hour_row_label, self.static_hour_edit)
        form_layout.addRow(self.duration_row_label, self.duration_input)
        form_layout.addRow(self.end_date_row_label, self.end_date_label)
        form_layout.addRow(self.resolution_row_label, self.resolution_combo)
        form_layout.addRow(self.mip_rel_gap_row_label, self.mip_rel_gap_input)
        form_layout.addRow(self.time_limit_row_label, self.time_limit_input)

        # Los dos últimos
        form_layout.addRow(self.discount_rate_row_label, self.discount_rate_input)
        form_layout.addRow(
            self.default_battery_lifetime_row_label,
            self.default_battery_lifetime_input
        )

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.btn_select)
        button_layout.addWidget(self.btn_run)

        main_layout = QVBoxLayout()
        main_layout.addWidget(self.file_label)
        main_layout.addLayout(form_layout)
        main_layout.addLayout(button_layout)
        main_layout.addWidget(self.progress_bar)
        main_layout.addWidget(self.status)

        self.setLayout(main_layout)

        self.btn_select.clicked.connect(self.select_file)
        self.btn_run.clicked.connect(self.execute_program)
        self.start_date_edit.dateChanged.connect(self.update_duration_limit)
        self.start_date_edit.dateChanged.connect(self.update_end_date_label)
        self.duration_input.valueChanged.connect(self.update_end_date_label)
        self.horizon_combo.currentTextChanged.connect(self.update_dynamic_fields_visibility)

        self.update_duration_limit()
        self.update_end_date_label()
        self.update_dynamic_fields_visibility()

    def get_selected_input_file(self) -> str:
        return self.input_path if self.input_path else "GridInputs.xlsx"

    def battery_optimization_enabled(self) -> bool:
        try:
            print("intentar usar la función de optimización")
            return bool(getBatteryOptimizationMode(self.get_selected_input_file()))
        except Exception:
            print("no se consiguió")
            return False

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar GridInputs.xlsx",
            "",
            "Excel Files (*.xlsx)"
        )
        if file_path:
            self.input_path = file_path
            self.file_label.setText(f"Archivo Excel: {file_path}")
            self.update_dynamic_fields_visibility()

    def get_system_parameters_from_gui(self) -> dict:
        horizon = self.horizon_combo.currentText()

        params = {
            "VOLL (€/MWh)": self.voll_input.value(),
            "Static / Multiperiod": horizon,
            "mip_rel_gap": self.mip_rel_gap_input.value() / 100,
            "time_limit": self.time_limit_input.value(),
        }

        if horizon == "Multiperiod":
            params.update({
                "Start date (dd/mm/aaaa)": self.start_date_edit.date().toPython(),
                "Simulation duration (days)": self.duration_input.value(),
                "Graph resolution": self.resolution_combo.currentText(),
            })

            if self.battery_optimization_enabled():
                params.update({
                    "Discount rate (%)": self.discount_rate_input.value(),
                    "Default battery lifetime (years)": self.default_battery_lifetime_input.value(),
                })

        elif horizon == "Static":
            static_date = self.static_date_edit.date().toPython()
            static_time = self.static_hour_edit.time().toPython()

            params.update({
                "Static snapshot date (dd/mm/aaaa)": static_date,
                "Static snapshot hour": static_time,
                "Static snapshot datetime": pd.Timestamp.combine(static_date, static_time),
            })

        return params

    def execute_program(self):
        if self.horizon_combo.currentText() == "Multiperiod":
            valid, error_msg = self.validate_simulation_dates()
            if not valid:
                QMessageBox.warning(self, "Fechas no válidas", error_msg)
                return

        elif self.horizon_combo.currentText() == "Static":
            valid, error_msg = self.validate_static_snapshot()
            if not valid:
                QMessageBox.warning(self, "Snapshot no válido", error_msg)
                return

        self.btn_run.setEnabled(False)
        self.progress_bar.setValue(0)
        self.status.setText("Estado: ejecutando...")

        system_parameters = self.get_system_parameters_from_gui()

        self.worker = Worker(
            input_path=self.input_path,
            system_parameters=system_parameters
        )
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_progress(self, value, message):
        self.progress_bar.setValue(value)
        self.status.setText(f"Estado: {message}")

    def on_finished(self):
        self.progress_bar.setValue(100)
        self.status.setText("Estado: terminado")
        self.btn_run.setEnabled(True)
        QMessageBox.information(self, "Éxito", "Programa ejecutado correctamente.")

    def on_error(self, error_msg):
        self.status.setText("Estado: error")
        self.btn_run.setEnabled(True)

        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Critical)
        msg.setWindowTitle("Error")
        msg.setText("Ha ocurrido un error.")
        
        # 👇 esto es la clave
        msg.setDetailedText(error_msg)

        msg.exec()

    def update_duration_limit(self):
        start_date = self.start_date_edit.date()
        max_days = start_date.daysTo(self.max_end_date) + 1

        self.duration_input.setMaximum(max_days)

        if self.duration_input.value() > max_days:
            self.duration_input.setValue(max_days)

    def update_end_date_label(self):
        start_date = self.start_date_edit.date()
        duration_days = self.duration_input.value()
        end_date = start_date.addDays(duration_days - 1)
        self.end_date_label.setText(end_date.toString("dd/MM/yyyy"))

    def validate_simulation_dates(self) -> tuple[bool, str]:
        start_date = self.start_date_edit.date()
        duration_days = self.duration_input.value()
        end_date = start_date.addDays(duration_days - 1)

        if start_date < self.min_start_date:
            return False, "La fecha de inicio no puede ser anterior al 05/01/2015."

        if end_date > self.max_end_date:
            return False, "La fecha final de la simulación no puede superar el 31/12/2024."

        return True, ""
    
    def validate_static_snapshot(self) -> tuple[bool, str]:
        static_date = self.static_date_edit.date()

        if static_date < self.min_start_date:
            return False, "La fecha del snapshot estático no puede ser anterior al 01/01/2015."

        if static_date > self.max_end_date:
            return False, "La fecha del snapshot estático no puede superar el 31/12/2024."

        return True, ""

    def update_dynamic_fields_visibility(self):
        horizon = self.horizon_combo.currentText()

        is_multiperiod = horizon == "Multiperiod"
        is_static = horizon == "Static"

        battery_opt_enabled = self.battery_optimization_enabled()
        show_battery_fields = is_multiperiod and battery_opt_enabled

        # Campos propios de simulación multiperiodo
        self.start_date_row_label.setVisible(is_multiperiod)
        self.start_date_edit.setVisible(is_multiperiod)

        self.duration_row_label.setVisible(is_multiperiod)
        self.duration_input.setVisible(is_multiperiod)

        self.end_date_row_label.setVisible(is_multiperiod)
        self.end_date_label.setVisible(is_multiperiod)

        self.resolution_row_label.setVisible(is_multiperiod)
        self.resolution_combo.setVisible(is_multiperiod)

        # Campos propios de simulación estática
        self.static_date_row_label.setVisible(is_static)
        self.static_date_edit.setVisible(is_static)

        self.static_hour_row_label.setVisible(is_static)
        self.static_hour_edit.setVisible(is_static)

        # Campos de optimización de batería
        # Solo aparecen si:
        # 1. La simulación es multiperiodo
        # 2. En el Excel hay alguna batería con modo de optimización
        self.discount_rate_row_label.setVisible(show_battery_fields)
        self.discount_rate_input.setVisible(show_battery_fields)

        self.default_battery_lifetime_row_label.setVisible(show_battery_fields)
        self.default_battery_lifetime_input.setVisible(show_battery_fields)

        # Opciones del solver
        self.mip_rel_gap_row_label.setVisible(True)
        self.mip_rel_gap_input.setVisible(True)

        self.time_limit_row_label.setVisible(True)
        self.time_limit_input.setVisible(True)


if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()
