from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QFileDialog,
    QVBoxLayout, QHBoxLayout, QMessageBox, QComboBox,
    QSpinBox, QDoubleSpinBox, QDateEdit, QFormLayout
)
from PySide6.QtCore import QThread, Signal, QDate

from GridReader import run_program


class Worker(QThread):
    finished = Signal()
    error = Signal(str)

    def __init__(self, input_path=None, system_parameters=None):
        super().__init__()
        self.input_path = input_path
        self.system_parameters = system_parameters

    def run(self):
        try:
            run_program(self.input_path, self.system_parameters)
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("DC OPF ")
        self.resize(650, 420)

        self.input_path = None
        self.worker = None

        self.file_label = QLabel("Excel archive: GridInputs.xlsx (default)")
        self.btn_select = QPushButton("Select excel")
        self.btn_run = QPushButton("Execute DC OPF")
        self.status = QLabel("State: Ready")

        self.voll_input = QDoubleSpinBox()
        self.voll_input.setRange(0, 1_000_000)
        self.voll_input.setValue(10000)
        self.voll_input.setDecimals(0)
        self.voll_input.setSuffix(" €/MWh")

        self.horizon_combo = QComboBox()
        self.horizon_combo.addItems(["Static", "Multiperiod"])
        self.horizon_combo.setCurrentText("Multiperiod")

        self.region_combo = QComboBox()
        self.region_combo.addItems([
            "Andalucía", "Aragón", "Asturias", "Baleares", "Canarias",
            "Cantabria", "Castilla-La Mancha", "Castilla y León",
            "Cataluña", "Ceuta", "Comunidad Valenciana", "Extremadura",
            "Galicia", "La Rioja", "Madrid", "Melilla", "Murcia",
            "Navarra", "País Vasco"
        ])
        self.region_combo.setCurrentText("Andalucía")
        
        self.min_start_date = QDate(2022, 1, 1)
        self.max_end_date = QDate(2024, 12, 31)

        self.start_date_edit = QDateEdit()
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDisplayFormat("dd/MM/yyyy")
        self.start_date_edit.setMinimumDate(self.min_start_date)
        self.start_date_edit.setMaximumDate(self.max_end_date)
        self.start_date_edit.setDate(QDate(2022, 1, 1))

        self.duration_input = QSpinBox()
        self.duration_input.setRange(1, 3650)
        self.duration_input.setValue(14)
        self.duration_input.setSuffix(" días")

        self.resolution_combo = QComboBox()
        self.resolution_combo.addItems(["Auto", "Hourly", "Daily", "Weekly"])
        self.resolution_combo.setCurrentText("Auto")

        self.end_date_label = QLabel("End date: -")

        form_layout = QFormLayout()
        form_layout.addRow("VOLL", self.voll_input)
        form_layout.addRow("Static / Multiperiod", self.horizon_combo)
        form_layout.addRow("Region", self.region_combo)
        form_layout.addRow("Start date", self.start_date_edit)
        form_layout.addRow("Simulation duration", self.duration_input)
        form_layout.addRow("End date", self.end_date_label)
        form_layout.addRow("Graph resolution", self.resolution_combo)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.btn_select)
        button_layout.addWidget(self.btn_run)

        main_layout = QVBoxLayout()
        main_layout.addWidget(self.file_label)
        main_layout.addLayout(form_layout)
        main_layout.addLayout(button_layout)
        main_layout.addWidget(self.status)

        self.setLayout(main_layout)

        self.btn_select.clicked.connect(self.select_file)
        self.btn_run.clicked.connect(self.execute_program)
        self.start_date_edit.dateChanged.connect(self.update_duration_limit)
        self.start_date_edit.dateChanged.connect(self.update_end_date_label)
        self.duration_input.valueChanged.connect(self.update_end_date_label)

        self.update_duration_limit()
        self.update_end_date_label()

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

    def get_system_parameters_from_gui(self) -> dict:
        return {
            "VOLL (€/MWh)": self.voll_input.value(),
            "Static / Multiperiod": self.horizon_combo.currentText(),
            "Region": self.region_combo.currentText(),
            "Start date (dd/mm/aaaa)": self.start_date_edit.date().toPython(),
            "Simulation duration (days)": self.duration_input.value(),
            "Graph resolution": self.resolution_combo.currentText(),
        }

    def execute_program(self):
        valid, error_msg = self.validate_simulation_dates()
        if not valid:
            QMessageBox.warning(self, "Fechas no válidas", error_msg)
            return

        self.btn_run.setEnabled(False)
        self.status.setText("Estado: ejecutando...")

        system_parameters = self.get_system_parameters_from_gui()

        self.worker = Worker(
            input_path=self.input_path,
            system_parameters=system_parameters
        )
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_finished(self):
        self.status.setText("Estado: terminado")
        self.btn_run.setEnabled(True)
        QMessageBox.information(self, "Éxito", "Programa ejecutado correctamente.")

    def on_error(self, error_msg):
        self.status.setText("Estado: error")
        self.btn_run.setEnabled(True)
        QMessageBox.critical(self, "Error", f"Ha ocurrido un error:\n\n{error_msg}")
    
    def update_duration_limit(self):
        start_date = self.start_date_edit.date()
        max_end_date = self.max_end_date

        max_days = start_date.daysTo(max_end_date) + 1

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
            return False, "La fecha de inicio no puede ser anterior al 01/01/2022."

        if end_date > self.max_end_date:
            return False, "La fecha final de la simulación no puede superar el 31/12/2024."

        return True, ""

if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()