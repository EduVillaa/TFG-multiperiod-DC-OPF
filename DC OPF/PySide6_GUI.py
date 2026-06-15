from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QFileDialog,
    QVBoxLayout, QHBoxLayout, QMessageBox, QComboBox,
    QSpinBox, QDoubleSpinBox, QDateEdit, QFormLayout,
    QProgressBar, QTimeEdit, QTextEdit, QDoubleSpinBox, QCheckBox,
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
    

    return bool(
        df["Optimization mode"].isin(
            ["Optimize both", "Optimize MW", "Optimize MWh"]
        ).any()
    )


def MILPvsLP(filename):
    df = pd.read_excel(
        filename,
        sheet_name="Gen_Dispatchable",
        header=2
    )

    df.columns = df.columns.astype(str).str.strip()

    if "Committable" not in df.columns:
        raise KeyError(
            f"No existe la columna 'Committable'. Columnas disponibles: {list(df.columns)}"
        )

    committable = (
        df["Committable"]
        .fillna(False)
        .astype(str)
        .str.strip()
        .str.lower()
    )

    return bool(committable.isin(["true", "1", "yes", "y", "si", "sí"]).any())


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
        self.voll_input.setValue(1000)
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
        self.start_date_edit.setDate(QDate(2024, 1, 1))

        self.duration_input = QSpinBox()
        self.duration_input.setRange(1, 3650)
        self.duration_input.setValue(365)
        self.duration_input.setSuffix(" days")

        # Rolling horizon: fase 1, solo configuración desde la GUI.
        self.rolling_horizon_enabled_input = QCheckBox()
        self.rolling_horizon_enabled_input.setChecked(False)
        self.rolling_horizon_enabled_input.setToolTip(
            "Enable rolling horizon mode. In this phase it only stores the configuration."
        )

        self.rolling_horizon_days_input = QSpinBox()
        self.rolling_horizon_days_input.setRange(1, 7)
        self.rolling_horizon_days_input.setValue(3)
        self.rolling_horizon_days_input.setSuffix(" days")
        self.rolling_horizon_days_input.setToolTip(
            "Duration of each rolling horizon segment. Allowed range: 1 to 7 days."
        )

        # Rolling horizon: banda terminal para la trayectoria agregada de SOC hydro.
        self.rolling_hydro_soc_band_input = QDoubleSpinBox()
        self.rolling_hydro_soc_band_input.setRange(0.0, 30.0)
        self.rolling_hydro_soc_band_input.setDecimals(2)
        self.rolling_hydro_soc_band_input.setSingleStep(1.0)
        self.rolling_hydro_soc_band_input.setValue(0.5)
        self.rolling_hydro_soc_band_input.setSuffix(" %")
        self.rolling_hydro_soc_band_input.setToolTip(
            "Allowed band around the interpolated hydro SOC target in rolling horizon."
        )

        # Rolling horizon: valor residual de energía final en BatteryStore.
        self.rolling_batterystore_residual_value_input = QDoubleSpinBox()
        self.rolling_batterystore_residual_value_input.setRange(0.0, 1_000_000.0)
        self.rolling_batterystore_residual_value_input.setDecimals(2)
        self.rolling_batterystore_residual_value_input.setSingleStep(1.0)
        self.rolling_batterystore_residual_value_input.setValue(0.0)
        self.rolling_batterystore_residual_value_input.setSuffix(" €/MWh")
        self.rolling_batterystore_residual_value_input.setToolTip(
            "Residual value subtracted from the rolling horizon objective for final BatteryStore SOC."
        )

        # Rolling horizon: SOC mínimo terminal individual para BatteryStore.
        self.rolling_batterystore_min_final_soc_input = QDoubleSpinBox()
        self.rolling_batterystore_min_final_soc_input.setRange(0.0, 95.0)
        self.rolling_batterystore_min_final_soc_input.setDecimals(2)
        self.rolling_batterystore_min_final_soc_input.setSingleStep(1.0)
        self.rolling_batterystore_min_final_soc_input.setValue(0.0)
        self.rolling_batterystore_min_final_soc_input.setSuffix(" %")
        self.rolling_batterystore_min_final_soc_input.setToolTip(
            "Minimum final SOC required for each BatteryStore at the end of every rolling window."
        )

        # Simulación anual sin rolling: restricciones terminales intermedias hydro/PHS.
        self.intermediate_storage_constraints_enabled_input = QCheckBox()
        self.intermediate_storage_constraints_enabled_input.setChecked(False)
        self.intermediate_storage_constraints_enabled_input.setToolTip(
            "Apply hydro/PHS terminal constraints at intermediate points without enabling rolling horizon."
        )

        self.intermediate_storage_constraint_days_input = QSpinBox()
        self.intermediate_storage_constraint_days_input.setRange(1, 7)
        self.intermediate_storage_constraint_days_input.setValue(3)
        self.intermediate_storage_constraint_days_input.setSuffix(" days")
        self.intermediate_storage_constraint_days_input.setToolTip(
            "Duration of each intermediate block. The optimization is still solved once."
        )

        self.intermediate_hydro_soc_band_input = QDoubleSpinBox()
        self.intermediate_hydro_soc_band_input.setRange(0.0, 30.0)
        self.intermediate_hydro_soc_band_input.setDecimals(2)
        self.intermediate_hydro_soc_band_input.setSingleStep(0.1)
        self.intermediate_hydro_soc_band_input.setValue(0.5)
        self.intermediate_hydro_soc_band_input.setSuffix(" %")
        self.intermediate_hydro_soc_band_input.setToolTip(
            "Allowed band around the interpolated hydro SOC target at intermediate points."
        )

        self.resolution_combo = QComboBox()
        self.resolution_combo.addItems(["Auto", "Hourly", "Daily", "Weekly"])
        self.resolution_combo.setCurrentText("Auto")

        self.end_date_label = QLabel("End date: -")

        # Campos para optimización de batería
        self.discount_rate_input = QDoubleSpinBox()
        self.discount_rate_input.setRange(0.0, 100.0)
        self.discount_rate_input.setDecimals(2)
        self.discount_rate_input.setValue(4.0)
        self.discount_rate_input.setSuffix(" %")

        self.default_battery_lifetime_input = QSpinBox()
        self.default_battery_lifetime_input.setRange(1, 100)
        self.default_battery_lifetime_input.setValue(15)
        self.default_battery_lifetime_input.setSuffix(" years")

        # Solver options
        self.solver_combo = QComboBox()
        self.solver_combo.addItems(["HiGHS", "Gurobi"])
        self.solver_combo.setCurrentText("Gurobi")
        self.solver_combo.setToolTip("Solver used by PyPSA/Linopy to solve the OPF problem.")

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
        self.time_limit_input.setValue(0)
        self.time_limit_input.setSuffix(" s")
        self.time_limit_input.setToolTip(
            "Solver time limit in seconds. 0 = no time limit. 3600 = 1 hour."
        )

        self.gurobi_threads_input = QSpinBox()
        self.gurobi_threads_input.setRange(0, 128)
        self.gurobi_threads_input.setValue(10)
        self.gurobi_threads_input.setToolTip(
            "Number of threads used by Gurobi. 0 = automatic."
        )

        self.gurobi_mip_focus_input = QComboBox()
        self.gurobi_mip_focus_input.addItems([
            "0 - Automatic",
            "1 - Find feasible solutions",
            "2 - Prove optimality",
            "3 - Improve bound",
        ])
        self.gurobi_mip_focus_input.setCurrentText("1 - Find feasible solutions")
        self.gurobi_mip_focus_input.setToolTip(
            "Gurobi MIPFocus. 1 is useful when the solver struggles to find a feasible solution."
        )

        self.gurobi_method_input = QComboBox()
        self.gurobi_method_input.addItems([
            "-1 - Automatic",
            "0 - Primal simplex",
            "1 - Dual simplex",
            "2 - Barrier",
            "3 - Concurrent",
        ])
        self.gurobi_method_input.setCurrentText("2 - Barrier")
        self.gurobi_method_input.setToolTip(
            "Gurobi Method parameter for continuous relaxations. "
            "3 = concurrent optimizer, useful when the log reports high concurrent spin time."
        )

        self.gurobi_crossover_input = QComboBox()
        self.gurobi_crossover_input.addItems([
            "-1 - Automatic",
            "0 - Disabled",
            "1 - Push",
            "2 - Automatic crossover",
        ])
        self.gurobi_crossover_input.setCurrentText("-1 - Automatic")
        self.gurobi_crossover_input.setToolTip(
            "Gurobi Crossover parameter. "
            "For large LP problems solved with barrier, setting Crossover = 0 can avoid spending a long time in crossover."
        )

        self.gurobi_numeric_focus_input = QSpinBox()
        self.gurobi_numeric_focus_input.setRange(0, 3)
        self.gurobi_numeric_focus_input.setValue(2)
        self.gurobi_numeric_focus_input.setToolTip(
            "Gurobi NumericFocus. Higher values make the solver more careful with numerical issues. "
            "0 = automatic/default, 1-3 = increasing numerical caution."
        )

        self.gurobi_bar_conv_tol_input = QComboBox()
        self.gurobi_bar_conv_tol_input.addItems([
            "None",
            "1e-10",
            "1e-9",
            "1e-8",
            "1e-7",
            "1e-6",
        ])
        self.gurobi_bar_conv_tol_input.setCurrentText("1e-8")
        self.gurobi_bar_conv_tol_input.setToolTip(
            "Gurobi BarConvTol. Barrier convergence tolerance. Mainly useful for large LPs solved with barrier."
        )

        self.gurobi_bar_homogeneous_input = QComboBox()
        self.gurobi_bar_homogeneous_input.addItems([
            "-1 - Automatic",
            "0 - Off",
            "1 - On",
        ])
        self.gurobi_bar_homogeneous_input.setCurrentText("1 - On")
        self.gurobi_bar_homogeneous_input.setToolTip(
            "Gurobi BarHomogeneous. Homogeneous barrier algorithm. "
            "Useful for numerically difficult LPs solved with barrier."
        )

        
        self.gurobi_feasibility_tol_input = QComboBox()
        self.gurobi_feasibility_tol_input.addItems([
            "None",
            "1e-9",
            "1e-8",
            "1e-7",
            "1e-6",
            "1e-5",
        ])
        self.gurobi_feasibility_tol_input.setCurrentText("1e-6")
        self.gurobi_feasibility_tol_input.setToolTip(
            "Gurobi FeasibilityTol. Constraint feasibility tolerance."
        )

        self.gurobi_optimality_tol_input = QComboBox()
        self.gurobi_optimality_tol_input.addItems([
            "None",
            "1e-9",
            "1e-8",
            "1e-7",
            "1e-6",
            "1e-5",
        ])
        self.gurobi_optimality_tol_input.setCurrentText("1e-6")
        self.gurobi_optimality_tol_input.setToolTip(
            "Gurobi OptimalityTol. Optimality tolerance for continuous optimization."
        )

        # Notes

        self.notes_input = QTextEdit()
        self.notes_input.setPlaceholderText(
            "Write notes for this simulation, e.g. Q1 2024, no UC, line penalty 0.05, battery CAPEX sweep..."
        )
        self.notes_input.setFixedHeight(80)
        self.notes_input.setToolTip(
            "Optional notes saved in the output Excel file to identify this simulation."
        )

        # Line flow penalty
        self.line_flow_penalty_input = QDoubleSpinBox()
        self.line_flow_penalty_input.setRange(0.0, 1000.0)
        self.line_flow_penalty_input.setDecimals(4)
        self.line_flow_penalty_input.setSingleStep(0.01)
        self.line_flow_penalty_input.setValue(0)
        self.line_flow_penalty_input.setSuffix(" €/MWh")
        self.line_flow_penalty_input.setToolTip(
            "Penalty applied to the absolute value of AC line flows. "
            "Set to 0 to disable line flow penalty."
        )

        self.line_flow_length_scaling_input = QCheckBox()
        self.line_flow_length_scaling_input.setChecked(True)
        self.line_flow_length_scaling_input.setToolTip(
            "If enabled, the line flow penalty is scaled according to the geographical length of each line."
        )

        self.form_layout = QFormLayout()
        form_layout = self.form_layout

        self.voll_row_label = QLabel("VOLL")
        self.line_flow_penalty_row_label = QLabel("Line flow penalty")
        self.line_flow_length_scaling_row_label = QLabel("Scale penalty by line length")
        self.horizon_row_label = QLabel("Static / Multiperiod")
        self.start_date_row_label = QLabel("Start date")
        self.duration_row_label = QLabel("Simulation duration")
        self.rolling_horizon_enabled_row_label = QLabel("Enable rolling horizon")
        self.rolling_horizon_days_row_label = QLabel("Rolling horizon segment duration")
        self.rolling_hydro_soc_band_row_label = QLabel("Hydro trajectory margin (%)")
        self.rolling_batterystore_residual_value_row_label = QLabel(
            "BatteryStore residual value [€/MWh]"
        )
        self.rolling_batterystore_min_final_soc_row_label = QLabel(
            "BatteryStore minimum final SOC (%)"
        )
        self.intermediate_storage_constraints_enabled_row_label = QLabel(
            "Apply intermediate hydro/PHS terminal constraints"
        )
        self.intermediate_storage_constraint_days_row_label = QLabel(
            "Intermediate hydro/PHS block duration"
        )
        self.intermediate_hydro_soc_band_row_label = QLabel(
            "Intermediate hydro trajectory margin (%)"
        )
        self.static_date_row_label = QLabel("Static snapshot date")
        self.static_hour_row_label = QLabel("Static snapshot hour")
        self.end_date_row_label = QLabel("End date")
        self.resolution_row_label = QLabel("Graph resolution")
        self.discount_rate_row_label = QLabel("Discount rate")
        self.default_battery_lifetime_row_label = QLabel("Default battery lifetime")


        self.solver_row_label = QLabel("Solver")

        self.problem_type_row_label = QLabel("Detected problem")
        self.problem_type_label = QLabel("-")

        self.mip_rel_gap_row_label = QLabel("MIP relative gap (%)")
        self.time_limit_row_label = QLabel("Solver time limit")
        self.gurobi_threads_row_label = QLabel("Gurobi threads")
        self.gurobi_mip_focus_row_label = QLabel("Gurobi MIPFocus")
        self.gurobi_method_row_label = QLabel("Gurobi Method")
        self.gurobi_bar_homogeneous_row_label = QLabel("Gurobi BarHomogeneous")
        self.gurobi_crossover_row_label = QLabel("Gurobi Crossover")
        self.gurobi_numeric_focus_row_label = QLabel("Gurobi NumericFocus")
        self.gurobi_bar_conv_tol_row_label = QLabel("Gurobi BarConvTol")
        self.gurobi_feasibility_tol_row_label = QLabel("Gurobi FeasibilityTol")
        self.gurobi_optimality_tol_row_label = QLabel("Gurobi OptimalityTol")
        self.notes_row_label = QLabel("Notes")

        form_layout.addRow(self.voll_row_label, self.voll_input)
        form_layout.addRow(self.line_flow_penalty_row_label, self.line_flow_penalty_input)
        form_layout.addRow(self.line_flow_length_scaling_row_label, self.line_flow_length_scaling_input)
        form_layout.addRow(self.horizon_row_label, self.horizon_combo)
        form_layout.addRow(self.start_date_row_label, self.start_date_edit)
        form_layout.addRow(self.static_date_row_label, self.static_date_edit)
        form_layout.addRow(self.static_hour_row_label, self.static_hour_edit)
        form_layout.addRow(self.duration_row_label, self.duration_input)
        form_layout.addRow(
            self.rolling_horizon_enabled_row_label,
            self.rolling_horizon_enabled_input
        )
        form_layout.addRow(
            self.rolling_horizon_days_row_label,
            self.rolling_horizon_days_input
        )
        form_layout.addRow(
            self.rolling_hydro_soc_band_row_label,
            self.rolling_hydro_soc_band_input
        )
        form_layout.addRow(
            self.rolling_batterystore_residual_value_row_label,
            self.rolling_batterystore_residual_value_input
        )
        form_layout.addRow(
            self.rolling_batterystore_min_final_soc_row_label,
            self.rolling_batterystore_min_final_soc_input
        )
        form_layout.addRow(
            self.intermediate_storage_constraints_enabled_row_label,
            self.intermediate_storage_constraints_enabled_input
        )
        form_layout.addRow(
            self.intermediate_storage_constraint_days_row_label,
            self.intermediate_storage_constraint_days_input
        )
        form_layout.addRow(
            self.intermediate_hydro_soc_band_row_label,
            self.intermediate_hydro_soc_band_input
        )
        form_layout.addRow(self.end_date_row_label, self.end_date_label)
        form_layout.addRow(self.resolution_row_label, self.resolution_combo)

        form_layout.addRow(self.solver_row_label, self.solver_combo)
        form_layout.addRow(self.problem_type_row_label, self.problem_type_label)
        
        form_layout.addRow(self.mip_rel_gap_row_label, self.mip_rel_gap_input)
        form_layout.addRow(self.time_limit_row_label, self.time_limit_input)
        form_layout.addRow(self.gurobi_threads_row_label, self.gurobi_threads_input)
        form_layout.addRow(self.gurobi_mip_focus_row_label, self.gurobi_mip_focus_input)
        form_layout.addRow(self.gurobi_method_row_label, self.gurobi_method_input)
        form_layout.addRow(
            self.gurobi_bar_homogeneous_row_label,
            self.gurobi_bar_homogeneous_input
        )
        form_layout.addRow(self.gurobi_crossover_row_label, self.gurobi_crossover_input)

        form_layout.addRow(
        self.gurobi_numeric_focus_row_label,
        self.gurobi_numeric_focus_input
        )

        form_layout.addRow(
            self.gurobi_bar_conv_tol_row_label,
            self.gurobi_bar_conv_tol_input
        )

        form_layout.addRow(
            self.gurobi_feasibility_tol_row_label,
            self.gurobi_feasibility_tol_input
        )

        form_layout.addRow(
            self.gurobi_optimality_tol_row_label,
            self.gurobi_optimality_tol_input
        )

        form_layout.addRow(self.discount_rate_row_label, self.discount_rate_input)
        form_layout.addRow(
            self.default_battery_lifetime_row_label,
            self.default_battery_lifetime_input
        )

        form_layout.addRow(self.notes_row_label, self.notes_input)

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
        self.rolling_horizon_enabled_input.toggled.connect(self.update_dynamic_fields_visibility)
        self.intermediate_storage_constraints_enabled_input.toggled.connect(self.update_dynamic_fields_visibility)
        self.solver_combo.currentTextChanged.connect(self.update_dynamic_fields_visibility)
        self.gurobi_method_input.currentTextChanged.connect(self.update_dynamic_fields_visibility)

        self.update_duration_limit()
        self.update_end_date_label()
        self.update_dynamic_fields_visibility()

    def get_selected_input_file(self) -> str:
        return self.input_path if self.input_path else "GridInputs.xlsx"

    def battery_optimization_enabled(self) -> bool:
        try:
    
            return bool(getBatteryOptimizationMode(self.get_selected_input_file()))
        except Exception:
            print("no se consiguió obtener los datos de optimización de la batería")
            return False
    
    def problem_is_milp(self) -> bool:
        try:
            return bool(MILPvsLP(self.get_selected_input_file()))
        except Exception as e:
            print(f"No se consiguió detectar si el problema es MILP o LP: {e}")
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
        
        def combo_float_or_none(combo):
            text = combo.currentText().strip()
            if text.lower() in ["none", "", "default", "auto"]:
                return None
            return float(text)
        
        horizon = self.horizon_combo.currentText()

        solver_display_name = self.solver_combo.currentText()

        if solver_display_name == "HiGHS":
            solver_name = "highs"
        elif solver_display_name == "Gurobi":
            solver_name = "gurobi"
        else:
            raise ValueError(f"Solver no soportado: {solver_display_name}")

        is_milp = self.problem_is_milp()
        is_lp = not is_milp
        battery_opt_enabled = self.battery_optimization_enabled()

        # Rolling horizon: parámetros opcionales. Si está desactivado,
        # el flujo de simulación individual conserva el comportamiento previo.
        rolling_horizon_enabled = (
            horizon == "Multiperiod"
            and not battery_opt_enabled
            and self.rolling_horizon_enabled_input.isChecked()
        )
        intermediate_storage_constraints_enabled = (
            horizon == "Multiperiod"
            and not rolling_horizon_enabled
            and self.intermediate_storage_constraints_enabled_input.isChecked()
        )

        params = {
            "VOLL (€/MWh)": self.voll_input.value(),
            "line_flow_penalty": self.line_flow_penalty_input.value(),
            "use_line_length_scaling": self.line_flow_length_scaling_input.isChecked(),
            "Static / Multiperiod": horizon,
            "solver_name": solver_name,
            "time_limit": self.time_limit_input.value(),
            "Notes": self.notes_input.toPlainText().strip(),
            "problem_type": "MILP" if is_milp else "LP",
            "rolling_horizon_enabled": rolling_horizon_enabled,
            "rolling_horizon_days": self.rolling_horizon_days_input.value(),
            "rolling_hydro_soc_band_percent": self.rolling_hydro_soc_band_input.value(),
            "rolling_batterystore_residual_value_eur_per_mwh": (
                self.rolling_batterystore_residual_value_input.value()
            ),
            "rolling_batterystore_min_final_soc_percent": (
                self.rolling_batterystore_min_final_soc_input.value()
            ),
            "intermediate_storage_constraints_enabled": intermediate_storage_constraints_enabled,
            "intermediate_storage_constraint_days": (
                self.intermediate_storage_constraint_days_input.value()
            ),
            "intermediate_hydro_soc_band_percent": (
                self.intermediate_hydro_soc_band_input.value()
            ),
        }

        if is_milp:
            params["mip_rel_gap"] = self.mip_rel_gap_input.value() / 100
        else:
            params["mip_rel_gap"] = None

        if solver_name == "gurobi":
            method_text = self.gurobi_method_input.currentText()
            bar_homogeneous_text = self.gurobi_bar_homogeneous_input.currentText()
            crossover_text = self.gurobi_crossover_input.currentText()

            params.update({
            "threads": self.gurobi_threads_input.value(),
            "numeric_focus": self.gurobi_numeric_focus_input.value(),
            "feasibility_tol": combo_float_or_none(self.gurobi_feasibility_tol_input),
            "optimality_tol": combo_float_or_none(self.gurobi_optimality_tol_input),})

            if is_milp:
                mip_focus_text = self.gurobi_mip_focus_input.currentText()
                params["mip_focus"] = int(mip_focus_text.split(" - ")[0])
            else:
                params["mip_focus"] = None

            if is_lp:
                params.update({
                "method": int(method_text.split(" - ")[0]),
                "bar_homogeneous": int(bar_homogeneous_text.split(" - ")[0]),
                "crossover": int(crossover_text.split(" - ")[0]),
                "bar_conv_tol": combo_float_or_none(self.gurobi_bar_conv_tol_input),
                })
            else:
                params.update({
                    "method": None,
                    "crossover": None,
                    "bar_conv_tol": None,
                })

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

            valid, error_msg = self.validate_rolling_horizon_settings()
            if not valid:
                QMessageBox.warning(self, "Rolling horizon no válido", error_msg)
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

    def validate_rolling_horizon_settings(self) -> tuple[bool, str]:
        # Rolling/intermediate: validación defensiva además de los rangos de la GUI.
        if self.horizon_combo.currentText() != "Multiperiod":
            return True, ""

        rolling_enabled = self.rolling_horizon_enabled_input.isChecked()
        intermediate_enabled = self.intermediate_storage_constraints_enabled_input.isChecked()

        if rolling_enabled and intermediate_enabled:
            return (
                False,
                "Las restricciones intermedias hydro/PHS solo aplican sin rolling horizon.",
            )

        if rolling_enabled:
            rolling_days = self.rolling_horizon_days_input.value()

            if not isinstance(rolling_days, int):
                return False, "La duración del rolling horizon debe ser un número entero de días."

            if rolling_days < 1:
                return False, "La duración mínima del rolling horizon es 1 día."

            if rolling_days > 7:
                return False, "La duración máxima del rolling horizon es 7 días."

            hydro_band = self.rolling_hydro_soc_band_input.value()

            if hydro_band < 0:
                return False, "El margen hydro del rolling horizon no puede ser negativo."

            if hydro_band > 30:
                return False, "El margen hydro máximo del rolling horizon es 30%."

            batterystore_residual_value = (
                self.rolling_batterystore_residual_value_input.value()
            )

            if batterystore_residual_value < 0:
                return False, "El valor residual BatteryStore no puede ser negativo."

            batterystore_min_final_soc = self.rolling_batterystore_min_final_soc_input.value()

            if batterystore_min_final_soc < 0:
                return False, "El SOC final mínimo BatteryStore no puede ser negativo."

            if batterystore_min_final_soc > 95:
                return False, "El SOC final mínimo BatteryStore no puede superar el 95%."

        if intermediate_enabled:
            intermediate_days = self.intermediate_storage_constraint_days_input.value()

            if not isinstance(intermediate_days, int):
                return False, "La duración de los bloques intermedios debe ser un entero."

            if intermediate_days < 1:
                return False, "La duración mínima de los bloques intermedios es 1 día."

            if intermediate_days > 7:
                return False, "La duración máxima de los bloques intermedios es 7 días."

            intermediate_hydro_band = self.intermediate_hydro_soc_band_input.value()

            if intermediate_hydro_band < 0:
                return False, "El margen hydro intermedio no puede ser negativo."

            if intermediate_hydro_band > 30:
                return False, "El margen hydro intermedio máximo es 30%."

        return True, ""
    

    def set_form_row_visible(self, label_widget, field_widget, visible: bool):
        row, role = self.form_layout.getWidgetPosition(label_widget)

        if row == -1:
            label_widget.setVisible(visible)
            field_widget.setVisible(visible)
            return

        try:
            self.form_layout.setRowVisible(row, visible)
        except AttributeError:
            label_widget.setVisible(visible)
            field_widget.setVisible(visible)
            self.form_layout.invalidate()

    def update_dynamic_fields_visibility(self):
        horizon = self.horizon_combo.currentText()

        is_multiperiod = horizon == "Multiperiod"
        is_static = horizon == "Static"

        battery_opt_enabled = self.battery_optimization_enabled()
        show_rolling_horizon = is_multiperiod and not battery_opt_enabled

        if not show_rolling_horizon and self.rolling_horizon_enabled_input.isChecked():
            self.rolling_horizon_enabled_input.blockSignals(True)
            self.rolling_horizon_enabled_input.setChecked(False)
            self.rolling_horizon_enabled_input.blockSignals(False)

        rolling_horizon_enabled = (
            show_rolling_horizon
            and self.rolling_horizon_enabled_input.isChecked()
        )
        show_intermediate_constraints = is_multiperiod and not rolling_horizon_enabled

        if rolling_horizon_enabled and self.intermediate_storage_constraints_enabled_input.isChecked():
            self.intermediate_storage_constraints_enabled_input.blockSignals(True)
            self.intermediate_storage_constraints_enabled_input.setChecked(False)
            self.intermediate_storage_constraints_enabled_input.blockSignals(False)

        intermediate_storage_constraints_enabled = (
            show_intermediate_constraints
            and self.intermediate_storage_constraints_enabled_input.isChecked()
        )
        show_battery_fields = is_multiperiod and battery_opt_enabled

        # Detectar solver y tipo de problema
        solver = self.solver_combo.currentText()
        is_gurobi = solver == "Gurobi"

        is_milp = self.problem_is_milp()
        is_lp = not is_milp

        # Detectar método de Gurobi
        method_text = self.gurobi_method_input.currentText()
        method_value = int(method_text.split(" - ")[0])
        is_barrier = method_value == 2

        # ---------------------------------------------------------
        # Campos propios de simulación multiperiodo
        # ---------------------------------------------------------
        self.set_form_row_visible(
            self.start_date_row_label,
            self.start_date_edit,
            is_multiperiod
        )

        self.set_form_row_visible(
            self.duration_row_label,
            self.duration_input,
            is_multiperiod
        )

        # Rolling horizon: se muestra solo en multiperiodo; la duración
        # solo aparece cuando el modo rolling horizon está activado.
        self.set_form_row_visible(
            self.rolling_horizon_enabled_row_label,
            self.rolling_horizon_enabled_input,
            show_rolling_horizon
        )

        self.set_form_row_visible(
            self.rolling_horizon_days_row_label,
            self.rolling_horizon_days_input,
            rolling_horizon_enabled
        )

        self.set_form_row_visible(
            self.rolling_hydro_soc_band_row_label,
            self.rolling_hydro_soc_band_input,
            rolling_horizon_enabled
        )

        self.set_form_row_visible(
            self.rolling_batterystore_residual_value_row_label,
            self.rolling_batterystore_residual_value_input,
            rolling_horizon_enabled
        )

        self.set_form_row_visible(
            self.rolling_batterystore_min_final_soc_row_label,
            self.rolling_batterystore_min_final_soc_input,
            rolling_horizon_enabled
        )

        self.set_form_row_visible(
            self.intermediate_storage_constraints_enabled_row_label,
            self.intermediate_storage_constraints_enabled_input,
            show_intermediate_constraints
        )

        self.set_form_row_visible(
            self.intermediate_storage_constraint_days_row_label,
            self.intermediate_storage_constraint_days_input,
            intermediate_storage_constraints_enabled
        )

        self.set_form_row_visible(
            self.intermediate_hydro_soc_band_row_label,
            self.intermediate_hydro_soc_band_input,
            intermediate_storage_constraints_enabled
        )

        self.set_form_row_visible(
            self.end_date_row_label,
            self.end_date_label,
            is_multiperiod
        )

        self.set_form_row_visible(
            self.resolution_row_label,
            self.resolution_combo,
            is_multiperiod
        )

        # ---------------------------------------------------------
        # Campos propios de simulación estática
        # ---------------------------------------------------------
        self.set_form_row_visible(
            self.static_date_row_label,
            self.static_date_edit,
            is_static
        )

        self.set_form_row_visible(
            self.static_hour_row_label,
            self.static_hour_edit,
            is_static
        )

        # ---------------------------------------------------------
        # Campos de optimización de batería
        # ---------------------------------------------------------
        self.set_form_row_visible(
            self.discount_rate_row_label,
            self.discount_rate_input,
            show_battery_fields
        )

        self.set_form_row_visible(
            self.default_battery_lifetime_row_label,
            self.default_battery_lifetime_input,
            show_battery_fields
        )

        # ---------------------------------------------------------
        # Opciones generales del solver
        # ---------------------------------------------------------
        self.set_form_row_visible(
            self.solver_row_label,
            self.solver_combo,
            True
        )

        self.set_form_row_visible(
            self.problem_type_row_label,
            self.problem_type_label,
            True
        )
        self.problem_type_label.setText("MILP" if is_milp else "LP")

        self.set_form_row_visible(
            self.time_limit_row_label,
            self.time_limit_input,
            True
        )

        # ---------------------------------------------------------
        # Opciones MIP
        # ---------------------------------------------------------
        self.set_form_row_visible(
            self.mip_rel_gap_row_label,
            self.mip_rel_gap_input,
            is_milp
        )

        self.set_form_row_visible(
            self.gurobi_mip_focus_row_label,
            self.gurobi_mip_focus_input,
            is_gurobi and is_milp
        )

        # ---------------------------------------------------------
        # Opciones Gurobi generales
        # ---------------------------------------------------------
        self.set_form_row_visible(
            self.gurobi_threads_row_label,
            self.gurobi_threads_input,
            is_gurobi
        )

        self.set_form_row_visible(
            self.gurobi_numeric_focus_row_label,
            self.gurobi_numeric_focus_input,
            is_gurobi
        )

        self.set_form_row_visible(
            self.gurobi_feasibility_tol_row_label,
            self.gurobi_feasibility_tol_input,
            is_gurobi
        )

        self.set_form_row_visible(
            self.gurobi_optimality_tol_row_label,
            self.gurobi_optimality_tol_input,
            is_gurobi
        )

        # ---------------------------------------------------------
        # Opciones Gurobi LP
        # ---------------------------------------------------------
        self.set_form_row_visible(
            self.gurobi_method_row_label,
            self.gurobi_method_input,
            is_gurobi and is_lp
        )

        # Solo tiene sentido si el LP se resuelve con barrier
        show_barrier_options = is_gurobi and is_lp and is_barrier

        self.set_form_row_visible(
            self.gurobi_crossover_row_label,
            self.gurobi_crossover_input,
            show_barrier_options
        )

        self.set_form_row_visible(
            self.gurobi_bar_conv_tol_row_label,
            self.gurobi_bar_conv_tol_input,
            show_barrier_options
        )

        self.set_form_row_visible(
            self.gurobi_bar_homogeneous_row_label,
            self.gurobi_bar_homogeneous_input,
            show_barrier_options
        )

        self.form_layout.invalidate()
        self.adjustSize()

if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()
