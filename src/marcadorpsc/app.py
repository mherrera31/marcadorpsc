import sys
# Changed from PyQt6 to PySide6 to work with Briefcase
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QMessageBox, QHBoxLayout, QMainWindow, QComboBox, QTableWidget, QTableWidgetItem, QHeaderView, QTabWidget, QFormLayout, QAbstractItemView, QCalendarWidget, QDialog, QDialogButtonBox, QFileDialog
from PySide6.QtCore import Qt, QTimer, QDate, QPoint, QTime
from PySide6.QtGui import QPainter, QColor, QFont
import psycopg2
import bcrypt
from datetime import datetime, date
import pytz
import openpyxl

# Database connection details
DB_HOST = "aws-1-us-east-1.pooler.supabase.com"
DB_NAME = "postgres"
DB_USER = "postgres.zupvdegnxrxtihapcuqr"
DB_PASS = "es7Wq7cb02QvsCLE"  # <--- Pega tu contraseña real aquí
DB_PORT = "5432"

def connect_to_db():
    """Connects to the Supabase database."""
    try:
        conn = psycopg2.connect(
            host=DB_HOST,
            database=DB_NAME,
            user=DB_USER,
            password=DB_PASS,
            port=DB_PORT
        )
        return conn
    except (Exception, psycopg2.Error) as error:
        print("Error al conectar a la base de datos:", error)
        return None

# --- Custom Analog Clock Widget ---

class AnalogClock(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(200, 200)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update)
        self.timer.start(1000)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        side = min(self.width(), self.height())
        painter.translate(self.width() / 2, self.height() / 2)
        painter.scale(side / 200.0, side / 200.0)

        panama_tz = pytz.timezone('America/Panama')
        current_time = datetime.now(panama_tz).time()

        # Draw clock face
        painter.setPen(Qt.GlobalColor.black)
        painter.setBrush(QColor(255, 255, 255))
        painter.drawEllipse(-99, -99, 198, 198)

        # Draw clock hands
        painter.setPen(Qt.GlobalColor.black)
        painter.setBrush(Qt.GlobalColor.black)
        
        painter.save()
        painter.rotate(6.0 * (current_time.hour % 12 * 5 + current_time.minute / 12.0))
        painter.drawConvexPolygon(QPoint(7, 0), QPoint(-7, 0), QPoint(0, -50))
        painter.restore()

        painter.save()
        painter.rotate(6.0 * (current_time.minute + current_time.second / 60.0))
        painter.drawConvexPolygon(QPoint(7, 0), QPoint(-7, 0), QPoint(0, -70))
        painter.restore()

        painter.save()
        painter.setPen(QColor(255, 0, 0))
        painter.rotate(6.0 * current_time.second)
        painter.drawConvexPolygon(QPoint(1, 0), QPoint(-1, 0), QPoint(0, -90))
        painter.restore()
        
        painter.setPen(Qt.GlobalColor.black)
        for i in range(12):
            painter.drawLine(0, -88, 0, -99)
            painter.rotate(30)

    def sizeHint(self):
        return self.minimumSizeHint()

# --- Main TimeClock Window ---

class TimeClockWindow(QMainWindow):
    def __init__(self, user_id, username):
        super().__init__()
        self.user_id = user_id
        self.username = username
        self.setWindowTitle("Marcación de Horas")
        self.setGeometry(100, 100, 400, 300)
        self.event_types = self.load_event_types()
        self.setup_ui()
        self.load_today_entries()

    def load_event_types(self):
        """Loads event types from the database."""
        connection = connect_to_db()
        if connection:
            cursor = connection.cursor()
            try:
                cursor.execute("SELECT id, event_name FROM event_types;")
                return {row[1]: row[0] for row in cursor.fetchall()}
            except (Exception, psycopg2.Error) as error:
                QMessageBox.critical(self, "Error de Base de Datos", f"Ocurrió un error al cargar tipos de eventos: {error}")
                return {}
            finally:
                cursor.close()
                connection.close()
        return {}
    
    def load_today_entries(self):
        """Loads and displays today's time entries for the current user."""
        connection = connect_to_db()
        if connection:
            cursor = connection.cursor()
            try:
                panama_tz = pytz.timezone('America/Panama')
                today = datetime.now(panama_tz).date()
                
                sql = """
                    SELECT et.event_name, te.timestamp::time
                    FROM time_entries te
                    JOIN event_types et ON te.event_type_id = et.id
                    WHERE te.user_id = %s AND DATE(te.timestamp) = %s
                    ORDER BY te.timestamp;
                """
                cursor.execute(sql, (self.user_id, today))
                entries = cursor.fetchall()

                self.today_entries_layout.setParent(None)
                self.today_entries_layout = QVBoxLayout()
                self.layout.addLayout(self.today_entries_layout)

                if entries:
                    for event_name, entry_time in entries:
                        formatted_time = entry_time.strftime("%I:%M:%S %p")
                        label_text = f"<b>{event_name.replace('_', ' ').title()}:</b> {formatted_time}"
                        label = QLabel(label_text)
                        label.setStyleSheet("font-size: 11pt;")
                        self.today_entries_layout.addWidget(label)
                else:
                    label = QLabel("No hay marcaciones para hoy.")
                    label.setStyleSheet("font-size: 11pt; font-style: italic;")
                    self.today_entries_layout.addWidget(label)

            except (Exception, psycopg2.Error) as error:
                QMessageBox.critical(self, "Error de Base de Datos", f"Ocurrió un error al cargar las marcaciones: {error}")
            finally:
                cursor.close()
                connection.close()


    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.layout = QVBoxLayout(central_widget)
        
        top_layout = QHBoxLayout()
        self.welcome_label = QLabel(f"Usuario: {self.username}")
        self.welcome_label.setStyleSheet("font-size: 12pt;")
        top_layout.addWidget(self.welcome_label)

        logout_button = QPushButton("Cerrar Sesión")
        logout_button.clicked.connect(self.logout)
        top_layout.addWidget(logout_button)
        self.layout.addLayout(top_layout)

        self.title_label = QLabel("Marcación de Horas")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title_label.setStyleSheet("font-size: 16pt; font-weight: bold;")
        self.layout.addWidget(self.title_label)
        
        # Center clock and date
        clock_layout = QVBoxLayout()
        self.analog_clock = AnalogClock(self)
        self.analog_clock.setFixedSize(200, 200)
        clock_layout.addWidget(self.analog_clock, alignment=Qt.AlignmentFlag.AlignCenter)

        self.date_label = QLabel()
        self.date_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.date_label.setStyleSheet("font-size: 14pt; font-weight: bold;")

        self.digital_clock_label = QLabel()
        self.digital_clock_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.digital_clock_label.setStyleSheet("font-size: 18pt; font-weight: bold;")

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_datetime)
        self.timer.start(1000)
        self.update_datetime()

        clock_layout.addWidget(self.date_label)
        clock_layout.addWidget(self.digital_clock_label)
        self.layout.addLayout(clock_layout)

        self.today_entries_layout = QVBoxLayout()
        self.layout.addLayout(self.today_entries_layout)

        button_layout = QHBoxLayout()
        
        custom_order = ['entrada', 'almuerzo_inicio', 'almuerzo_fin', 'salida']
        all_event_names = sorted(list(self.event_types.keys()))
        other_events = [name for name in all_event_names if name not in custom_order]
        sorted_events = custom_order[:-1] + other_events + [custom_order[-1]]

        for event_name in sorted_events:
            button = QPushButton(event_name.replace('_', ' ').title())
            button.clicked.connect(lambda _, name=event_name: self.handle_time_entry(name))
            button_layout.addWidget(button)
        
        self.layout.addLayout(button_layout)
        
    def update_datetime(self):
        panama_tz = pytz.timezone('America/Panama')
        panama_time = datetime.now(panama_tz)
        formatted_date = panama_time.strftime("%A, %d de %B del %Y")
        formatted_time = panama_time.strftime("%I:%M:%S %p")
        self.date_label.setText(formatted_date)
        self.digital_clock_label.setText(formatted_time)

    def handle_time_entry(self, entry_type):
        """Maneja los clics de los botones de marcación y guarda en la base de datos."""
        connection = connect_to_db()
        if connection:
            cursor = connection.cursor()
            try:
                event_type_id = self.event_types.get(entry_type)
                if not event_type_id:
                    QMessageBox.critical(self, "Error", "Tipo de evento no encontrado.")
                    return

                panama_tz = pytz.timezone('America/Panama')
                today = datetime.now(panama_tz).date()

                sql_check = "SELECT COUNT(*) FROM time_entries WHERE user_id = %s AND event_type_id = %s AND DATE(timestamp) = %s;"
                cursor.execute(sql_check, (self.user_id, event_type_id, today))
                if cursor.fetchone()[0] > 0:
                    QMessageBox.warning(self, "Advertencia", f"Ya has marcado '{entry_type}' hoy. No puedes marcarlo dos veces en el mismo día.")
                    return

                sql_insert = "INSERT INTO time_entries (user_id, event_type_id, timestamp) VALUES (%s, %s, %s);"
                panama_time = datetime.now(panama_tz)
                cursor.execute(sql_insert, (self.user_id, event_type_id, panama_time))
                connection.commit()
                self.load_today_entries() # Reload entries after successful marking
            except (Exception, psycopg2.Error) as error:
                QMessageBox.critical(self, "Error de Base de Datos", f"Ocurrió un error: {error}")
            finally:
                cursor.close()
                connection.close()
        else:
            QMessageBox.critical(self, "Error de Conexión", "No se pudo conectar a la base de datos.")

    def logout(self):
        self.hide()
        self.login_window = LoginWindow()
        self.login_window.show()

# --- Admin Window ---

class AdminWindow(QMainWindow):
    def __init__(self, user_id, username):
        super().__init__()
        self.user_id = user_id
        self.username = username
        self.setWindowTitle("Área de Administrador")
        self.setGeometry(100, 100, 800, 600)
        self.all_users = {}
        self.all_event_types = {}
        self.setup_ui()
        self.load_users()
        self.load_event_types()
        self.load_reports()

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        top_layout = QHBoxLayout()
        self.welcome_label = QLabel(f"Admin: {self.username}")
        self.welcome_label.setStyleSheet("font-size: 12pt;")
        top_layout.addWidget(self.welcome_label)

        logout_button = QPushButton("Cerrar Sesión")
        logout_button.clicked.connect(self.logout)
        top_layout.addWidget(logout_button)
        layout.addLayout(top_layout)

        self.title_label = QLabel("Panel de Administración")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title_label.setStyleSheet("font-size: 16pt; font-weight: bold;")
        layout.addWidget(self.title_label)
        
        # Admin Portal using QTabWidget
        self.admin_tabs = QTabWidget()
        
        # User Management Tab
        self.user_management_tab = QWidget()
        self.setup_user_management_tab()
        self.admin_tabs.addTab(self.user_management_tab, "Gestión de Usuarios")

        # Report Tab
        self.report_tab = QWidget()
        self.setup_report_tab()
        self.admin_tabs.addTab(self.report_tab, "Reportes")
        
        # Event Types Tab
        self.event_types_tab = QWidget()
        self.setup_event_types_tab()
        self.admin_tabs.addTab(self.event_types_tab, "Tipos de Eventos")

        layout.addWidget(self.admin_tabs)

    def setup_user_management_tab(self):
        tab_layout = QHBoxLayout(self.user_management_tab)
        
        # Table of users
        users_table_widget = QWidget()
        users_table_layout = QVBoxLayout(users_table_widget)
        users_table_layout.addWidget(QLabel("Lista de Usuarios:"))
        self.users_table = QTableWidget()
        self.users_table.setColumnCount(2)
        self.users_table.setHorizontalHeaderLabels(["Usuario", "Rol"])
        header = self.users_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.users_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.users_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.users_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        users_table_layout.addWidget(self.users_table)
        
        # User management controls
        management_controls_widget = QWidget()
        management_controls_layout = QVBoxLayout(management_controls_widget)

        # Create user
        management_controls_layout.addWidget(QLabel("Crear Nuevo Usuario:"))
        self.new_user_entry = QLineEdit()
        self.new_user_entry.setPlaceholderText("Nombre de Usuario")
        management_controls_layout.addWidget(self.new_user_entry)
        self.new_pass_entry = QLineEdit()
        self.new_pass_entry.setPlaceholderText("Contraseña")
        self.new_pass_entry.setEchoMode(QLineEdit.EchoMode.Password)
        management_controls_layout.addWidget(self.new_pass_entry)
        self.role_label = QLabel("Rol:")
        management_controls_layout.addWidget(self.role_label)
        self.role_combo = QComboBox()
        self.role_combo.addItem("employee")
        self.role_combo.addItem("admin")
        management_controls_layout.addWidget(self.role_combo)
        self.create_user_button = QPushButton("Crear Usuario")
        self.create_user_button.clicked.connect(self.create_new_user)
        management_controls_layout.addWidget(self.create_user_button)

        # Password change and delete user
        management_controls_layout.addWidget(QLabel("Cambiar/Eliminar Usuario:"))
        self.pass_change_user_entry = QLineEdit()
        self.pass_change_user_entry.setPlaceholderText("Usuario seleccionado")
        self.pass_change_user_entry.setReadOnly(True)
        management_controls_layout.addWidget(self.pass_change_user_entry)
        self.new_pass_change_entry = QLineEdit()
        self.new_pass_change_entry.setPlaceholderText("Nueva Contraseña")
        self.new_pass_change_entry.setEchoMode(QLineEdit.EchoMode.Password)
        management_controls_layout.addWidget(self.new_pass_change_entry)
        self.change_pass_button = QPushButton("Cambiar Contraseña")
        self.change_pass_button.clicked.connect(self.change_user_password)
        management_controls_layout.addWidget(self.change_pass_button)

        self.delete_user_button = QPushButton("Eliminar Usuario")
        self.delete_user_button.setStyleSheet("background-color: #f44336; color: white;")
        self.delete_user_button.clicked.connect(self.delete_user)
        management_controls_layout.addWidget(self.delete_user_button)

        tab_layout.addWidget(users_table_widget)
        tab_layout.addWidget(management_controls_widget)
        
        self.users_table.itemSelectionChanged.connect(self.on_user_selected)
        
    def on_user_selected(self):
        selected_items = self.users_table.selectedItems()
        if selected_items:
            username = selected_items[0].text()
            self.pass_change_user_entry.setText(username)

    def setup_report_tab(self):
        tab_layout = QVBoxLayout(self.report_tab)
        tab_layout.addWidget(QLabel("Reportes de Horas:"))
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Usuario:"))
        self.user_filter_combo = QComboBox()
        filter_layout.addWidget(self.user_filter_combo)
        filter_layout.addWidget(QLabel("Desde:"))
        self.start_date_entry = QLineEdit()
        self.start_date_entry.setPlaceholderText("YYYY-MM-DD")
        filter_layout.addWidget(self.start_date_entry)
        filter_layout.addWidget(QLabel("Hasta:"))
        self.end_date_entry = QLineEdit()
        self.end_date_entry.setPlaceholderText("YYYY-MM-DD")
        filter_layout.addWidget(self.end_date_entry)
        self.filter_button = QPushButton("Filtrar")
        self.filter_button.clicked.connect(self.load_reports)
        filter_layout.addWidget(self.filter_button)
        self.export_button = QPushButton("Exportar a .xlsx")
        self.export_button.clicked.connect(self.export_report)
        self.export_button.setEnabled(False)
        filter_layout.addWidget(self.export_button)
        tab_layout.addLayout(filter_layout)
        self.report_table = QTableWidget()
        header = self.report_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        tab_layout.addWidget(self.report_table)

    def setup_event_types_tab(self):
        tab_layout = QVBoxLayout(self.event_types_tab)
        tab_layout.addWidget(QLabel("Gestión de Tipos de Eventos:"))

        form_layout = QFormLayout()
        self.event_name_entry = QLineEdit()
        form_layout.addRow("Nombre del Evento:", self.event_name_entry)

        add_event_button = QPushButton("Agregar Evento")
        add_event_button.clicked.connect(self.add_event_type)
        form_layout.addRow(add_event_button)

        tab_layout.addLayout(form_layout)
        
        self.event_types_list = QComboBox()
        tab_layout.addWidget(self.event_types_list)
        
        delete_event_button = QPushButton("Eliminar Evento Seleccionado")
        delete_event_button.clicked.connect(self.delete_event_type)
        tab_layout.addWidget(delete_event_button)
        
    def load_users(self):
        connection = connect_to_db()
        if connection:
            cursor = connection.cursor()
            try:
                cursor.execute("SELECT id, username, role FROM users ORDER BY username;")
                users_data = cursor.fetchall()
                self.users_table.setRowCount(len(users_data))
                self.users_table.setColumnCount(2)
                for row_idx, row_data in enumerate(users_data):
                    self.users_table.setItem(row_idx, 0, QTableWidgetItem(row_data[1]))
                    self.users_table.setItem(row_idx, 1, QTableWidgetItem(row_data[2]))
                
                self.all_users = {row[1]: row[0] for row in users_data}
                self.user_filter_combo.clear()
                self.user_filter_combo.addItem("Todos los empleados")
                for username in self.all_users.keys():
                    self.user_filter_combo.addItem(username)

            except (Exception, psycopg2.Error) as error:
                QMessageBox.critical(self, "Error de Base de Datos", f"Ocurrió un error al cargar usuarios: {error}")
            finally:
                cursor.close()
                connection.close()

    def load_event_types(self):
        connection = connect_to_db()
        if connection:
            cursor = connection.cursor()
            try:
                cursor.execute("SELECT id, event_name FROM event_types ORDER BY event_name;")
                self.all_event_types = {row[1]: row[0] for row in cursor.fetchall()}
                self.event_types_list.clear()
                for event_name in self.all_event_types.keys():
                    self.event_types_list.addItem(event_name)
            except (Exception, psycopg2.Error) as error:
                QMessageBox.critical(self, "Error de Base de Datos", f"Ocurrió un error al cargar tipos de eventos: {error}")
            finally:
                cursor.close()
                connection.close()

    def add_event_type(self):
        event_name = self.event_name_entry.text().lower().replace(" ", "_")
        if not event_name:
            QMessageBox.warning(self, "Error", "El nombre del evento no puede estar vacío.")
            return

        connection = connect_to_db()
        if connection:
            cursor = connection.cursor()
            try:
                sql = "INSERT INTO event_types (event_name) VALUES (%s);"
                cursor.execute(sql, (event_name,))
                connection.commit()
                self.event_name_entry.clear()
                self.load_event_types()
            except psycopg2.IntegrityError:
                QMessageBox.warning(self, "Error", "El tipo de evento ya existe.")
            except (Exception, psycopg2.Error) as error:
                QMessageBox.critical(self, "Error de Base de Datos", f"Ocurrió un error: {error}")
            finally:
                cursor.close()
                connection.close()
        else:
            QMessageBox.critical(self, "Error de Conexión", "No se pudo conectar a la base de datos.")

    def delete_event_type(self):
        event_name = self.event_types_list.currentText()
        event_id = self.all_event_types.get(event_name)
        
        if not event_id:
            QMessageBox.warning(self, "Error", "Por favor, selecciona un tipo de evento para eliminar.")
            return
        
        reply = QMessageBox.question(self, 'Confirmar Eliminación', 
                                    f"¿Estás seguro de que quieres eliminar el evento '{event_name}'?\nEsto podría afectar los registros de marcación existentes.", 
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            connection = connect_to_db()
            if connection:
                cursor = connection.cursor()
                try:
                    sql = "DELETE FROM event_types WHERE id = %s;"
                    cursor.execute(sql, (event_id,))
                    connection.commit()
                    self.load_event_types()
                except psycopg2.errors.ForeignKeyViolation:
                    QMessageBox.critical(self, "Error", "No se puede eliminar este tipo de evento porque hay marcaciones de tiempo que lo utilizan.")
                except (Exception, psycopg2.Error) as error:
                    QMessageBox.critical(self, "Error de Base de Datos", f"Ocurrió un error: {error}")
                finally:
                    cursor.close()
                    connection.close()
            else:
                QMessageBox.critical(self, "Error de Conexión", "No se pudo conectar a la base de datos.")

    def load_reports(self):
        connection = connect_to_db()
        if not connection:
            QMessageBox.critical(self, "Error de Conexión", "No se pudo conectar a la base de datos.")
            return
        
        cursor = connection.cursor()
        try:
            query = """
                SELECT u.username, DATE(te.timestamp), et.event_name, te.timestamp::time
                FROM time_entries te
                JOIN users u ON te.user_id = u.id
                JOIN event_types et ON te.event_type_id = et.id
                WHERE 1=1
            """
            params = []

            # Filter by user
            selected_user = self.user_filter_combo.currentText()
            if selected_user and selected_user != "Todos los empleados":
                query += " AND u.username = %s"
                params.append(selected_user)
            
            # Filter by date range
            start_date = self.start_date_entry.text()
            end_date = self.end_date_entry.text()
            if start_date:
                query += " AND DATE(te.timestamp) >= %s"
                params.append(start_date)
            if end_date:
                query += " AND DATE(te.timestamp) <= %s"
                params.append(end_date)
            
            query += " ORDER BY u.username, DATE(te.timestamp), te.timestamp;"
            cursor.execute(query, params)
            rows = cursor.fetchall()
            
            report_data = {}
            # Custom order of report columns
            custom_order = ['entrada', 'almuerzo_inicio', 'almuerzo_fin', 'salida']
            all_event_names = sorted(list(self.all_event_types.keys()))
            other_events = [name for name in all_event_names if name not in custom_order]
            sorted_event_names = custom_order[:-1] + other_events + [custom_order[-1]]

            for row in rows:
                username, date, event_name, time = row
                date_str = date.strftime("%Y-%m-%d")
                
                key = (username, date_str)
                if key not in report_data:
                    report_data[key] = {name: "Sin Marcación" for name in sorted_event_names}
                    report_data[key]["Usuario"] = username
                    report_data[key]["Fecha"] = date_str
                    
                report_data[key][event_name] = time.strftime("%I:%M:%S %p")

            # Prepare table for display
            headers = ["Usuario", "Fecha"] + [name.replace('_', ' ').title() for name in sorted_event_names]
            self.report_table.setColumnCount(len(headers))
            self.report_table.setHorizontalHeaderLabels(headers)
            self.report_table.setRowCount(len(report_data))
            
            sorted_keys = sorted(report_data.keys())
            for row_idx, key in enumerate(sorted_keys):
                record = report_data[key]
                self.report_table.setItem(row_idx, 0, QTableWidgetItem(record["Usuario"]))
                self.report_table.setItem(row_idx, 1, QTableWidgetItem(record["Fecha"]))
                for col_idx, event_name in enumerate(sorted_event_names):
                    display_time = record.get(event_name, "Sin Marcación")
                    self.report_table.setItem(row_idx, col_idx + 2, QTableWidgetItem(display_time))

            self.export_button.setEnabled(len(report_data) > 0)
        except (Exception, psycopg2.Error) as error:
            QMessageBox.critical(self, "Error de Base de Datos", f"Ocurrió un error al cargar reportes: {error}")
        finally:
            cursor.close()
            connection.close()
            
    def export_report(self):
        if self.report_table.rowCount() == 0:
            QMessageBox.warning(self, "Error", "No hay datos para exportar.")
            return

        file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Reporte", "reporte.xlsx", "Excel Files (*.xlsx)")
        if not file_path:
            return

        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Reporte de Horas"

        headers = [self.report_table.horizontalHeaderItem(i).text() for i in range(self.report_table.columnCount())]
        sheet.append(headers)

        for row_idx in range(self.report_table.rowCount()):
            row_data = [self.report_table.item(row_idx, col_idx).text() for col_idx in range(self.report_table.columnCount())]
            sheet.append(row_data)

        try:
            workbook.save(file_path)
            QMessageBox.information(self, "Éxito", f"Reporte exportado correctamente a:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Ocurrió un error al guardar el archivo:\n{e}")

    def create_new_user(self):
        new_username = self.new_user_entry.text()
        new_password = self.new_pass_entry.text()
        new_role = self.role_combo.currentText()
        
        if not new_username or not new_password:
            QMessageBox.warning(self, "Error", "Usuario y contraseña no pueden estar vacíos.")
            return
        
        connection = connect_to_db()
        if connection:
            cursor = connection.cursor()
            try:
                hashed_password = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt())
                sql = "INSERT INTO users (username, password_hash, role) VALUES (%s, %s, %s);"
                cursor.execute(sql, (new_username, hashed_password.decode('utf-8'), new_role))
                connection.commit()
                self.new_user_entry.clear()
                self.new_pass_entry.clear()
                self.load_users()
            except psycopg2.IntegrityError:
                QMessageBox.warning(self, "Error", "El nombre de usuario ya existe.")
            except (Exception, psycopg2.Error) as error:
                QMessageBox.critical(self, "Error de Base de Datos", f"Ocurrió un error: {error}")
            finally:
                cursor.close()
                connection.close()
        else:
            QMessageBox.critical(self, "Error de Conexión", "No se pudo conectar a la base de datos.")

    def change_user_password(self):
        username_to_change = self.pass_change_user_entry.text()
        new_password = self.new_pass_change_entry.text()
        
        if not username_to_change or not new_password:
            QMessageBox.warning(self, "Error", "Por favor, selecciona un usuario y escribe una nueva contraseña.")
            return

        connection = connect_to_db()
        if connection:
            cursor = connection.cursor()
            try:
                hashed_password = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt())
                sql = "UPDATE users SET password_hash = %s WHERE username = %s;"
                cursor.execute(sql, (hashed_password.decode('utf-8'), username_to_change))
                connection.commit()
                self.new_pass_change_entry.clear()
            except (Exception, psycopg2.Error) as error:
                QMessageBox.critical(self, "Error de Base de Datos", f"Ocurrió un error: {error}")
            finally:
                cursor.close()
                connection.close()
        else:
            QMessageBox.critical(self, "Error de Conexión", "No se pudo conectar a la base de datos.")
    
    def delete_user(self):
        username_to_delete = self.pass_change_user_entry.text()
        if not username_to_delete:
            QMessageBox.warning(self, "Error", "Por favor, selecciona un usuario para eliminar.")
            return
        
        if username_to_delete == self.username:
            QMessageBox.critical(self, "Error", "No puedes eliminar tu propio usuario de administrador.")
            return

        reply = QMessageBox.question(self, 'Confirmar Eliminación', 
                                    f"¿Estás seguro de que quieres eliminar al usuario '{username_to_delete}'?\nTodos sus registros de marcación también serán eliminados.", 
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            connection = connect_to_db()
            if connection:
                cursor = connection.cursor()
                try:
                    # Get user_id before deleting
                    cursor.execute("SELECT id FROM users WHERE username = %s;", (username_to_delete,))
                    user_id_to_delete = cursor.fetchone()[0]

                    # Delete time entries first
                    sql_entries = "DELETE FROM time_entries WHERE user_id = %s;"
                    cursor.execute(sql_entries, (user_id_to_delete,))

                    # Then delete the user
                    sql_user = "DELETE FROM users WHERE id = %s;"
                    cursor.execute(sql_user, (user_id_to_delete,))

                    connection.commit()
                    self.pass_change_user_entry.clear()
                    self.load_users()
                    self.load_reports()

                except (Exception, psycopg2.Error) as error:
                    QMessageBox.critical(self, "Error de Base de Datos", f"Ocurrió un error: {error}")
                finally:
                    cursor.close()
                    connection.close()
            else:
                QMessageBox.critical(self, "Error de Conexión", "No se pudo conectar a la base de datos.")

    def logout(self):
        self.hide()
        self.login_window = LoginWindow()
        self.login_window.show()

# --- Login Window ---

class LoginWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Programa de Marcación de Horas")
        self.setGeometry(100, 100, 400, 300)
        self.setup_ui()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_clock)
        self.timer.start(1000)

    def setup_ui(self):
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)
        
        self.clock_label = QLabel()
        self.clock_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.clock_label.setStyleSheet("font-size: 14pt; font-weight: bold;")
        self.layout.addWidget(self.clock_label)

        self.title_label = QLabel("Iniciar Sesión")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title_label.setStyleSheet("font-size: 16pt;")
        self.layout.addWidget(self.title_label)

        self.user_label = QLabel("Usuario:")
        self.layout.addWidget(self.user_label)
        self.user_entry = QLineEdit()
        self.layout.addWidget(self.user_entry)

        self.pass_label = QLabel("Contraseña:")
        self.layout.addWidget(self.pass_label)
        self.pass_entry = QLineEdit()
        self.pass_entry.setEchoMode(QLineEdit.EchoMode.Password)
        self.layout.addWidget(self.pass_entry)

        self.login_button = QPushButton("Entrar")
        self.layout.addWidget(self.login_button)
        self.login_button.clicked.connect(self.handle_login)

    def update_clock(self):
        panama_tz = pytz.timezone('America/Panama')
        panama_time = datetime.now(panama_tz)
        formatted_time = panama_time.strftime("%I:%M:%S %p")
        self.clock_label.setText(formatted_time)

    def handle_login(self):
        username = self.user_entry.text()
        password = self.pass_entry.text().encode('utf-8')
        
        connection = connect_to_db()
        if connection:
            cursor = connection.cursor()
            try:
                cursor.execute("SELECT id, password_hash, role FROM users WHERE username = %s;", (username,))
                result = cursor.fetchone()
                
                if result:
                    user_id, password_hash, user_role = result
                    password_hash = password_hash.encode('utf-8')
                    if bcrypt.checkpw(password, password_hash):
                        self.hide()
                        if user_role == 'admin':
                            self.admin_window = AdminWindow(user_id, username)
                            self.admin_window.show()
                        else:
                            self.time_clock_window = TimeClockWindow(user_id, username)
                            self.time_clock_window.show()
                    else:
                        QMessageBox.warning(self, "Error de Login", "Usuario o contraseña incorrectos.")
                else:
                    QMessageBox.warning(self, "Error de Login", "Usuario o contraseña incorrectos.")
            except (Exception, psycopg2.Error) as error:
                QMessageBox.critical(self, "Error de Base de Datos", f"Ocurrió un error: {error}")
            finally:
                cursor.close()
                connection.close()
        else:
            QMessageBox.critical(self, "Error de Conexión", "No se pudo conectar a la base de datos.")

def main():
    """
    This is the entry point for the Briefcase app.
    """
    app = QApplication(sys.argv)
    main_window = LoginWindow()
    main_window.show()
    return app.exec()

if __name__ == '__main__':
    main()