import sys
from calendar import monthrange
import pandas as pd
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QBoxLayout,
    QGridLayout,
    QButtonGroup,
    QRadioButton,
    QPushButton,
    QGroupBox,
    QListWidget,
    QComboBox,
    QLineEdit,
    QTableWidget,
    QTableWidgetItem,
    QAbstractItemView
)


FILE_NAME = "data.xlsx"
COMPANY_NAME = "Компания «Ресейл Опт»"
INPUT_DATA_TEXT = "Входные данные"
OUTPUT_DATA_TEXT = "Выходные данные"
SHOW_TEXT = "Показать"
WINDOW_COORDS = {
    "H": 100,
    "V": 20
}
INPUT_DATA_WIDTH = 300
OUTPUT_DATA_WIDTH = 800
SHOW_BUTTON_WIDTH = INPUT_DATA_WIDTH
PARAMETER_NAMES = [
    "Оценка",
    "Общая цена",
    "Общий вес"
]
OPERATION_NAMES = {
    "Максимум": "max",
    "Минимум": "min",
    "Среднее": "mean",
    "Медиана": "median",
    "Количество": "count",
    "Сумма": "sum"
}
TIME_PERIOD_TYPES = {
    "День": "D",
    "Месяц": "M",
    "Год": "Y",
    "Весь": "Весь"
}

objects_from_db = {
    "Процессы": {
        "Основное поле": "ID",
    },
    "Сотрудники": {
        "Основное поле": "Фамилия",
        "Элементы": []
    },
    "Компании": {
        "Основное поле": "Название",
        "Элементы": []
    },
    "Товары": {
        "Основное поле": "Название",
        "Элементы": []
    },
    "Склады": {
        "Основное поле": "Название",
        "Элементы": []
    }
}
data_from_form = {
    "Операции": {
        "Элементы": []
    },
    "Параметры": {
        "Элементы": []
    },
    "Объекты": {
        "Тип": "Сотрудники",
        "Элементы": []
    },
    "Временные промежутки": {
        "Тип": "",
        "Дата": {
            "Начало": "",
            "Конец": ""
        }
    }
}


def get_sheets(file_name):
    """
    Получает все листы excel-файла
    :param file_name: строка, имя excel-файла
    :return: словарь, где ключ - имя листа, значение - лист
    """
    data_file = pd.ExcelFile(file_name)
    return {sheet: data_file.parse(sheet) for sheet in data_file.sheet_names}


def get_column_values(sheets_, table_name, column_name):
    """
    Получает все значения столбца
    :param sheets_:  словарь DataFrame'ов, таблицы
    :param table_name: строка, имя таблицы, содержащей столбец
    :param column_name: строка, имя столбца, содержащего возвращаемые значения
    :return: список, значения столбца
    """
    return sheets_[table_name][column_name].tolist()


def add_processes_info(sheets_, table_names):
    """
    Добавляет в таблицу процессов столбцы из связанных таблиц по соответствующему ID-столбцу
    :param sheets_: список DataFrame'ов, все таблицы
    :param table_names: список, названия связанных таблиц, из которых в процессы добавляются их основные столбцы, например: ID (сотрудник) и Фамилия (Сотрудник)
    :return: таблица, процессы с добавленными столбцами
    """
    processes_ = sheets_["Процессы"]
    for table_name in table_names:
        processes_ = pd.merge(processes_, sheets_[table_name], left_on="ID (" + table_name + ")", right_on="ID")
        processes_ = processes_.rename(columns={"ID_x": "ID"})
        processes_.drop("ID (" + table_name + ")", axis=1, inplace=True)
        processes_ = processes_.rename(columns={"ID_y": "ID (" + table_name + ")"})
        processes_ = processes_.rename(
            columns={
                objects_from_db[table_name]["Основное поле"]: objects_from_db[table_name][
                                                                  "Основное поле"] + " (" + table_name + ")"
            }
        )
    return processes_


def filter_processes(sheets_, date_range=None, conditions=None):
    """
    Получает из списка процессов
        - находящиеся во временном диапазоне
        - соответствующие значениям объектов из других таблиц, например, только для двух указанных сотрудников и одной указанной компании
    :param sheets_: список DataFrame'ов, все таблицы
    :param date_range: словарь, диапазон дат: {"start": "2020-01-01", "end": "2025-12-31"}
    :param conditions: словарь списков, названия таблиц и соответствующие значения их основных столбцов для поиска
    :return: DataFrame, отфильтрованная таблица процессов
    """
    processes_ = sheets_["Процессы"]
    if conditions is not None:
        processes_ = add_processes_info(sheets_, conditions.keys())
        for condition_name, condition_values in conditions.items():
            table_name = objects_from_db[condition_name]["Основное поле"]
            column_name = condition_name
            column_values = condition_values
            processes_ = processes_[processes_[table_name + " (" + column_name + ")"].isin(column_values)]
    if date_range is not None:
        processes_ = processes_[
            (processes_["Дата"] >= pd.to_datetime(date_range["start"])) &
            (processes_["Дата"] <= pd.to_datetime(date_range["end"]))
            ]
    return processes_


def calculate_grouped_processes(sheets_, func_names, date_group, date_range=None, conditions=None):
    """
    Вычисляет значения для процессов
    :param sheets_: список DataFrame'ов, все таблицы
    :param func_names: список строк, название вычислительных операций, проводимых над значениями
    :param date_group: строка, тип группировки по датам: день, месяц, год, весь; вычисляется только для одного значения одного condition
    :param date_range: словарь, диапазон дат, например: {"start": "2020-01-01", "end": "2025-12-31"}
    :param conditions: словарь списков, названия таблиц и соответствующие значения их основных столбцов для поиска
    :return: DataFrame, столбцы - дата и по столбцу на каждую функцию; не содержит единственное значение единственного condition, так как оно не табличное, а единственное, и есть во входном параметре
    """
    date_group = TIME_PERIOD_TYPES[date_group]
    filtered_processes = filter_processes(sheets_, date_range, conditions)
    filtered_processes.merge(
        sheets_[list(conditions.keys())[0]],
        left_on="ID (" + list(conditions.keys())[0] + ")",
        right_on="ID"
    )
    grouped_processes1 = pd.DataFrame()
    grouped_processes2 = pd.DataFrame()
    if "Сумма" in func_names:
        filtered_processes = filtered_processes.merge(
            sheets_["Товары"],
            left_on="ID (Товары)",
            right_on="ID"
        )
        filtered_processes = filtered_processes.rename(columns={"ID_x": "ID"})
        filtered_processes = filtered_processes.drop("ID_y", axis=1)
        filtered_processes["Общая цена"] = filtered_processes["Цена"] * filtered_processes["Количество"]
        filtered_processes["Общий вес"] = filtered_processes["Вес"] * filtered_processes["Количество"]
        group = [objects_from_db[list(conditions.keys())[0]]["Основное поле"]
                 + " (" + list(conditions.keys())[0]
                 + ")"]
        if date_group != "Весь":
            group.append(filtered_processes["Дата"].dt.to_period(date_group))
        grouped_processes1 = filtered_processes.groupby(group).agg(
            price=("Общая цена", "sum"),
            weight=("Общий вес", "sum")).round(2)
        grouped_processes1.rename(
            columns={"price": "Общая цена"},
            inplace=True
        )
        grouped_processes1.rename(
            columns={"weight": "Общий вес"},
            inplace=True
        )
        grouped_processes1 = grouped_processes1.reset_index()
    agg_funcs = []
    for key, value in OPERATION_NAMES.items():
        if key in func_names and key != "Сумма":
            agg_funcs.append(value)
    if agg_funcs:
        group = [objects_from_db[list(conditions.keys())[0]]["Основное поле"] +
                 " (" +
                 list(conditions.keys())[0] +
                 ")"]
        if date_group != "Весь":
            group.append(filtered_processes["Дата"].dt.to_period(date_group))
        grouped_processes2 = filtered_processes.groupby(group)["Оценка"].agg(agg_funcs).round(2)
        rename_columns = {}
        for agg_func in agg_funcs:
            rename_columns[agg_func] = [key for key, value in OPERATION_NAMES.items() if value == agg_func][0]
        grouped_processes2 = grouped_processes2.rename(columns=rename_columns)
        grouped_processes2 = grouped_processes2.reset_index()
    if not grouped_processes1.empty and grouped_processes2.empty:
        grouped_processes = grouped_processes1
    elif grouped_processes1.empty and not grouped_processes2.empty:
        grouped_processes = grouped_processes2
    else:
        group_columns = [
            objects_from_db[list(conditions.keys())[0]]["Основное поле"] +
            " (" +
            list(conditions.keys())[0] +
            ")"
        ]
        if date_group != "Весь":
            group_columns.append("Дата")
        grouped_processes = pd.merge(
            grouped_processes1,
            grouped_processes2,
            on=group_columns
        )
    grouped_processes.rename(
        columns={
            objects_from_db[list(conditions.keys())[0]]["Основное поле"] + " (" + list(conditions.keys())[0] + ")":
                list(conditions.keys())[0]
        },
        inplace=True
    )
    return grouped_processes


def fill_objects_from_db(sheets_):
    """
    Заполняет переменную, хранящую списки значений объектов
    :param sheets_: словарь DataFrame'ов, таблицы
    :return: None
    """
    for key, value in objects_from_db.items():
        if key != "Процессы":
            objects_from_db[key]["Элементы"] = get_column_values(
                sheets_,
                key,
                objects_from_db[key]["Основное поле"]
            )


def create_gb_with_lo(name, direction, *elements):
    """
    Создаёт GroupBox с обводкой, выровненных внутри при помощи Layout, и заполняет его элементами
    :param name: строка, название GroupBox (отображаемый текст)
    :param direction: QBoxLayout.Direction, направление выравнивания в Layout (по горизонтали/вертикали)
    :param elements: список Widget/Layout, располагаемых внутри GroupBox
    :return: GroupBox
    """
    gb = QGroupBox(name)
    lo = QBoxLayout(direction)
    for element in elements:
        if type(element) is QBoxLayout or type(element) is QGridLayout:
            lo.addLayout(element)
        else:
            lo.addWidget(element)
    gb.setLayout(lo)
    return gb


def create_lo(direction, *elements):
    """
    Создаёт Layout без обводки и заполняет его элементами
    :param direction: QBoxLayout.Direction, направление выравнивания в Layout (по горизонтали/вертикали)
    :param elements: список Widget/Layout, располагаемых внутри GroupBox
    :return: Layout
    """
    lo = QBoxLayout(direction)
    for element in elements:
        if type(element) is QBoxLayout or type(element) is QGridLayout:
            lo.addLayout(element)
        else:
            lo.addWidget(element)
    return lo


def create_lb(items):
    """
    Создаёт список с множественным выбором ListWidget и заполняет его элементами
    :param items: список строк, текст элементов ListWidget
    :return: ListBox
    """
    lb = QListWidget()
    lb.addItems(items)
    lb.setSelectionMode(QAbstractItemView.ExtendedSelection)
    return lb


def create_cb(items):
    """
    Создаёт выпадающий список ComboBox
    :param items: список строк, текст элементов ComboBox
    :return: ComboBox
    """
    cb = QComboBox()
    cb.addItems(items)
    return cb


def create_lo_with_rbg(direction, *rb_names):
    """
    Создаёт Layout с множеством RadioButton'ов
    :param direction: QBoxLayout.Direction, направление выравнивания в Layout (по горизонтали/вертикали)
    :param rb_names: список строк, текст RadioButton'ов
    :return: Layout
    """
    rbg = QButtonGroup()
    layout = QBoxLayout(direction)
    for number, rb_name in enumerate(rb_names):
        button = QRadioButton(rb_name)
        if number == 0:
            button.setChecked(True)
        rbg.addButton(button, number)
        layout.addWidget(button)
    return layout


sheets = get_sheets(FILE_NAME)
pd.set_option("display.max_columns", None)
fill_objects_from_db(sheets)


def send_data_from_form_to_db_response():
    """
    Преобразует формат данных из собранных с формы в отправляемые запросом к БД, и отправляет запрос
    :return: данные в преобразованном формате
    """
    operations = data_from_form["Операции"]["Элементы"]
    parameters = data_from_form["Параметры"]["Элементы"]
    if ("Общая цена" in parameters or "Общий вес" in parameters) and "Сумма" not in operations:
        operations.append("Сумма")
    if "Сумма" in operations:
        if "Общая цена" not in parameters:
            parameters.append("Общая цена")
        if "Общий вес" not in parameters:
            parameters.append("Общий вес")
    objects_type = data_from_form["Объекты"]["Тип"]
    objects_name = data_from_form["Объекты"]["Элементы"]
    time_periods_type = data_from_form["Временные промежутки"]["Тип"]
    time_periods_name = data_from_form["Временные промежутки"]["Дата"]
    if time_periods_type == "Год":
        year_start = time_periods_name["Начало"].split(".")[2]
        year_end = time_periods_name["Конец"].split(".")[2]
        time_periods = {"start": f"{year_start}-01-01", "end": f"{year_end}-12-31"}
    elif time_periods_type == "Месяц":
        date_start_parts = time_periods_name["Начало"].split(".")
        date_end_parts = time_periods_name["Конец"].split(".")
        month_start = int(date_start_parts[1])
        year_start = int(date_start_parts[2])
        year_end = int(date_end_parts[2])
        month_end = int(date_end_parts[1])
        days_end = monthrange(year_end, month_end)[1]
        time_periods = {
            "start": f"{year_start}-{month_start}-01",
            "end": f"{year_end}-{month_end}-{days_end}"
        }
    else:
        date_start_parts = time_periods_name["Начало"].split(".")
        date_end_parts = time_periods_name["Конец"].split(".")
        time_periods = {
            "start": f"{date_start_parts[2]}-{date_start_parts[1]}-{date_start_parts[0]}",
            "end": f"{date_end_parts[2]}-{date_end_parts[1]}-{date_end_parts[0]}"
        }
    return calculate_grouped_processes(
        sheets,
        func_names=operations,
        date_group=time_periods_type,
        date_range=time_periods,
        conditions={objects_type: objects_name}
    )


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.table_output = None
        self.gb_time_periods_date_end = None
        self.le_time_periods_end = None
        self.gb_time_periods_date_start = None
        self.le_time_periods_start = None
        self.lo_main = None
        self.pb_show = None
        self.lo_data = None
        self.gb_output_data = None
        self.gb_input_data = None
        self.gb_time_periods = None
        self.gb_time_periods_name = None
        self.gb_time_periods_type = None
        self.cb_time_periods_type = None
        self.gb_time_periods_count = None
        self.gb_objects = None
        self.gb_objects_name = None
        self.lb_objects_name = None
        self.cb_objects_name = None
        self.gb_objects_type = None
        self.cb_objects_type = None
        self.gb_objects_count = None
        self.gb_parameters = None
        self.lb_parameters = None
        self.cb_parameters = None
        self.gb_parameters_count = None
        self.gb_operations = None
        self.lb_operations = None
        self.cb_operations = None
        self.gb_operations_count = None
        self.create_window()

    def get_data_from_form(self):
        """
        Собирает все данные с формы в переменную
        :return: словарь с данными с формы
        """
        if self.gb_operations_count.itemAt(0).widget().isChecked():
            data_from_form["Операции"]["Элементы"] = [self.cb_operations.currentText()]
        else:
            selected_indexes = self.lb_operations.selectedIndexes()
            selected_texts = []
            for index in selected_indexes:
                text = index.data(Qt.ItemDataRole.DisplayRole)
                if text is not None:
                    selected_texts.append(text)
            data_from_form["Операции"]["Элементы"] = selected_texts
        if self.gb_parameters_count.itemAt(0).widget().isChecked():
            data_from_form["Параметры"]["Элементы"] = [self.cb_parameters.currentText()]
        else:
            selected_indexes = self.lb_parameters.selectedIndexes()
            selected_texts = []
            for index in selected_indexes:
                text = index.data(Qt.ItemDataRole.DisplayRole)
                if text is not None:
                    selected_texts.append(text)
            data_from_form["Параметры"]["Элементы"] = selected_texts

        data_from_form["Объекты"]["Тип"] = self.cb_objects_type.currentText()
        if self.gb_objects_count.itemAt(0).widget().isChecked():
            data_from_form["Объекты"]["Элементы"] = [self.cb_objects_name.currentText()]
        else:
            selected_indexes = self.lb_objects_name.selectedIndexes()
            selected_texts = []
            for index in selected_indexes:
                text = index.data(Qt.ItemDataRole.DisplayRole)
                if text is not None:
                    selected_texts.append(text)
            data_from_form["Объекты"]["Элементы"] = selected_texts
        data_from_form["Временные промежутки"]["Тип"] = self.cb_time_periods_type.currentText()
        data_from_form["Временные промежутки"]["Дата"]["Начало"] = self.le_time_periods_start.text()
        data_from_form["Временные промежутки"]["Дата"]["Конец"] = self.le_time_periods_end.text()
        return data_from_form

    def show_output_data(self, output_data):
        self.table_output.setRowCount(output_data.shape[0])
        self.table_output.setColumnCount(output_data.shape[1])
        self.table_output.setHorizontalHeaderLabels(output_data.columns)
        for row in range(output_data.shape[0]):
            for col in range(output_data.shape[1]):
                item = QTableWidgetItem(str(output_data.iloc[row, col]))
                self.table_output.setItem(row, col, item)
        self.table_output.resizeColumnsToContents()

    def pb_clicked(self):
        self.get_data_from_form()
        df_data = send_data_from_form_to_db_response()
        self.show_output_data(df_data)

    def cb_time_periods_type_changed(self):
        """
        Событие переключения RadioButton'ов "Один"/"Много" в блоке временных промежутков
        :return: None
        """
        if self.gb_time_periods_count.itemAt(0).widget().isChecked():
            self.gb_time_periods_type.hide()
        else:
            self.gb_time_periods_type.show()

    def cb_objects_type_changed(self):
        """
        Событие изменения выбранного элемента ComboBox в блоке объектов
        :return: None
        """
        element = objects_from_db[self.cb_objects_type.currentText()]
        self.gb_objects_name.setTitle(element["Основное поле"])
        self.cb_objects_name.clear()
        self.cb_objects_name.addItems(element["Элементы"])
        self.lb_objects_name.clear()
        self.lb_objects_name.addItems(element["Элементы"])

    def rbg_changed(self, widget_show, widget_hide):
        """
        Событие переключения RadioButton'ов "Один"/"Много"
        :param widget_show: RadioButton, требуемый к включению переключатель
        :param widget_hide: RadioButton, требуемый к выключению переключатель
        :return: None
        """
        widget_show.show()
        widget_hide.hide()

    def create_gb_count_block(self, widget_one, widget_many):
        """
        Создаёт подблок с переключателями "Один"/"Много"
        :param widget_one: RadioButton, переключатель "Один"
        :param widget_many: RadioButton, переключатель "Много"
        :return: ButtonGroup, сгруппированные RadioButton'ы
        """
        rgb = create_lo_with_rbg(QBoxLayout.LeftToRight, "Один", "Много")
        rgb.itemAt(0).widget().clicked.connect(lambda: self.rbg_changed(widget_one, widget_many))
        rgb.itemAt(1).widget().clicked.connect(lambda: self.rbg_changed(widget_many, widget_one))
        return rgb

    def create_operations_block(self):
        """
        Создаёт блок операций
        :return: None
        """
        operation_names = OPERATION_NAMES
        self.cb_operations = create_cb(operation_names)
        self.lb_operations = create_lb(operation_names)
        self.lb_operations.setCurrentItem(self.lb_operations.item(0))
        self.lb_operations.hide()
        self.gb_operations_count = self.create_gb_count_block(self.cb_operations, self.lb_operations)
        self.gb_operations = create_gb_with_lo(
            "Операции",
            QBoxLayout.TopToBottom,
            self.gb_operations_count,
            self.cb_operations,
            self.lb_operations
        )

    def create_parameters_block(self):
        """
        Создаёт блок параметров
        :return: None
        """
        parameter_names = PARAMETER_NAMES
        self.cb_parameters = create_cb(parameter_names)
        self.lb_parameters = create_lb(parameter_names)
        self.lb_parameters.setCurrentItem(self.lb_parameters.item(0))
        self.lb_parameters.hide()
        self.gb_parameters_count = self.create_gb_count_block(self.cb_parameters, self.lb_parameters)
        self.gb_parameters = create_gb_with_lo(
            "Параметры",
            QBoxLayout.TopToBottom,
            self.gb_parameters_count,
            self.cb_parameters,
            self.lb_parameters
        )

    def create_objects_block(self):
        """
        Создаёт блок объектов
        :return: None
        """
        self.cb_objects_type = create_cb([every_element for every_element in list(objects_from_db.keys())[1:]])
        self.cb_objects_type.currentIndexChanged.connect(self.cb_objects_type_changed)
        self.gb_objects_type = create_gb_with_lo("Тип", QBoxLayout.TopToBottom, self.cb_objects_type)
        self.cb_objects_name = create_cb(objects_from_db[list(objects_from_db)[1]]["Элементы"])
        self.lb_objects_name = create_lb(objects_from_db[self.cb_objects_type.currentText()]["Элементы"])
        self.lb_objects_name.setCurrentItem(self.lb_objects_name.item(0))
        self.lb_objects_name.hide()
        self.gb_objects_count = self.create_gb_count_block(self.cb_objects_name, self.lb_objects_name)
        self.gb_objects_name = create_gb_with_lo(
            objects_from_db[self.cb_objects_type.currentText()]["Основное поле"],
            QBoxLayout.TopToBottom,
            self.cb_objects_name,
            self.lb_objects_name)
        self.gb_objects = create_gb_with_lo(
            "Объекты",
            QBoxLayout.TopToBottom,
            self.gb_objects_count,
            self.gb_objects_type,
            self.gb_objects_name
        )

    def create_time_periods_block(self):
        """
        Создаёт блок временных промежутков
        :return: None
        """
        self.cb_time_periods_type = create_cb(TIME_PERIOD_TYPES.keys())
        self.gb_time_periods_type = create_gb_with_lo(
            "Периодичность",
            QBoxLayout.TopToBottom,
            self.cb_time_periods_type
        )
        self.le_time_periods_start = QLineEdit()
        self.gb_time_periods_date_start = create_gb_with_lo(
            "Начало",
            QBoxLayout.TopToBottom,
            self.le_time_periods_start
        )
        self.le_time_periods_end = QLineEdit()
        self.gb_time_periods_date_end = create_gb_with_lo(
            "Конец",
            QBoxLayout.TopToBottom,
            self.le_time_periods_end
        )
        self.gb_time_periods_name = create_gb_with_lo(
            "Дата",
            QBoxLayout.TopToBottom,
            self.gb_time_periods_date_start,
            self.gb_time_periods_date_end
        )
        self.gb_time_periods = create_gb_with_lo(
            "Временные промежутки",
            QBoxLayout.TopToBottom,
            self.gb_time_periods_count,
            self.gb_time_periods_type,
            self.gb_time_periods_name
        )

    def create_input_data_block(self):
        """
        Создаёт блок входных данных
        :return: None
        """
        self.gb_input_data = create_gb_with_lo(
            INPUT_DATA_TEXT,
            QBoxLayout.TopToBottom,
            self.gb_operations,
            self.gb_parameters,
            self.gb_objects,
            self.gb_time_periods
        )
        self.gb_input_data.setFixedWidth(INPUT_DATA_WIDTH)

    def create_output_data_block(self):
        """
        Создаёт блок выходных данных
        :return: None
        """
        self.table_output = QTableWidget()
        self.gb_output_data = create_gb_with_lo(
            OUTPUT_DATA_TEXT,
            QBoxLayout.TopToBottom,
            self.table_output
        )
        self.gb_output_data.setFixedWidth(OUTPUT_DATA_WIDTH)

    def create_data_block(self):
        """
        Создаёт блок данных
        :return: None
        """
        self.lo_data = create_lo(
            QBoxLayout.LeftToRight,
            self.gb_input_data,
            self.gb_output_data
        )

    def create_main_block(self):
        """
        Создаёт основной блок
        :return: None
        """
        self.pb_show = QPushButton(SHOW_TEXT)
        self.pb_show.setFixedWidth(SHOW_BUTTON_WIDTH)
        self.pb_show.clicked.connect(self.pb_clicked)
        self.lo_main = create_lo(
            QBoxLayout.TopToBottom,
            self.lo_data,
            self.pb_show
        )
        self.setLayout(self.lo_main)

    def create_window(self):
        """
        Создаёт окно
        :return: None
        """
        self.setWindowTitle(COMPANY_NAME)
        self.move(WINDOW_COORDS["H"], WINDOW_COORDS["V"])
        self.create_operations_block()
        self.create_parameters_block()
        self.create_objects_block()
        self.create_time_periods_block()
        self.create_input_data_block()
        self.create_output_data_block()
        self.create_data_block()
        self.create_main_block()


app = QApplication(sys.argv)
main_window = MainWindow()
main_window.show()
sys.exit(app.exec_())
