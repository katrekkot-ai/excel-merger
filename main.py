"""Excel Merger — перенос остатков из выгрузки 1С в материальный отчёт (форма М-19)."""

import os
import re
import shutil
from collections import defaultdict
from copy import copy

import openpyxl
import pandas as pd
from kivy.app import App
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.metrics import dp, sp
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.widget import Widget
from kivy.utils import platform

# Для работы с системным выбором файла
if platform == 'android':
    from android import activity as android_activity
    from android.permissions import request_permissions, Permission
    from android.storage import app_storage_path, primary_external_storage_path
    from jnius import autoclass

REQUEST_OPEN_DOCUMENT = 1001
RESULT_OK = -1

# ---------- правила разбора таблиц ----------

# Название материала: в выгрузках 1С колонка «Материал» — это код,
# а человекочитаемое название лежит в «Краткий текст материала».
MATERIAL_NAME_KEYWORDS = ('краткий текст материала', 'наименование',
                          'название', 'номенклатура')

# Что переносим из эталона в отчёт. Ключи проверяются по порядку:
# сначала точные имена из выгрузки 1С, потом «человеческие» из формы.
METRIC_KEYWORDS = {
    'остаток': ('остзапнп', 'остаток на начало', 'остаток'),
    'приход': ('постматер', 'приход'),
    'расход': ('отпуск мат', 'расход'),
    'конец': ('остзапкп', 'остаток на конец', 'конец'),
}

# Слова, по которым узнаём строку с подзаголовками («наименование», «код»...).
HEADER_WORDS = ('наименование', 'материал', 'код', 'кол-во', 'кол- во',
                'количество', 'сумма', 'остаток', 'остзап', 'приход',
                'постматер', 'расход', 'отпуск', 'цена', 'единица',
                'номер по порядку')

# Слова, по которым узнаём строку-группу над подзаголовками
# («Остаток на начало отчетного периода» и т.п.).
GROUP_WORDS = ('товарно', 'остаток', 'приход', 'расход', 'единица',
               'цена', 'номер по порядку')

# Колонка-единица измерения: «Ед.», «Ед. Unnamed: 2_level_1» и т.п.
UNIT_TOKENS = {'ед', 'шт', 'кг', 'т', 'л', 'мл', 'м', 'уп', 'пач', 'рул'}

# Ссылка на ячейку внутри формулы: H170, $H$170, H$170.
FORMULA_REF = re.compile(r'(\$?[A-Z]{1,3})(\$?)(\d+)')


def _is_unit_column(name):
    """True, если в заголовке нет ничего, кроме единицы измерения."""
    tokens = [t for t in re.split(r'[^0-9a-zа-я]+', str(name).lower()) if t]
    tokens = [t for t in tokens
              if not t.isdigit() and t not in ('unnamed', 'level', 'уровень')]
    return bool(tokens) and all(t in UNIT_TOKENS for t in tokens)


def _norm_code(value):
    """Код материала -> строка из 18 цифр (в 1С ведущие нули теряются)."""
    if value is None:
        return ''
    if isinstance(value, float):
        if pd.isna(value):
            return ''
        if value.is_integer():
            value = int(value)
    text = str(value).strip()
    text = re.sub(r'\.0+$', '', text)
    if text.isdigit():
        return text.zfill(18)
    return text


def _norm_name(value):
    """Название материала для сравнения: без лишних пробелов, в нижнем регистре."""
    return re.sub(r'\s+', ' ', str(value or '')).strip().lower()


# ВАЖНО: размеры только в dp/sp. В пикселях (height=44) на телефоне с плотностью 3
# кнопка получается ~15 dp и текст в неё не влезает.
BTN_HEIGHT = dp(52)
BTN_FONT = sp(17)
STATUS_FONT = sp(15)
STATUS_MIN_HEIGHT = dp(72)
STATUS_MAX_HEIGHT = dp(160)


class ExcelMergerApp(App):
    _pending_type = None

    def build(self):
        self.etalon_path = None
        self.edit_path = None
        self._diag = []

        # Запрашиваем разрешения при запуске
        if platform == 'android':
            request_permissions([Permission.READ_EXTERNAL_STORAGE])
            # Результат системного диалога выбора файла приходит сюда
            android_activity.bind(on_activity_result=self._on_activity_result)

        layout = BoxLayout(orientation='vertical', padding=dp(16), spacing=dp(12))

        Window.clearcolor = (0.96, 0.96, 0.96, 1)

        self.status = Label(text='Загрузите эталон и файл для редактирования',
                            size_hint_y=None, height=STATUS_MIN_HEIGHT,
                            halign='center', valign='middle', font_size=STATUS_FONT,
                            color=(0.12, 0.12, 0.12, 1), padding=(dp(8), dp(8)))
        # halign работает только вместе с text_size; высота подстраивается под текст
        self.status.bind(size=self._sync_status_text_size)
        self.status.bind(texture_size=self._sync_status_height)
        layout.add_widget(self.status)

        btn_etalon = Button(text='Выбрать эталон', size_hint_y=None, height=BTN_HEIGHT,
                            font_size=BTN_FONT)
        btn_etalon.bind(on_press=self._select_file('etalon'))
        layout.add_widget(btn_etalon)

        btn_edit = Button(text='Выбрать файл для редактирования', size_hint_y=None,
                          height=BTN_HEIGHT, font_size=BTN_FONT)
        btn_edit.bind(on_press=self._select_file('edit'))
        layout.add_widget(btn_edit)

        btn_run = Button(text='Объединить и сохранить', size_hint_y=None,
                         height=BTN_HEIGHT, font_size=BTN_FONT)
        btn_run.bind(on_press=self._process)
        layout.add_widget(btn_run)

        # BoxLayout(vertical) укладывает детей от нижнего края вверх — без распорки
        # фиксированные по высоте кнопки «прилипают» к низу экрана.
        layout.add_widget(Widget())

        return layout

    def _sync_status_text_size(self, instance, size):
        instance.text_size = (max(0, size[0] - dp(8)), None)

    def _sync_status_height(self, instance, texture_size):
        height = texture_size[1] + dp(24)
        instance.height = min(STATUS_MAX_HEIGHT, max(STATUS_MIN_HEIGHT, height))

    def _clean_number(self, value):
        """Очистка чисел: «1 000,50 ₽» -> 1000.5."""
        if value is None:
            return 0
        if isinstance(value, float) and pd.isna(value):
            return 0
        if isinstance(value, (int, float)):
            return value
        s = re.sub(r'[^\d.,-]', '', str(value)).replace(',', '.')
        try:
            return float(s)
        except ValueError:
            return 0

    @staticmethod
    def _to_number(value):
        """Значение ячейки -> число (или None, если это не число)."""
        if value is None or isinstance(value, bool):
            return None
        if isinstance(value, (int, float)):
            return float(value)
        text = str(value).replace(' ', '').replace('\xa0', '').replace(',', '.')
        try:
            return float(text)
        except ValueError:
            return None

    def _find_first_column(self, df, keywords):
        """Первая колонка, в заголовке которой есть один из ключей (по порядку)."""
        columns = [(str(col).strip().lower(), col) for col in df.columns]
        for key in keywords:
            for low, col in columns:
                if _is_unit_column(low):
                    continue
                if key in low:
                    return col
        return None

    def _find_metric_column(self, df, keywords):
        return self._find_first_column(df, keywords)

    def _find_material_column(self, df):
        col = self._find_first_column(df, MATERIAL_NAME_KEYWORDS)
        if col is not None:
            return col
        return self._find_first_column(df, ('материал',))

    def _output_dir(self):
        """Первая папка, в которую реально можно писать."""
        candidates = []
        if platform == 'android':
            candidates.append(os.path.join(primary_external_storage_path(), 'Download'))
            candidates.append(app_storage_path())
        else:
            candidates.append(os.path.join(os.path.expanduser('~'), 'Download'))
        candidates.append(os.path.expanduser('~'))
        candidates.append(os.getcwd())

        for path in candidates:
            try:
                os.makedirs(path, exist_ok=True)
                test = os.path.join(path, '.write_test')
                with open(test, 'w') as f:
                    f.write('ok')
                os.remove(test)
                return path
            except Exception:
                continue
        return os.path.expanduser('~')

    # MIME-тип книги Excel — нужен и «Загрузкам», и просмотрщику на телефоне.
    XLSX_MIME = ('application/vnd.openxmlformats-officedocument'
                 '.spreadsheetml.sheet')

    def _publish_to_download(self, source_path):
        """Кладёт результат в общую папку «Загрузки», чтобы его было где найти.

        На Android 10+ прямая запись в /storage/emulated/0/Download запрещена,
        поэтому файл добавляется через MediaStore. Возвращает описание места
        сохранения или None, если не получилось (тогда файл остаётся в папке
        приложения).
        """
        if platform != 'android':
            return None

        # уже сохранили прямо в «Загрузки» — второй копии не нужно
        try:
            download_dir = os.path.join(primary_external_storage_path(), 'Download')
            if os.path.abspath(os.path.dirname(source_path)) == os.path.abspath(download_dir):
                return source_path
        except Exception:
            pass

        name = os.path.basename(source_path)
        mime = self.XLSX_MIME if name.lower().endswith(('.xlsx', '.xlsm')) else 'text/plain'
        try:
            sdk = autoclass('android.os.Build$VERSION').SDK_INT
        except Exception:
            sdk = 0

        if sdk >= 29:
            try:
                resolver = (autoclass('org.kivy.android.PythonActivity')
                            .mActivity.getContentResolver())
                MediaStore = autoclass('android.provider.MediaStore')
                ContentValues = autoclass('android.content.ContentValues')
                values = ContentValues()
                values.put(MediaStore.MediaColumns.DISPLAY_NAME, name)
                values.put(MediaStore.MediaColumns.MIME_TYPE, mime)
                values.put(MediaStore.MediaColumns.RELATIVE_PATH, 'Download')
                uri = resolver.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, values)
                if uri is None:
                    return None
                stream = resolver.openOutputStream(uri)
                with open(source_path, 'rb') as source:
                    while True:
                        chunk = source.read(65536)
                        if not chunk:
                            break
                        stream.write(bytearray(chunk))
                stream.flush()
                stream.close()
                return 'Загрузки/' + name
            except Exception:
                return None

        # Android 9 и ниже: разрешение WRITE_EXTERNAL_STORAGE позволяет писать напрямую
        try:
            folder = os.path.join(primary_external_storage_path(), 'Download')
            os.makedirs(folder, exist_ok=True)
            shutil.copyfile(source_path, os.path.join(folder, name))
            return os.path.join(folder, name)
        except Exception:
            return None

    def _select_file(self, file_type):
        """Функция-обертка для вызова диалога."""
        def inner(instance):
            if platform != 'android':
                # На ПК просто открываем старый FileChooser
                chooser_popup = Popup(title='Выберите Excel файл', size_hint=(0.95, 0.95))
                chooser = FileChooserListView(filters=['*.xlsx', '*.xls'])
                btn_select = Button(text='Выбрать', size_hint_y=None, height=BTN_HEIGHT,
                                    font_size=BTN_FONT)

                def on_select(_):
                    if not chooser.selection:
                        return
                    chooser_popup.dismiss()
                    self._set_path(file_type, os.path.join(chooser.path, chooser.selection[0]))

                btn_select.bind(on_press=on_select)
                content = BoxLayout(orientation='vertical')
                content.add_widget(chooser)
                content.add_widget(btn_select)
                chooser_popup.content = content
                chooser_popup.open()
            else:
                # На Андроид используем системный выбор
                activity = autoclass('org.kivy.android.PythonActivity').mActivity
                Intent = autoclass('android.content.Intent')
                intent = Intent(Intent.ACTION_OPEN_DOCUMENT)
                intent.addCategory(Intent.CATEGORY_OPENABLE)
                intent.setType('*/*')
                self._pending_type = file_type
                activity.startActivityForResult(intent, REQUEST_OPEN_DOCUMENT)
        return inner

    def _on_activity_result(self, request_code, result_code, intent):
        """Колбэк системного диалога Android (приходит из Java-потока)."""
        if request_code != REQUEST_OPEN_DOCUMENT or not self._pending_type:
            return
        file_type = self._pending_type
        self._pending_type = None

        if result_code != RESULT_OK or intent is None or intent.getData() is None:
            Clock.schedule_once(lambda dt: self._update_status('Файл не выбран'))
            return

        try:
            # content:// нельзя открыть напрямую — копируем файл к себе
            path = self._copy_uri_to_file(intent.getData())
        except Exception as e:
            message = f'Не удалось прочитать файл: {e}'
            Clock.schedule_once(lambda dt: self._update_status(message))
            return

        Clock.schedule_once(lambda dt: self._set_path(file_type, path))

    def _copy_uri_to_file(self, uri):
        resolver = autoclass('org.kivy.android.PythonActivity').mActivity.getContentResolver()

        name = None
        cursor = resolver.query(uri, None, None, None, None)
        if cursor is not None:
            try:
                index = cursor.getColumnIndex(autoclass('android.provider.OpenableColumns').DISPLAY_NAME)
                if index >= 0 and cursor.moveToFirst():
                    name = cursor.getString(index)
            finally:
                cursor.close()
        if not name:
            name = 'input.xlsx'
        if not name.lower().endswith(('.xlsx', '.xls')):
            name += '.xlsx'

        dest_dir = os.path.join(app_storage_path(), 'input')
        os.makedirs(dest_dir, exist_ok=True)
        dest = os.path.join(dest_dir, name)

        BufferedInputStream = autoclass('java.io.BufferedInputStream')
        FileOutputStream = autoclass('java.io.FileOutputStream')
        source = BufferedInputStream(resolver.openInputStream(uri))
        target = FileOutputStream(dest)
        buffer = bytearray(65536)
        try:
            while True:
                read = source.read(buffer)
                if read == -1:
                    break
                target.write(buffer, 0, read)
        finally:
            source.close()
            target.close()
        return dest

    def _set_path(self, file_type, path):
        if file_type == 'etalon':
            self.etalon_path = path
        else:
            self.edit_path = path
        self._update_status()

    def _update_status(self, message=None):
        if message is not None:
            self.status.text = message
            return
        etalon = os.path.basename(self.etalon_path) if self.etalon_path else 'не выбран'
        edit = os.path.basename(self.edit_path) if self.edit_path else 'не выбран'
        self.status.text = f'Эталон: {etalon}\nФайл: {edit}'

    # ---------- чтение эталона (pandas) ----------

    @staticmethod
    def _flatten_columns(df):
        """Многоуровневые заголовки -> обычные строки с уникальными именами."""
        names = []
        counts = {}
        for col in df.columns:
            name = ' '.join(map(str, col)) if isinstance(col, tuple) else str(col)
            # pandas дописывает к пустым ячейкам шапки 'Unnamed: 3_level_1'
            name = re.sub(r'unnamed:\s*\d+_level_\d+', ' ', name, flags=re.IGNORECASE)
            name = re.sub(r'\s+', ' ', name).strip()
            if not name:
                name = 'Без названия'
            if name in counts:
                counts[name] += 1
                name = f'{name} ({counts[name]})'
            else:
                counts[name] = 0
            names.append(name)
        df.columns = names
        return df

    @staticmethod
    def _detect_header(raw):
        """Строка шапки = та, где больше всего «шапочных» слов."""
        best_row, best_score = 0, 0
        for row in range(min(40, len(raw))):
            values = [str(v).strip().lower() for v in raw.iloc[row].tolist()
                      if pd.notna(v)]
            score = sum(1 for v in values
                        if any(word in v for word in HEADER_WORDS))
            if score > best_score:
                best_row, best_score = row, score

        if best_score < 3:
            return 0

        # над подзаголовками может быть строка-группа («Остаток на начало» ...)
        if best_row > 0:
            values = [str(v).strip().lower() for v in raw.iloc[best_row - 1].tolist()
                      if pd.notna(v)]
            score = sum(1 for v in values if any(word in v for word in GROUP_WORDS))
            if score >= 2:
                return [best_row - 1, best_row]

        return best_row

    @staticmethod
    def _sheet_order(sheets, exact=(), contains=()):
        """Порядок листов: сначала точные имена, потом по слову, потом остальные."""
        ordered = [s for s in sheets if str(s).strip() in exact]
        for word in contains:
            ordered += [s for s in sheets
                        if word in str(s).lower() and s not in ordered]
        ordered += [s for s in sheets if s not in ordered]
        return ordered

    def _read_table(self, path, exact=(), contains=()):
        """Читаем Excel: сами выбираем лист и шапку, не полагаясь на их названия."""
        excel = pd.ExcelFile(path)
        sheets = list(excel.sheet_names)

        fallback = None
        for sheet in self._sheet_order(sheets, exact, contains):
            try:
                raw = pd.read_excel(path, sheet_name=sheet, header=None)
            except Exception:
                continue
            if raw.empty:
                continue

            header = self._detect_header(raw)
            try:
                df = pd.read_excel(path, sheet_name=sheet, header=header)
            except Exception:
                df = raw
            df = self._flatten_columns(df)

            if fallback is None:
                fallback = (df, sheet)
            # берём первый лист, где вообще есть колонка материала
            if self._find_material_column(df) is not None:
                return df, sheet

        if fallback is None:
            raise ValueError('в файле нет читаемых листов: '
                             + ', '.join(str(s) for s in sheets))
        return fallback

    def _load_etalon(self, path):
        """Эталон -> таблица с колонками code/name/unit/price/остаток/приход/расход/конец."""
        df, sheet = self._read_table(path)
        self._diag = [
            'Excel Merger — диагностика',
            '',
            '=== ЭТАЛОН ===',
            f'Файл: {path}',
            f'Лист: {sheet}',
            f'Колонок: {len(df.columns)}',
        ]
        for i, col in enumerate(df.columns):
            self._diag.append(f'  [{i}] {col}')
        self._diag.append('')

        code_col = self._find_first_column(df, ('код', 'материал'))
        name_col = self._find_first_column(df, MATERIAL_NAME_KEYWORDS)
        price_col = self._find_first_column(df, ('цена',))
        unit_col = self._find_first_column(df, ('беи', 'ед. изм', 'ед.изм', 'единица'))
        warehouse_col = self._find_first_column(df, ('склд', 'склад'))
        plant_col = self._find_first_column(df, ('з-д', 'завод'))

        metrics = {name: self._find_metric_column(df, keys)
                   for name, keys in METRIC_KEYWORDS.items()}
        missing = [name for name in ('остаток', 'приход', 'расход')
                   if metrics.get(name) is None]
        if missing:
            raise ValueError('в эталоне нет колонок: ' + ', '.join(missing))

        if name_col is None:
            name_col = code_col
        if name_col is None:
            raise ValueError('в эталоне не найдена колонка с материалом')

        data = pd.DataFrame(index=df.index)
        data['code'] = df[code_col].apply(_norm_code) if code_col else ''
        data['code_text'] = data['code']
        data['name'] = df[name_col].apply(_norm_name)
        data['raw_name'] = df[name_col]
        data['unit'] = df[unit_col] if unit_col else ''
        data['price'] = df[price_col].apply(self._clean_number) if price_col else 0
        for name in METRIC_KEYWORDS:
            column = metrics.get(name)
            data[name] = df[column].apply(self._clean_number) if column else 0
        # склад и завод приводим к числам: 1С пишет их то числом, то текстом
        data['warehouse'] = (df[warehouse_col].apply(self._to_number)
                             if warehouse_col else None)
        data['plant'] = df[plant_col] if plant_col else None

        # строки без кода и названия не нужны
        data = data[(data['code'] != '') | (data['name'] != '')]
        if data.empty:
            raise ValueError('в эталоне нет строк с материалами')
        return data

    # ---------- разбор формы отчёта (openpyxl) ----------

    @staticmethod
    def _row_values(ws, row, max_col):
        values = []
        for column in range(1, max_col + 1):
            value = ws.cell(row, column).value
            if value is None:
                continue
            values.append(re.sub(r'\s+', ' ', str(value)).strip().lower())
        return values

    def _report_layout(self, ws):
        """Определяет шапку, колонки и строки данных формы. None, если это не таблица."""
        max_row = ws.max_row
        max_col = ws.max_column
        if max_row < 3 or max_col < 3:
            return None

        best_score, best_row = 0, 0
        for row in range(1, min(max_row, 60) + 1):
            values = self._row_values(ws, row, max_col)
            score = sum(1 for v in values if any(w in v for w in HEADER_WORDS))
            if score > best_score:
                best_score, best_row = score, row

        if best_score < 3:
            return None

        # строка-группа над подзаголовками может быть через пустую строку
        header_rows = [best_row]
        for row in range(max(1, best_row - 4), best_row):
            values = self._row_values(ws, row, max_col)
            score = sum(1 for v in values if any(w in v for w in GROUP_WORDS))
            if score >= 2:
                header_rows.append(row)
        header_rows.sort()

        labels = {}
        for column in range(1, max_col + 1):
            parts = []
            for row in header_rows:
                value = ws.cell(row, column).value
                if value is None:
                    continue
                text = re.sub(r'\s+', ' ', str(value)).strip().lower()
                if text and (not parts or parts[-1] != text):
                    parts.append(text)
            labels[column] = ' '.join(parts)

        def find(keys):
            for key in keys:
                for column in range(1, max_col + 1):
                    if _is_unit_column(labels[column]):
                        continue
                    if key in labels[column]:
                        return column
            return None

        name_col = find(MATERIAL_NAME_KEYWORDS) or find(('материал',))
        if name_col is None:
            return None

        columns = {
            'name': name_col,
            'code': find(('код',)),
            'unit': find(('ед. изм', 'ед.изм', 'единица измерения', 'единица')),
            'okei': find(('океи',)),
            'price': find(('цена',)),
            'seq': find(('номер по порядку',)),
            'plant': find(('з-д', 'завод')),
            'osn': find(METRIC_KEYWORDS['остаток']),
            'pri': find(METRIC_KEYWORDS['приход']),
            'ras': find(METRIC_KEYWORDS['расход']),
            'konec': find(METRIC_KEYWORDS['конец']),
        }
        for required in ('osn', 'pri', 'ras'):
            if columns[required] is None:
                return None

        rows = []
        started = False
        for row in range(best_row + 1, max_row + 1):
            value = ws.cell(row, name_col).value
            text = str(value).strip() if value is not None else ''
            if text.lower().startswith(('итого', 'всего')):
                break
            # «1», «2» — это строка нумерации колонок, а не материал
            numeric = bool(text) and re.sub(r'[.,\s]', '', text).isdigit()
            if text and not numeric:
                started = True
                rows.append(row)
            elif started:
                break

        if not rows:
            return None

        return {'header_rows': header_rows, 'rows': rows, 'col': columns,
                'labels': labels}

    @staticmethod
    def _report_warehouse(ws):
        """Склад из шапки формы («склад | 7720»)."""
        limit_row = min(ws.max_row, 40)
        for row in range(1, limit_row + 1):
            for column in range(1, ws.max_column + 1):
                value = ws.cell(row, column).value
                if isinstance(value, str) and value.strip().lower() in ('склад', 'склд'):
                    for shift in range(1, 4):
                        if column + shift > ws.max_column:
                            break
                        neighbour = ws.cell(row, column + shift).value
                        number = ExcelMergerApp._to_number(neighbour)
                        if number is not None and float(number).is_integer():
                            return int(number)
        return None

    # ---------- запись результата ----------

    @staticmethod
    def _insert_rows(ws, at, count):
        """Вставляет count строк перед строкой at, сохраняя стили и объединения.

        openpyxl.insert_rows не переносит объединённые ячейки и высоты строк,
        поэтому сдвигаем блок вниз вручную.
        """
        max_row = ws.max_row
        max_col = ws.max_column

        merges = [(m.min_col, m.min_row, m.max_col, m.max_row)
                  for m in ws.merged_cells.ranges if m.min_row >= at]
        for merge in list(ws.merged_cells.ranges):
            if merge.min_row >= at:
                ws.unmerge_cells(str(merge))

        for row in range(max_row, at - 1, -1):
            for column in range(1, max_col + 1):
                source = ws.cell(row, column)
                target = ws.cell(row + count, column)
                target.value = source.value
                if source.has_style:
                    target._style = copy(source._style)
                source.value = None
            height = ws.row_dimensions[row].height
            ws.row_dimensions[row + count].height = height
            ws.row_dimensions[row].height = None

        for min_col, min_row, max_col_, max_row_ in merges:
            ws.merge_cells(start_row=min_row + count, start_column=min_col,
                           end_row=max_row_ + count, end_column=max_col_)

    @staticmethod
    def _shift_formula_rows(formula, at, count):
        def replace(match):
            column, dollar, row = match.group(1), match.group(2), int(match.group(3))
            if row >= at:
                row += count
            return f'{column}{dollar}{row}'
        return FORMULA_REF.sub(replace, formula)

    def _fix_formulas(self, ws, at, count):
        """После вставки строк: сдвигаем ссылки и расширяем диапазоны SUM."""
        last_data = at - 1
        for row in range(at, ws.max_row + 1):
            for cell in ws[row]:
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    cell.value = self._shift_formula_rows(cell.value, at, count)

        # строка итогов теперь ниже вставленных строк
        pattern = re.compile(r'(\$?[A-Z]{1,3}\$?)%d(?!\d)' % last_data)
        for row in range(at + count, min(at + count + 5, ws.max_row + 1)):
            for cell in ws[row]:
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    cell.value = pattern.sub(
                        lambda m: f'{m.group(1)}{last_data + count}', cell.value)

    @staticmethod
    def _copy_row_style(ws, source_row, target_row, max_col):
        for column in range(1, max_col + 1):
            source = ws.cell(source_row, column)
            if source.has_style:
                ws.cell(target_row, column)._style = copy(source._style)
        height = ws.row_dimensions[source_row].height
        if height:
            ws.row_dimensions[target_row].height = height

    @staticmethod
    def _retarget_formula(formula, old_row, new_row):
        def replace(match):
            column, dollar, row = match.group(1), match.group(2), int(match.group(3))
            if row == old_row:
                row = new_row
            return f'{column}{dollar}{row}'
        return FORMULA_REF.sub(replace, formula)

    def _append_rows(self, ws, layout, records):
        """Добавляет строки новых материалов перед строкой итогов."""
        rows = layout['rows']
        columns = layout['col']
        max_col = ws.max_column
        style_row = rows[-1]
        at = style_row + 1
        count = len(records)

        self._insert_rows(ws, at, count)
        self._fix_formulas(ws, at, count)

        # продолжаем нумерацию строк
        numbers = [self._to_number(ws.cell(row, columns['seq']).value)
                   for row in rows if columns.get('seq')]
        numbers = [number for number in numbers if number is not None]
        next_number = int(max(numbers)) + 1 if numbers else 1

        # ОКЕИ по единице измерения — из уже заполненных строк формы
        okei = {}
        if columns.get('okei') and columns.get('unit'):
            for row in rows:
                unit = ws.cell(row, columns['unit']).value
                code = ws.cell(row, columns['okei']).value
                key = str(unit).strip().upper() if unit is not None else ''
                if key and code is not None and key not in okei:
                    okei[key] = code

        plant = None
        plant_col = columns.get('plant')
        if plant_col is None and columns.get('seq'):
            # в форме М-19 завод идёт сразу после номера строки
            plant_col = columns['seq'] + 1
        if plant_col:
            plant = ws.cell(style_row, plant_col).value
        # в форме завод записан текстом ('4301'), а из эталона приходит числом
        plant_as_text = isinstance(plant, str)

        for index, record in enumerate(records):
            row = at + index
            self._copy_row_style(ws, style_row, row, max_col)

            if columns.get('seq'):
                ws.cell(row, columns['seq']).value = next_number + index
            if plant_col:
                value = record.get('plant')
                if value is None or (isinstance(value, float) and pd.isna(value)):
                    value = plant
                if plant_as_text and self._to_number(value) is not None:
                    number = self._to_number(value)
                    value = str(int(number)) if number.is_integer() else str(number)
                ws.cell(row, plant_col).value = value

            ws.cell(row, columns['name']).value = record['raw_name']
            if columns.get('code'):
                cell = ws.cell(row, columns['code'])
                cell.value = record['code_text']
                cell.number_format = '@'
            if columns.get('unit'):
                ws.cell(row, columns['unit']).value = record['unit']
            if columns.get('okei'):
                key = str(record.get('unit') or '').strip().upper()
                ws.cell(row, columns['okei']).value = okei.get(key)
            if columns.get('price'):
                ws.cell(row, columns['price']).value = record['price']

            ws.cell(row, columns['osn']).value = record['остаток']
            ws.cell(row, columns['pri']).value = record['приход']
            ws.cell(row, columns['ras']).value = record['расход']

            # формулы (суммы, «остаток на конец») копируем из последней строки
            for column in range(1, max_col + 1):
                formula = ws.cell(style_row, column).value
                if isinstance(formula, str) and formula.startswith('='):
                    ws.cell(row, column).value = self._retarget_formula(
                        formula, style_row, row)

    def _merge_into_report(self, etalon):
        path = self.edit_path
        extension = os.path.splitext(str(path))[1].lower()
        if extension not in ('.xlsx', '.xlsm'):
            raise ValueError('чтобы сохранить оформление, отчёт должен быть '
                             'в формате .xlsx — пересохраните его из Excel')
        workbook = openpyxl.load_workbook(path, keep_vba=(extension == '.xlsm'))

        worksheet = layout = None
        for sheet in workbook.worksheets:
            found = self._report_layout(sheet)
            if found:
                worksheet, layout = sheet, found
                break
        if worksheet is None:
            raise ValueError('в файле для редактирования не найдена таблица '
                             'с материалами (нужны колонки «наименование», '
                             '«остаток», «приход», «расход»)')

        columns = layout['col']
        rows = layout['rows']

        self._diag += [
            '=== ФАЙЛ ДЛЯ РЕДАКТИРОВАНИЯ ===',
            f'Файл: {path}',
            f'Лист: {worksheet.title}',
            f'Строки шапки: {layout["header_rows"]}',
            f'Строк с материалами: {len(rows)} ({rows[0]}..{rows[-1]})',
            'Колонки: ' + ', '.join(
                f'{key}={value}' for key, value in columns.items() if value),
            '',
        ]

        warehouse = self._report_warehouse(worksheet)
        if warehouse is not None and 'warehouse' in etalon.columns:
            subset = etalon[etalon['warehouse'] == warehouse]
            if not subset.empty:
                etalon = subset
        self._diag.append(f'Склад формы: {warehouse}; '
                          f'строк эталона после фильтра: {len(etalon)}')
        self._diag.append('')

        records = []
        for _, row in etalon.iterrows():
            records.append({
                'code': row['code'],
                'code_text': row['code_text'],
                'name': row['name'],
                'raw_name': row['raw_name'],
                'unit': row['unit'],
                'price': row['price'],
                'plant': row['plant'],
                'остаток': row['остаток'],
                'приход': row['приход'],
                'расход': row['расход'],
                'конец': row['конец'],
                'used': False,
            })

        index_code = defaultdict(list)
        index_name = defaultdict(list)
        for position, record in enumerate(records):
            if record['code']:
                index_code[record['code']].append(position)
            if record['name']:
                index_name[record['name']].append(position)

        def take(key, index):
            for position in index.get(key, ()):
                if not records[position]['used']:
                    records[position]['used'] = True
                    return records[position]
            return None

        updated = 0
        untouched = 0
        matched = 0
        for row in rows:
            code = _norm_code(worksheet.cell(row, columns['code']).value) \
                if columns.get('code') else ''
            name = _norm_name(worksheet.cell(row, columns['name']).value)

            record = take(code, index_code) if code else None
            if record is None and name:
                record = take(name, index_name)
            if record is None:
                untouched += 1
                continue

            matched += 1
            changed = False
            for key, column in (('остаток', columns['osn']),
                                ('приход', columns['pri']),
                                ('расход', columns['ras'])):
                value = record[key]
                cell = worksheet.cell(row, column)
                if self._to_number(cell.value) != value:
                    cell.value = value
                    changed = True
            if changed:
                updated += 1

        new_records = [record for record in records if not record['used']]
        if new_records:
            self._append_rows(worksheet, layout, new_records)

        base_name = os.path.splitext(os.path.basename(path))[0]
        output_path = os.path.join(self._output_dir(),
                                   f'{base_name}_ОБНОВЛЕНО{extension}')
        try:
            workbook.calculation.fullCalcOnLoad = True
        except Exception:
            pass
        workbook.save(output_path)

        published = self._publish_to_download(output_path)

        return {
            'updated': updated,
            'matched': matched,
            'added': len(new_records),
            'untouched': untouched,
            'sheet': worksheet.title,
            'path': output_path,
            'published': published,
        }

    # ---------- запуск ----------

    def _process(self, instance):
        if not self.etalon_path or not self.edit_path:
            self._update_status('Ошибка: выберите оба файла')
            return

        # Длинная операция: сначала даём интерфейсу перерисовать статус
        self._update_status('Обрабатываю файлы…')
        Clock.schedule_once(self._run_merge, 0.1)

    def _run_merge(self, _dt):
        try:
            etalon = self._load_etalon(self.etalon_path)
            result = self._merge_into_report(etalon)
            message = (f'Готово! Строк с парой: {result["matched"]}, '
                       f'изменено: {result["updated"]}, '
                       f'добавлено: {result["added"]}.')
            if result['untouched']:
                message += f'\nБез пары в эталоне: {result["untouched"]}'
            message += (f'\nЛист: {result["sheet"]}\n'
                        f'Файл: {result["published"] or os.path.basename(result["path"])}')
            if result.get('published'):
                message += '\nОткройте папку «Загрузки»'
            self._update_status(message)
        except Exception as e:
            self._update_status(f'Ошибка: {e}\n'
                                f'Подробности: excel_merger_diagnostics.txt')
            self._write_diagnostics(e)

    def _write_diagnostics(self, error):
        """Пишем рядом с результатом расшифровку — по ней видно, что не так."""
        lines = list(self._diag) + ['=== ОШИБКА ===', str(error)]
        try:
            path = os.path.join(self._output_dir(), 'excel_merger_diagnostics.txt')
            with open(path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(lines))
            self._publish_to_download(path)
        except Exception:
            pass


if __name__ == '__main__':
    ExcelMergerApp().run()
