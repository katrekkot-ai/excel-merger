import os
import re

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

# Признаки строки шапки: ищем не по точному имени листа/строки, а по содержимому,
# иначе любое переименование листа («Сентябрь_table_4») ломает всё приложение.
MATERIAL_HEADER_PREFIXES = ('материал', 'наименование', 'краткий текст',
                            'название', 'номенклатура')
DATA_HEADER_PREFIXES = ('остаток', 'приход', 'расход', 'цена', 'количество', 'кол-во', 'ед.')

# Что переносим из эталона в отчёт. Ключи идут по приоритету: сначала точные
# названия из эталона, потом «человеческие» из отчёта.
METRIC_KEYWORDS = {
    'остаток': ('остзапнп', 'остаток на начало', 'остаток'),
    'приход': ('постматер', 'приход'),
    'расход': ('отпуск мат', 'расход'),
}

# Колонка-единица измерения: «Ед.», «Ед. Unnamed: 2_level_1» и т.п.
UNIT_TOKENS = {'ед', 'шт', 'кг', 'т', 'л', 'мл', 'м', 'уп', 'пач', 'рул'}

# Как называется колонка с названием материала (по приоритету).
MATERIAL_NAME_KEYWORDS = ('краткий текст материала', 'наименование',
                          'название', 'номенклатура')


def _is_unit_column(name):
    """True, если в заголовке нет ничего, кроме единицы измерения."""
    tokens = [t for t in re.split(r'[^0-9a-zа-я]+', str(name).lower()) if t]
    tokens = [t for t in tokens
              if not t.isdigit() and t not in ('unnamed', 'level', 'уровень')]
    return bool(tokens) and all(t in UNIT_TOKENS for t in tokens)

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
        """Очистка чисел."""
        if pd.isna(value):
            return 0
        if isinstance(value, (int, float)):
            return value
        s = re.sub(r'[^\d.,]', '', str(value).replace(',', '.'))
        try:
            return float(s)
        except ValueError:
            return 0

    def _find_columns(self, df, keywords):
        """Гибкий поиск колонок."""
        df_columns = [str(col).lower() for col in df.columns]
        found = {}
        for key in keywords:
            for i, col in enumerate(df_columns):
                if key in col:
                    found[key] = df.columns[i]
                    break
        return found

    def _find_metric_column(self, df, keywords):
        """Колонка метрики (остаток/приход/расход) с учётом приоритета ключей.

        Ключи проверяются по очереди, поэтому 'остзапнп' важнее простого
        'остаток'. Колонки-единицы измерения («Ед.») пропускаются.
        """
        return self._find_first_column(df, keywords)

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

    def _find_material_column(self, df):
        """Колонка с наименованием материала.

        В выгрузках 1С колонка «Материал» — это код (000000000011316154),
        а название лежит в «Краткий текст материала». Поэтому сначала ищем
        человекочитаемое название и только потом падаем на код.
        """
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

    # ---------- чтение Excel ----------

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
        """Находим строку шапки по ключевым словам (лист мог быть переименован)."""
        limit = min(20, len(raw))
        for row in range(limit):
            values = [str(v).strip().lower() for v in raw.iloc[row].tolist()]
            if not any(v.startswith(MATERIAL_HEADER_PREFIXES) for v in values):
                continue
            if row + 1 < len(raw):
                next_values = [str(v).strip().lower() for v in raw.iloc[row + 1].tolist()]
                if any(v.startswith(DATA_HEADER_PREFIXES) for v in next_values):
                    return [row, row + 1]
            return row
        return 0

    def _read_sheet(self, path, sheet):
        """Читаем один лист, определяя строку шапки по содержимому."""
        try:
            raw = pd.read_excel(path, sheet_name=sheet, header=None)
        except Exception:
            return None
        if raw.empty:
            return None

        header = self._detect_header(raw)
        try:
            df = pd.read_excel(path, sheet_name=sheet, header=header)
        except Exception:
            df = raw
        return self._flatten_columns(df)

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
            df = self._read_sheet(path, sheet)
            if df is None:
                continue
            if fallback is None:
                fallback = (df, sheet)
            # берём первый лист, где вообще есть колонка материала
            if self._find_material_column(df) is not None:
                return df, sheet

        if fallback is None:
            raise ValueError('в файле нет читаемых листов: '
                             + ', '.join(str(s) for s in sheets))
        return fallback

    def _read_edit(self, path):
        return self._read_table(path, exact=('Сентябрь_table_4',),
                                contains=('сентябр',))

    def _read_etalon(self, path):
        return self._read_table(path)

    def _process(self, instance):
        if not self.etalon_path or not self.edit_path:
            self._update_status('Ошибка: выберите оба файла')
            return

        df_etalon = df_edit = None
        sheet_etalon = sheet_edit = None
        try:
            df_etalon, sheet_etalon = self._read_etalon(self.etalon_path)
            df_edit, sheet_edit = self._read_edit(self.edit_path)

            # Ищем колонку с материалом в обоих файлах
            etalon_material = self._find_material_column(df_etalon)
            edit_material = self._find_material_column(df_edit)
            if etalon_material is None:
                raise ValueError('в эталоне не найдена колонка с материалом')
            if edit_material is None:
                raise ValueError('в файле периода не найдена колонка с материалом')

            # Соответствие колонок: (что, колонка в отчёте, колонка в эталоне).
            # Ищем по ключевым словам с обеих сторон, потому что после склейки
            # двух строк шапки имена выглядят как 'ОстЗапНП ОстЗапНП'.
            mapping = []
            missing_etalon = []
            missing_edit = []
            for name, keys in METRIC_KEYWORDS.items():
                target_col = self._find_metric_column(df_edit, keys)
                source_col = self._find_metric_column(df_etalon, keys)
                if source_col is None:
                    missing_etalon.append(name)
                if target_col is None:
                    missing_edit.append(name)
                if target_col is not None and source_col is not None:
                    mapping.append((name, target_col, source_col))

            if not mapping:
                parts = []
                if missing_etalon:
                    parts.append('в эталоне — ' + ', '.join(missing_etalon))
                if missing_edit:
                    parts.append('в отчёте — ' + ', '.join(missing_edit))
                raise ValueError('не найдены колонки: ' + '; '.join(parts))

            # Числа в эталоне могут быть текстом («1 000,50»), чистим их.
            for _, _, source_col in mapping:
                df_etalon[source_col] = df_etalon[source_col].apply(self._clean_number)

            df_etalon_i = df_etalon.set_index(etalon_material)
            # Один материал в выгрузке 1С встречается несколько раз (склад, партия),
            # поэтому остатки/приход/расход суммируем по материалу.
            metric_cols = [source_col for _, _, source_col in mapping]
            etalon_values = df_etalon_i[metric_cols].groupby(level=0).sum()
            df_edit_i = df_edit.set_index(edit_material).copy()

            # 1. Добавляем новые материалы из эталона в сентябрь.
            new_materials = etalon_values.index.difference(df_edit_i.index)
            if not new_materials.empty:
                new_rows = pd.DataFrame(index=new_materials,
                                        columns=list(df_edit_i.columns),
                                        dtype=object)
                for _, target_col, source_col in mapping:
                    new_rows[target_col] = etalon_values.loc[new_materials, source_col].to_numpy()
                # иначе при concat теряется имя колонки материала
                new_rows.index = new_rows.index.rename(df_edit_i.index.name)
                df_edit_i = pd.concat([df_edit_i, new_rows])

            # 2. Обновляем данные для существующих материалов.
            # Цену в отчёте НЕ ТРОГАЕМ: только остаток/приход/расход.
            updated = 0
            for _, target_col, source_col in mapping:
                # map по индексу работает и при повторяющихся наименованиях
                # в отчёте, поэтому используем его, а не .loc/позиции.
                mapped = df_edit_i.index.to_series().map(etalon_values[source_col]).to_numpy()
                mask = pd.notna(mapped)
                if not mask.any():
                    continue
                values = df_edit_i[target_col].to_numpy().astype(object)
                values[mask] = mapped[mask]
                df_edit_i[target_col] = values
                updated += 1

            if updated == 0:
                raise ValueError('в эталоне нет данных для обновления '
                                 '(остаток/приход/расход пустые)')

            # Сбрасываем индекс
            result_df = df_edit_i.reset_index()

            # Сохраняем результат
            base_name = os.path.splitext(os.path.basename(self.edit_path))[0]
            output_path = os.path.join(self._output_dir(), f'{base_name}_ОБНОВЛЕНО.xlsx')
            result_df.to_excel(output_path, index=False)

            self._update_status(
                f'Готово! Добавлено: {len(new_materials)}, '
                f'обновлено колонок: {updated}.\n'
                f'Лист: {sheet_edit}\n'
                f'Файл: {os.path.basename(output_path)}'
            )

        except Exception as e:
            self._update_status(f'Ошибка: {e}\n'
                                f'Подробности: excel_merger_diagnostics.txt')
            self._write_diagnostics(e, sheet_etalon, df_etalon,
                                    sheet_edit, df_edit)

    def _write_diagnostics(self, error, sheet_etalon, df_etalon,
                           sheet_edit, df_edit):
        """Пишем рядом с результатом список колонок — по нему видно, что не так."""
        lines = ['Excel Merger — диагностика', f'Ошибка: {error}', '']

        for title, sheet, df in (('ЭТАЛОН', sheet_etalon, df_etalon),
                                 ('ФАЙЛ ПЕРИОДА', sheet_edit, df_edit)):
            lines.append(f'=== {title} ===')
            lines.append(f'Файл: {self.etalon_path if df is df_etalon else self.edit_path}')
            lines.append(f'Лист: {sheet if sheet is not None else "(не определён)"}')
            if df is not None:
                lines.append(f'Колонок: {len(df.columns)}')
                for i, col in enumerate(df.columns):
                    lines.append(f'  [{i}] {col}')
            lines.append('')

        try:
            path = os.path.join(self._output_dir(), 'excel_merger_diagnostics.txt')
            with open(path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(lines))
        except Exception:
            pass


if __name__ == '__main__':
    ExcelMergerApp().run()
