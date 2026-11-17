"""Excel Merger — слияние эталонного файла и файла за период."""

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

ANDROID = platform == 'android'

# Для системного выбора файла (Storage Access Framework)
if ANDROID:
    from android import mActivity
    from android import activity as android_activity
    from android.permissions import request_permissions, Permission
    from jnius import autoclass

REQUEST_OPEN_DOCUMENT = 0x0E01
RESULT_OK = -1

EXCEL_MIME_TYPES = [
    'application/vnd.ms-excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
]

# ВАЖНО: все размеры в dp/sp, иначе на телефоне с плотностью 3
# «height=44» превращается в ~15 реальных dp и текст не влезает.
BTN_HEIGHT = dp(52)
BTN_FONT = sp(17)
STATUS_FONT = sp(15)
STATUS_MIN_HEIGHT = dp(72)
STATUS_MAX_HEIGHT = dp(160)

STATUS_START = 'Загрузите эталон и файл для редактирования'


def _start_dir():
    """Стартовая папка для файлового менеджера (только для ПК)."""
    if ANDROID and os.path.isdir('/storage/emulated/0'):
        return '/storage/emulated/0'
    return os.path.expanduser('~')


def _output_dir():
    """Первая папка, в которую реально можно писать."""
    candidates = []
    if ANDROID:
        candidates += [
            '/storage/emulated/0/Download',
            '/sdcard/Download',
            os.path.join(os.path.expanduser('~'), 'Download'),
        ]
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


class ExcelMergerApp(App):
    _pending_type = None

    def build(self):
        self.etalon_path = None
        self.edit_path = None

        Window.clearcolor = (0.96, 0.96, 0.96, 1)

        if ANDROID:
            # SAF не требует разрешений, но доступ к папке Download на Android 11+
            # зависит от «Все файлы» — просим то, что можно попросить программно.
            try:
                request_permissions(
                    [Permission.READ_EXTERNAL_STORAGE,
                     Permission.WRITE_EXTERNAL_STORAGE]
                )
            except Exception:
                pass
            try:
                android_activity.bind(on_activity_result=self._on_activity_result)
            except Exception:
                pass

        root = BoxLayout(orientation='vertical', padding=dp(16), spacing=dp(12))

        self.status = Label(
            text=STATUS_START,
            size_hint_y=None,
            height=STATUS_MIN_HEIGHT,
            halign='center',
            valign='middle',
            color=(0.12, 0.12, 0.12, 1),
            font_size=STATUS_FONT,
            padding=(dp(8), dp(8)),
        )
        # halign работает только вместе с text_size; высота подстраивается под текст
        self.status.bind(size=self._sync_status_text_size)
        self.status.bind(texture_size=self._sync_status_height)
        root.add_widget(self.status)

        btn_etalon = Button(text='1. Выбрать эталон', size_hint_y=None,
                            height=BTN_HEIGHT, font_size=BTN_FONT)
        btn_etalon.bind(on_press=self._select_file('etalon'))
        root.add_widget(btn_etalon)

        btn_edit = Button(text='2. Выбрать файл для редактирования', size_hint_y=None,
                          height=BTN_HEIGHT, font_size=BTN_FONT)
        btn_edit.bind(on_press=self._select_file('edit'))
        root.add_widget(btn_edit)

        btn_run = Button(text='3. Объединить и сохранить', size_hint_y=None,
                         height=BTN_HEIGHT, font_size=BTN_FONT)
        btn_run.bind(on_press=self._process)
        root.add_widget(btn_run)

        # BoxLayout(vertical) укладывает детей от нижнего края вверх,
        # поэтому без распорки кнопки «прилипают» к низу экрана.
        root.add_widget(Widget())

        return root

    # ---------- интерфейс ----------

    def _sync_status_text_size(self, instance, size):
        instance.text_size = (max(0, size[0] - dp(8)), None)

    def _sync_status_height(self, instance, texture_size):
        height = texture_size[1] + dp(24)
        instance.height = min(STATUS_MAX_HEIGHT, max(STATUS_MIN_HEIGHT, height))

    def _update_status(self, message=None):
        if message is not None:
            self.status.text = message
            return
        etalon = os.path.basename(self.etalon_path) if self.etalon_path else 'не выбран'
        edit = os.path.basename(self.edit_path) if self.edit_path else 'не выбран'
        self.status.text = f'Эталон: {etalon}\nФайл: {edit}'

    # ---------- выбор файла ----------

    def _select_file(self, file_type):
        """Функция-обертка для вызова диалога."""
        def inner(instance):
            if ANDROID:
                self._open_android_picker(file_type)
            else:
                self._open_desktop_picker(file_type)
        return inner

    def _open_desktop_picker(self, file_type):
        chooser_popup = Popup(title='Выберите Excel файл', size_hint=(0.95, 0.95))
        content = BoxLayout(orientation='vertical', spacing=dp(8), padding=dp(8))
        chooser = FileChooserListView(filters=['*.xlsx', '*.xls'], path=_start_dir())

        def on_select(instance):
            selection = chooser.selection
            if not selection:
                return
            chooser_popup.dismiss()
            self._set_path(file_type, os.path.join(chooser.path, selection[0]))

        btn_select = Button(text='Выбрать', size_hint_y=None, height=BTN_HEIGHT,
                            font_size=BTN_FONT)
        btn_select.bind(on_press=on_select)
        content.add_widget(chooser)
        content.add_widget(btn_select)
        chooser_popup.content = content
        chooser_popup.open()

    def _open_android_picker(self, file_type):
        """Системный диалог Android: работает без доступа ко всей файловой системе."""
        try:
            Intent = autoclass('android.content.Intent')
            intent = Intent(Intent.ACTION_OPEN_DOCUMENT)
            intent.addCategory(Intent.CATEGORY_OPENABLE)
            intent.setType('*/*')
            try:
                intent.putExtra(Intent.EXTRA_MIME_TYPES, EXCEL_MIME_TYPES)
            except Exception:
                pass
            self._pending_type = file_type
            mActivity.startActivityForResult(intent, REQUEST_OPEN_DOCUMENT)
        except Exception as e:
            self._update_status(f'Ошибка открытия выбора файла: {e}')

    def _on_activity_result(self, request_code, result_code, intent):
        # Колбэк приходит из Java-потока: Kivy-виджеты обновляем через Clock.
        if request_code != REQUEST_OPEN_DOCUMENT or not self._pending_type:
            return
        file_type = self._pending_type
        self._pending_type = None

        if result_code != RESULT_OK or intent is None or intent.getData() is None:
            Clock.schedule_once(lambda dt: self._update_status('Файл не выбран'))
            return

        try:
            path = self._copy_uri_to_file(intent.getData())
        except Exception as e:
            message = f'Не удалось прочитать файл: {e}'
            Clock.schedule_once(lambda dt: self._update_status(message))
            return

        Clock.schedule_once(lambda dt: self._set_path(file_type, path))

    def _copy_uri_to_file(self, uri):
        """content:// из SAF нельзя открыть напрямую — копируем во внутреннюю папку."""
        resolver = mActivity.getContentResolver()
        name = self._uri_display_name(uri, resolver)

        dest_dir = os.path.join(self.user_data_dir, 'input')
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

    def _uri_display_name(self, uri, resolver):
        name = None
        try:
            OpenableColumns = autoclass('android.provider.OpenableColumns')
            cursor = resolver.query(uri, None, None, None, None)
            if cursor is not None:
                try:
                    index = cursor.getColumnIndex(OpenableColumns.DISPLAY_NAME)
                    if index >= 0 and cursor.moveToFirst():
                        name = cursor.getString(index)
                finally:
                    cursor.close()
        except Exception:
            name = None

        if not name:
            name = 'input.xlsx'
        if not name.lower().endswith(('.xlsx', '.xls')):
            name += '.xlsx'
        return name

    def _set_path(self, file_type, path):
        if file_type == 'etalon':
            self.etalon_path = path
        else:
            self.edit_path = path
        self._update_status()

    # ---------- обработка Excel ----------

    def _clean_number(self, value):
        """Очистка чисел."""
        if value is None:
            return 0
        if isinstance(value, (int, float)):
            return 0 if pd.isna(value) else value
        s = re.sub(r'[^\d.,]', '', str(value)).replace(',', '.')
        if not s:
            return 0
        try:
            return float(s)
        except ValueError:
            return 0

    def _find_columns(self, df, keywords):
        """Гибкий поиск колонок по подстроке: {ключ: имя колонки}."""
        found = {}
        columns = list(df.columns)
        lowered = [str(col).lower() for col in columns]
        for key in keywords:
            key_low = key.lower()
            for i, col in enumerate(lowered):
                if key_low in col:
                    found[key] = columns[i]
                    break
        return found

    @staticmethod
    def _flat_columns(df):
        """Многоуровневые заголовки -> обычные строки."""
        df.columns = [' '.join(map(str, col)).strip()
                      if isinstance(col, tuple) else str(col).strip()
                      for col in df.columns]
        return df

    def _read_etalon(self, path):
        try:
            return self._flat_columns(pd.read_excel(path, header=[0, 1]))
        except Exception:
            return self._flat_columns(pd.read_excel(path, header=0))

    def _read_edit(self, path):
        excel = pd.ExcelFile(path)
        sheets = excel.sheet_names
        sheet = next((s for s in sheets if 'сентябрь' in str(s).lower()), sheets[0])
        try:
            return self._flat_columns(pd.read_excel(path, sheet_name=sheet, header=[5, 6]))
        except Exception:
            return self._flat_columns(pd.read_excel(path, sheet_name=sheet, header=0))

    def _process(self, instance):
        if not self.etalon_path or not self.edit_path:
            self._update_status('Ошибка: выберите оба файла')
            return

        try:
            df_etalon = self._read_etalon(self.etalon_path)
            df_edit = self._read_edit(self.edit_path)

            # --- ищем ключевые колонки, не полагаясь на точные названия ---
            candidate = self._find_columns(df_etalon, ['материал'])
            mat_etalon = candidate.get('материал')
            if mat_etalon is None:
                raise ValueError('в эталоне не найдена колонка «Материал»')

            candidate = self._find_columns(df_edit, ['наименование,', 'наименование', 'материал'])
            mat_edit = (candidate.get('наименование,')
                        or candidate.get('наименование')
                        or candidate.get('материал'))
            if mat_edit is None:
                raise ValueError('в файле периода не найдена колонка с наименованием материала')

            # соответствие: колонка в отчёте <- колонка в эталоне
            update_map = {
                'остаток на начало': 'остзапнп',
                'приход': 'постматер',
                'расход': 'отпуск мат',
            }
            found_etalon = self._find_columns(df_etalon, list(update_map.values()))
            found_edit = self._find_columns(df_edit, list(update_map.keys()))

            # чистим числовые колонки эталона от «1 234,50 ₽»
            for source_key, source_col in found_etalon.items():
                df_etalon[source_col] = df_etalon[source_col].apply(self._clean_number)

            df_etalon_i = df_etalon.set_index(mat_etalon)
            df_etalon_i = df_etalon_i[~df_etalon_i.index.duplicated(keep='first')]
            df_edit_i = df_edit.set_index(mat_edit)

            # 1. Добавляем новые материалы из эталона в отчёт
            new_materials = df_etalon_i.index.difference(df_edit_i.index)
            if not new_materials.empty:
                df_edit_i = pd.concat([df_edit_i, df_etalon_i.loc[new_materials]])

            # 2. Обновляем данные для существующих материалов
            common_materials = df_etalon_i.index.intersection(df_edit_i.index)
            updated = 0

            for target_key, source_key in update_map.items():
                target_col = found_edit.get(target_key)
                source_col = found_etalon.get(source_key)
                if target_col is None or source_col is None:
                    continue
                # Цену в отчёте НЕ ТРОГАЕМ: обновляем только остаток/приход/расход
                values = df_etalon_i[source_col].reindex(df_edit_i.index)
                mask = values.notna().values
                df_edit_i.loc[mask, target_col] = values[mask].values
                updated += 1

            if updated == 0:
                raise ValueError('не найдены колонки остатков/прихода/расхода — '
                                 'проверьте структуру файлов')

            result_df = df_edit_i.reset_index()

            base_name = os.path.splitext(os.path.basename(self.edit_path))[0]
            output_path = os.path.join(_output_dir(), f'{base_name}_ОБНОВЛЕНО.xlsx')
            result_df.to_excel(output_path, index=False)

            # Длинный путь в статус не влезет — показываем только имя файла
            self._update_status(
                f'Готово! Добавлено: {len(new_materials)}, '
                f'обновлено колонок: {updated}.\n'
                f'Файл: {os.path.basename(output_path)}'
            )

        except Exception as e:
            self._update_status(f'Ошибка: {e}')


if __name__ == '__main__':
    ExcelMergerApp().run()
