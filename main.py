import os
import re

import pandas as pd
from kivy.app import App
from kivy.clock import Clock
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.utils import platform

# Для работы с системным выбором файла
if platform == 'android':
    from android import activity as android_activity
    from android.permissions import request_permissions, Permission
    from android.storage import app_storage_path, primary_external_storage_path
    from jnius import autoclass

REQUEST_OPEN_DOCUMENT = 1001
RESULT_OK = -1


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

        layout = BoxLayout(orientation='vertical', padding=10, spacing=10)

        self.status = Label(text='Загрузите эталон и файл для редактирования', size_hint_y=None, height=44)
        layout.add_widget(self.status)

        btn_etalon = Button(text='Выбрать эталон', size_hint_y=None, height=44)
        btn_etalon.bind(on_press=self._select_file('etalon'))
        layout.add_widget(btn_etalon)

        btn_edit = Button(text='Выбрать файл для редактирования', size_hint_y=None, height=44)
        btn_edit.bind(on_press=self._select_file('edit'))
        layout.add_widget(btn_edit)

        btn_run = Button(text='Объединить и сохранить', size_hint_y=None, height=44)
        btn_run.bind(on_press=self._process)
        layout.add_widget(btn_run)

        return layout

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
                chooser_popup = Popup(title='Выберите Excel файл')
                chooser = FileChooserListView(filters=['*.xlsx', '*.xls'])
                btn_select = Button(text='Выбрать', size_hint_y=None, height=44)

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

    def _process(self, instance):
        if not self.etalon_path or not self.edit_path:
            self._update_status('Ошибка: выберите оба файла')
            return

        try:
            # Читаем файлы с учетом их сложной структуры
            df_etalon = pd.read_excel(self.etalon_path, header=[0, 1])
            df_edit = pd.read_excel(self.edit_path, sheet_name='Сентябрь_table_4', header=[5, 6])

            # Превращаем многоуровневые заголовки в обычные строки
            df_etalon.columns = [' '.join(map(str, col)).strip() for col in df_etalon.columns]
            df_edit.columns = [' '.join(map(str, col)).strip() for col in df_edit.columns]

            # Ищем колонку с материалом в обоих файлах
            etalon_material = self._find_columns(df_etalon, ['материал']).get('материал')
            edit_material = (self._find_columns(df_edit, ['наименование']).get('наименование')
                             or self._find_columns(df_edit, ['материал']).get('материал'))
            if etalon_material is None or edit_material is None:
                raise ValueError('не найдена колонка с материалом')

            df_etalon_i = df_etalon.set_index(etalon_material)
            df_edit_i = df_edit.set_index(edit_material)

            # 1. Добавляем новые материалы из эталона в сентябрь
            new_materials = df_etalon_i.index.difference(df_edit_i.index)
            if not new_materials.empty:
                df_new_rows = df_etalon_i.loc[new_materials]
                df_edit_i = pd.concat([df_edit_i, df_new_rows])

            # 2. Обновляем данные для существующих материалов
            common_materials = df_etalon_i.index.intersection(df_edit_i.index)

            # ВАЖНО: Обновляем только Остатки, Приход и Расход. Цену в "Сентябре" НЕ ТРОГАЕМ.
            update_map = {
                'остаток': 'ОстЗапНП',
                'приход': 'ПостМатер',
                'расход': 'Отпуск мат'
            }

            for keyword, source_col in update_map.items():
                if source_col not in df_etalon_i.columns:
                    continue
                target_col = self._find_columns(df_edit_i, [keyword]).get(keyword)
                if target_col is None:
                    continue
                df_etalon_i[source_col] = df_etalon_i[source_col].apply(self._clean_number)
                df_edit_i.loc[common_materials, target_col] = df_etalon_i.loc[common_materials, source_col]

            # Сбрасываем индекс
            result_df = df_edit_i.reset_index()

            # Сохраняем результат
            base_name = os.path.splitext(os.path.basename(self.edit_path))[0]
            output_path = os.path.join(self._output_dir(), f'{base_name}_ОБНОВЛЕНО.xlsx')
            result_df.to_excel(output_path, index=False)

            self._update_status(
                f'Готово! Добавлено: {len(new_materials)}, обновлено: {len(common_materials)}.\n'
                f'Файл: {os.path.basename(output_path)}'
            )

        except Exception as e:
            self._update_status(f'Ошибка: {e}')


if __name__ == '__main__':
    ExcelMergerApp().run()
