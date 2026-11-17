import os

import pandas as pd
from kivy.app import App
from kivy.core.window import Window
from kivy.metrics import dp, sp
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.widget import Widget

ANDROID = 'ANDROID_ARGUMENT' in os.environ

STATUS_START = 'Загрузите эталон и файл для редактирования'

# ВАЖНО: все размеры задаём в dp/sp, а не в пикселях. На телефоне с плотностью 3
# «height=44» превращается в ~15 реальных dp, и текст в кнопку не влезает.
BTN_HEIGHT = dp(52)
BTN_FONT = sp(17)
STATUS_FONT = sp(15)
STATUS_MIN_HEIGHT = dp(72)
STATUS_MAX_HEIGHT = dp(160)


def _start_dir():
    """Стартовая папка для файлового менеджера."""
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


class FileChooserPopup(Popup):
    def __init__(self, callback, **kwargs):
        super().__init__(**kwargs)
        self.title = 'Выберите Excel файл'
        self.size_hint = (0.95, 0.95)
        self.callback = callback

        layout = BoxLayout(orientation='vertical', spacing=dp(8), padding=dp(8))
        chooser = FileChooserListView(filters=['*.xlsx', '*.xls'], path=_start_dir())
        btn_select = Button(text='Выбрать', size_hint_y=None, height=BTN_HEIGHT,
                            font_size=BTN_FONT)
        btn_select.bind(on_press=lambda x: self._select(chooser.path, chooser.selection))
        layout.add_widget(chooser)
        layout.add_widget(btn_select)
        self.content = layout

    def _select(self, path, selection):
        self.dismiss()
        if selection:
            self.callback(os.path.join(path, selection[0]))


class ExcelMergerApp(App):
    def build(self):
        self.etalon_path = None
        self.edit_path = None

        # Светлый фон: на чёрном подписи читались плохо
        Window.clearcolor = (0.96, 0.96, 0.96, 1)

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
            markup=True,
        )
        # halign работает только вместе с text_size; высота подстраивается под текст
        self.status.bind(size=self._sync_status_text_size)
        self.status.bind(texture_size=self._sync_status_height)
        root.add_widget(self.status)

        btn_etalon = Button(text='1. Выбрать эталон', size_hint_y=None,
                            height=BTN_HEIGHT, font_size=BTN_FONT)
        btn_etalon.bind(on_press=lambda x: self._open_chooser(self._set_etalon))
        root.add_widget(btn_etalon)

        btn_edit = Button(text='2. Выбрать файл для редактирования', size_hint_y=None,
                          height=BTN_HEIGHT, font_size=BTN_FONT)
        btn_edit.bind(on_press=lambda x: self._open_chooser(self._set_edit))
        root.add_widget(btn_edit)

        btn_run = Button(text='3. Объединить и сохранить', size_hint_y=None,
                         height=BTN_HEIGHT, font_size=BTN_FONT)
        btn_run.bind(on_press=self._process)
        root.add_widget(btn_run)

        # Распорка. BoxLayout с orientation='vertical' укладывает детей от нижнего
        # края вверх, поэтому без неё фиксированные по высоте кнопки «прилипают»
        # к низу экрана. Эта распорка съедает остаток высоты и прижимает блок к верху.
        root.add_widget(Widget())

        return root

    def _sync_status_text_size(self, instance, size):
        instance.text_size = (max(0, size[0] - dp(8)), None)

    def _sync_status_height(self, instance, texture_size):
        height = texture_size[1] + dp(24)
        instance.height = min(STATUS_MAX_HEIGHT, max(STATUS_MIN_HEIGHT, height))

    def _open_chooser(self, callback):
        self.chooser_popup = FileChooserPopup(callback)
        self.chooser_popup.open()

    def _set_etalon(self, path):
        self.etalon_path = path
        edit_name = os.path.basename(self.edit_path) if self.edit_path else 'не выбран'
        self.status.text = (f'Эталон: [b]{os.path.basename(path)}[/b]\n'
                            f'Редактируем: {edit_name}')

    def _set_edit(self, path):
        self.edit_path = path
        etalon_name = os.path.basename(self.etalon_path) if self.etalon_path else 'не выбран'
        self.status.text = (f'Эталон: {etalon_name}\n'
                            f'Редактируем: [b]{os.path.basename(path)}[/b]')

    def _clean_number(self, value):
        if pd.isna(value):
            return 0
        s = (str(value).replace('₽', '').replace(' ', '')
             .replace('\xa0', '').replace(',', '.'))
        try:
            if '.' in s:
                return float(s)
            else:
                return int(s)
        except ValueError:
            return 0

    def _process(self, instance):
        if not self.etalon_path or not self.edit_path:
            self.status.text = 'Ошибка: выберите оба файла'
            return

        try:
            # Читаем файлы с учетом их сложной структуры
            df_etalon = pd.read_excel(self.etalon_path, header=[0, 1])
            df_edit = pd.read_excel(self.edit_path, sheet_name='Сентябрь_table_4', header=[5, 6])

            # Превращаем многоуровневые заголовки в обычные строки
            df_etalon.columns = [' '.join(map(str, col)).strip() for col in df_etalon.columns]
            df_edit.columns = [' '.join(map(str, col)).strip() for col in df_edit.columns]

            # Очищаем числовые данные в эталоне от мусора (символы, пробелы)
            num_cols_etalon = ['ЦенаКПер', 'ОстЗапНП', 'ПостМатер', 'Отпуск мат']
            for col in num_cols_etalon:
                if col in df_etalon.columns:
                    df_etalon[col] = df_etalon[col].apply(self._clean_number)

            # Устанавливаем "Материал" как индекс для быстрого поиска
            # В эталоне столбец называется "Материал", в сентябре — "наименование, характеристика"
            df_etalon_i = df_etalon.set_index('Материал')
            df_edit_i = df_edit.set_index('наименование,\nхарактеристика')

            # 1. Добавляем новые материалы из эталона в сентябрь
            new_materials = df_etalon_i.index.difference(df_edit_i.index)
            if not new_materials.empty:
                df_new_rows = df_etalon_i.loc[new_materials]
                df_edit_i = pd.concat([df_edit_i, df_new_rows])

            # 2. Обновляем данные для существующих материалов
            common_materials = df_etalon_i.index.intersection(df_edit_i.index)

            # ВАЖНО: Обновляем только Остатки, Приход и Расход. Цену в "Сентябре" НЕ ТРОГАЕМ.
            update_map = {
                'Остаток на начало\n отчетного периода': 'ОстЗапНП',
                'Приход за отчетный период': 'ПостМатер',
                'Расход за отчетный период': 'Отпуск мат'
            }

            for target_col, source_col in update_map.items():
                if target_col in df_edit_i.columns and source_col in df_etalon_i.columns:
                    df_edit_i.loc[common_materials, target_col] = df_etalon_i.loc[common_materials, source_col]

            # Сбрасываем индекс
            result_df = df_edit_i.reset_index()

            # Сохраняем результат
            base_name = os.path.splitext(os.path.basename(self.edit_path))[0]
            output_path = os.path.join(_output_dir(), f'{base_name}_ОБНОВЛЕНО.xlsx')
            result_df.to_excel(output_path, index=False)

            added_count = len(new_materials)
            updated_count = len(common_materials)
            # Длинный путь в статус не влезет — показываем только имя файла
            self.status.text = (f'Готово! Добавлено: {added_count}, Обновлено: {updated_count}.\n'
                                f'Файл: {os.path.basename(output_path)}')

        except Exception as e:
            self.status.text = f'Ошибка: {e}'


if __name__ == '__main__':
    ExcelMergerApp().run()
