"""Локальный рецепт kivy для python-for-android.

Штатный рецепт kivy объявляет:

    python_depends = ['certifi', 'chardet', 'idna', 'requests', 'urllib3', 'filetype']

`requests` тянет за собой `charset-normalizer`, и p4a ломается на нём:
для pure-python модулей версия отбрасывается, ставится charset-normalizer 3.5.2,
у которого нет универсального wheel, а p4a подсовывает переименованный под android
wheel -> "charset_normalizer-...-android_24_arm64_v8a.whl is not a supported wheel".

Этому приложению сеть не нужна, поэтому requests (и его зависимости) просто убираем.
Подключается через `p4a.local_recipes = ./p4a-recipes` в buildozer.spec.
"""

from pythonforandroid.recipes.kivy import KivyRecipe as _KivyRecipe


class KivyRecipe(_KivyRecipe):
    python_depends = ['certifi', 'chardet', 'idna', 'urllib3', 'filetype']


recipe = KivyRecipe()
