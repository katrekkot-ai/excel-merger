[app]

# Название приложения (видно на телефоне)
title = Excel Merger

# Имя пакета: org.test.excelmerger -> org.test.excelmerger-0.1-armeabi-v7a-debug.apk
package.name = excelmerger
package.domain = org.test

source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json,ttf

# Версия
version = 0.1

# Зависимости. pandas тянет за собой numpy, openpyxl читает .xlsx, xlrd — .xls
# numpy пиним на v2.3.4 (именно с "v" — p4a делает git checkout тега):
# в 2.3.0-2.3.2 нет #include <unordered_map> в unique.cpp -> сборка под Android падает
# Python 3.12: p4a по умолчанию берёт 3.14, под который pandas 2.3.x ещё не собирается.
# pandas v2.3.3 и numpy v2.3.4 — версии, проверенные рецептами p4a (теги с "v" для git checkout).
requirements = hostpython3==3.12.10,python3==3.12.10,kivy==2.3.1,numpy==v2.3.4,pandas==v2.3.3,openpyxl,xlrd

orientation = portrait
fullscreen = 0

android.api = 30
android.minapi = 24
android.ndk = 25b
# Только arm64 (все телефоны с 2018 г.). Для старых моделей добавьте: arm64-v8a, armeabi-v7a
android.archs = arm64-v8a
android.accept_sdk_license = True

# Доступ к файлам на телефоне
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE,MANAGE_EXTERNAL_STORAGE

# Разрешаем старый режим доступа к памяти (нужно для записи в /storage/emulated/0/Download)
android.extra_manifest_application_arguments = android:requestLegacyExternalStorage="true"

android.allow_backup = True

[buildozer]

log_level = 2
warn_on_root = 1
