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

# Стек по умолчанию (Python 3.14 от p4a). pandas v3.1.1 — первая ветка с официальной
# поддержкой Python 3.14; numpy v2.3.4 — минимально требуемая для 3.14 и с фиксом
# сборки unique.cpp. Теги указаны с "v", т.к. p4a делает `git checkout <версия>`.
requirements = python3,kivy==2.3.1,numpy==v2.3.4,pandas==v3.1.1,openpyxl,xlrd,charset-normalizer==3.3.2

# Локальный рецепт pandas (p4a.local_recipes ниже): штатный рецепт p4a тянет pandas
# v2.3.0, которая под Python 3.14 не компилируется, а его патч не применяется к 3.x.
p4a.local_recipes = ./p4a-recipes

# charset-normalizer 3.5.x поставляется скомпилированным wheel под чужие платформы,
# p4a не может его поставить. 3.3.2 — чистый python.
# (приезжает транзитивно через requests)

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
