[app]

# Название приложения (видно на телефоне)
title = Excel Merger

# Имя пакета: org.test.excelmerger -> org.test.excelmerger-0.1-armeabi-v7a-debug.apk
package.name = excelmerger
package.domain = org.test

source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json,ttf

# Версия
version = 0.6

# Стек по умолчанию (Python 3.14 от p4a). pandas v3.1.1 — первая ветка с официальной
# поддержкой Python 3.14; numpy v2.3.4 — минимально требуемая для 3.14 и с фиксом
# сборки unique.cpp. Теги указаны с "v", т.к. p4a делает `git checkout <версия>`.
requirements = python3,kivy==2.3.1,numpy==v2.3.4,pandas==v3.1.1,openpyxl,xlrd,android

# Локальные рецепты (p4a.local_recipes):
#   pandas — штатный рецепт p4a тянет v2.3.0, которая под Python 3.14 не компилируется,
#            а его патч не применяется к 3.x;
#   kivy   — штатный рецепт тянет requests -> charset-normalizer, у которого нет
#            универсального wheel (сборка падает на "not a supported wheel").
p4a.local_recipes = ./p4a-recipes

orientation = portrait
# Полноэкранный режим: без него SDL-окно может оказаться меньше экрана
fullscreen = 1

# compileSdk 35: Java-код SDL2 в p4a использует API 31+ (Manifest.permission.BLUETOOTH_CONNECT).
# minapi остаётся 24 — старые телефоны поддерживаются.
android.api = 35
android.minapi = 24
android.ndk = 25b
# Только arm64 (все телефоны с 2018 г.). Для старых моделей добавьте: arm64-v8a, armeabi-v7a
android.archs = arm64-v8a
android.accept_sdk_license = True

# Доступ к файлам на телефоне
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE,MANAGE_EXTERNAL_STORAGE

# Разрешаем старый режим доступа к памяти (нужно для записи в /storage/emulated/0/Download)
# ВАЖНО: buildozer читает значение как ПУТЬ К ФАЙЛУ с атрибутами <application>, а не как строку.
android.extra_manifest_application_arguments = android-manifest-app-args.txt

android.allow_backup = True

[buildozer]

log_level = 2
warn_on_root = 1
