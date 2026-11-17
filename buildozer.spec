[app]

# РќР°Р·РІР°РЅРёРµ РїСЂРёР»РѕР¶РµРЅРёСЏ (РІРёРґРЅРѕ РЅР° С‚РµР»РµС„РѕРЅРµ)
title = Excel Merger

# РРјСЏ РїР°РєРµС‚Р°: org.test.excelmerger -> org.test.excelmerger-0.1-armeabi-v7a-debug.apk
package.name = excelmerger
package.domain = org.test

source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json,ttf

# Р’РµСЂСЃРёСЏ
version = 0.10
# versionCode для Android: должен расти при каждой сборке, иначе телефон
# может отказаться устанавливать APK поверх предыдущего.
android.numeric_version = 10

# РЎС‚РµРє РїРѕ СѓРјРѕР»С‡Р°РЅРёСЋ (Python 3.14 РѕС‚ p4a). pandas v3.1.1 вЂ” РїРµСЂРІР°СЏ РІРµС‚РєР° СЃ РѕС„РёС†РёР°Р»СЊРЅРѕР№
# РїРѕРґРґРµСЂР¶РєРѕР№ Python 3.14; numpy v2.3.4 вЂ” РјРёРЅРёРјР°Р»СЊРЅРѕ С‚СЂРµР±СѓРµРјР°СЏ РґР»СЏ 3.14 Рё СЃ С„РёРєСЃРѕРј
# СЃР±РѕСЂРєРё unique.cpp. РўРµРіРё СѓРєР°Р·Р°РЅС‹ СЃ "v", С‚.Рє. p4a РґРµР»Р°РµС‚ `git checkout <РІРµСЂСЃРёСЏ>`.
requirements = python3,kivy==2.3.1,numpy==v2.3.4,pandas==v3.1.1,openpyxl,xlrd,android

# Р›РѕРєР°Р»СЊРЅС‹Рµ СЂРµС†РµРїС‚С‹ (p4a.local_recipes):
#   pandas вЂ” С€С‚Р°С‚РЅС‹Р№ СЂРµС†РµРїС‚ p4a С‚СЏРЅРµС‚ v2.3.0, РєРѕС‚РѕСЂР°СЏ РїРѕРґ Python 3.14 РЅРµ РєРѕРјРїРёР»РёСЂСѓРµС‚СЃСЏ,
#            Р° РµРіРѕ РїР°С‚С‡ РЅРµ РїСЂРёРјРµРЅСЏРµС‚СЃСЏ Рє 3.x;
#   kivy   вЂ” С€С‚Р°С‚РЅС‹Р№ СЂРµС†РµРїС‚ С‚СЏРЅРµС‚ requests -> charset-normalizer, Сѓ РєРѕС‚РѕСЂРѕРіРѕ РЅРµС‚
#            СѓРЅРёРІРµСЂСЃР°Р»СЊРЅРѕРіРѕ wheel (СЃР±РѕСЂРєР° РїР°РґР°РµС‚ РЅР° "not a supported wheel").
p4a.local_recipes = ./p4a-recipes

orientation = portrait
# РџРѕР»РЅРѕСЌРєСЂР°РЅРЅС‹Р№ СЂРµР¶РёРј: Р±РµР· РЅРµРіРѕ SDL-РѕРєРЅРѕ РјРѕР¶РµС‚ РѕРєР°Р·Р°С‚СЊСЃСЏ РјРµРЅСЊС€Рµ СЌРєСЂР°РЅР°
fullscreen = 1

# compileSdk 35: Java-РєРѕРґ SDL2 РІ p4a РёСЃРїРѕР»СЊР·СѓРµС‚ API 31+ (Manifest.permission.BLUETOOTH_CONNECT).
# minapi РѕСЃС‚Р°С‘С‚СЃСЏ 24 вЂ” СЃС‚Р°СЂС‹Рµ С‚РµР»РµС„РѕРЅС‹ РїРѕРґРґРµСЂР¶РёРІР°СЋС‚СЃСЏ.
android.api = 35
android.minapi = 24
android.ndk = 25b
# РўРѕР»СЊРєРѕ arm64 (РІСЃРµ С‚РµР»РµС„РѕРЅС‹ СЃ 2018 Рі.). Р”Р»СЏ СЃС‚Р°СЂС‹С… РјРѕРґРµР»РµР№ РґРѕР±Р°РІСЊС‚Рµ: arm64-v8a, armeabi-v7a
android.archs = arm64-v8a
android.accept_sdk_license = True

# Р”РѕСЃС‚СѓРї Рє С„Р°Р№Р»Р°Рј РЅР° С‚РµР»РµС„РѕРЅРµ
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE,MANAGE_EXTERNAL_STORAGE

# Р Р°Р·СЂРµС€Р°РµРј СЃС‚Р°СЂС‹Р№ СЂРµР¶РёРј РґРѕСЃС‚СѓРїР° Рє РїР°РјСЏС‚Рё (РЅСѓР¶РЅРѕ РґР»СЏ Р·Р°РїРёСЃРё РІ /storage/emulated/0/Download)
# Р’РђР–РќРћ: buildozer С‡РёС‚Р°РµС‚ Р·РЅР°С‡РµРЅРёРµ РєР°Рє РџРЈРўР¬ Рљ Р¤РђР™Р›РЈ СЃ Р°С‚СЂРёР±СѓС‚Р°РјРё <application>, Р° РЅРµ РєР°Рє СЃС‚СЂРѕРєСѓ.
android.extra_manifest_application_arguments = android-manifest-app-args.txt

android.allow_backup = True

[buildozer]

log_level = 2
warn_on_root = 1
