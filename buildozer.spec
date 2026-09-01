[app]
title = SinKa PvP 100
package.name = sinkapvp100
package.domain = tr.sinka
source.dir = .
source.include_exts = py,png,jpg,jpeg,kv,atlas,ttf
version = 1.0.0
requirements = python3,kivy,pillow,openpyxl
orientation = landscape
fullscreen = 0
android.api = 35
android.minapi = 23
android.ndk = 27c
android.archs = arm64-v8a, armeabi-v7a
android.permissions = INTERNET
android.accept_sdk_license = True

[buildozer]
log_level = 2
warn_on_root = 1

[app:android]
android.private_storage = True
android.add_src =
android.entrypoint = org.kivy.android.PythonActivity
