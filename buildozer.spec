[app]

title = Paper Formatter
package.name = paperformatter
package.domain = org.example

source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json

version = 0.1

requirements = python3,kivy,python-docx,markdown,plyer,pillow

orientation = portrait

osx.python_version = 3
osx.kivy_version = 2.1.0

fullscreen = 0

android.permissions = STORAGE
android.archs = arm64-v8a,armeabi-v7a
android.accept_sdk_license = True
