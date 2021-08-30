setlocal
set PATH=%PATH%;%~dp0build\Debug
set PYTHONPATH=%~dp0build\Debug
set QT_PLUGIN_PATH=%~dp0..\..\vcpkg_installed\x64-windows\plugins
%~dp0..\..\vcpkg_installed\x64-windows\tools\python3\python.exe test.py
