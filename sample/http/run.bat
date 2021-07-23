setlocal
set PATH=%PATH%;%~dp0build\Release
set PYTHONPATH=%~dp0build\Release
set QT_PLUGIN_PATH=%~dp0..\..\vcpkg_installed\x64-windows\plugins
%~dp0..\..\vcpkg_installed\x64-windows\tools\python3\python.exe test.py
