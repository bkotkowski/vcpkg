%~dp0..\..\downloads\tools\cmake-3.20.2-windows\cmake-3.20.2-windows-i386\bin\cmake.exe -B build -S . -G "Visual Studio 15 2017" -A x64 -DCMAKE_TOOLCHAIN_FILE=%~dp0..\..\scripts\buildsystems\vcpkg.cmake -DVCPKG_TARGET_TRIPLET=x64-windows -DVCPKG_INSTALLED_DIR=%~dp0..\..\vcpkg_installed%
"%DevEnvDir%devenv.exe" build\http.sln /build Debug

