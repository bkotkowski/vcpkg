#!/bin/sh
/vcpkg/vcpkg/downloads/tools/cmake-3.20.2-linux/cmake-3.20.2-linux-x86_64/bin/cmake -B build -S . -DCMAKE_TOOLCHAIN_FILE=/vcpkg/vcpkg/scripts/buildsystems/vcpkg.cmake -DVCPKG_TARGET_TRIPLET=x64-linux-dynamic -DVCPKG_OVERLAY_TRIPLETS=custom-triplets
/vcpkg/vcpkg/downloads/tools/cmake-3.20.2-linux/cmake-3.20.2-linux-x86_64/bin/cmake --build build -DCMAKE_BUILD_TYPE=Release
