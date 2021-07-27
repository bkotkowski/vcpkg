#!/bin/sh
for f in `find vcpkg_installed/x64-linux-dynamic/lib/*.so* -type f -print`; do
    patchelf --set-rpath '$ORIGIN' $f
done

for f in `find vcpkg_installed/x64-linux-dynamic/lib/python*/lib-dynload/*.so* -type f -print`; do
    patchelf --set-rpath '$ORIGIN:$ORIGIN/../..' $f
done