#include <Python.h>
#include <pybind11/pybind11.h>

int openDownloadApp();

namespace py = pybind11;

PYBIND11_MODULE(http_py, m) {
    m.doc() = "pybind11 wrapper for openssl test";
    m.def("OpenDownloadApp", &openDownloadApp, "Open Download App");
}