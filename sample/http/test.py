import requests
import http_py

r = requests.get('https://www.qt.io')
with open('index1.html', 'wb') as f:
	f.write(bytes(r.text, r.encoding))

http_py.OpenDownloadApp()

qt_lines = []
with open('index.html', 'r') as f:
	qt_lines = f.readlines()

py_lines = []
with open('index1.html', 'r') as f:
	py_lines = f.readlines()

n = min(len(qt_lines), len(py_lines))
for i in range(n):
	if qt_lines[i] != py_lines[i]:
		print(f'Line {i} differs\n\tpy: {py_lines[i]}\tqt: {qt_lines[i]}')