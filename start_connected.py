"""Launch the original UOL dashboard against its bundled live Excel workbook files."""
from pathlib import Path
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from functools import partial
import socket
import threading
import webbrowser
import sys

ROOT = Path(__file__).resolve().parent
SOURCES = [
    ROOT / 'data' / 'UOL_Separated_All_Activity_Data_Cleaned.xlsx',
    ROOT / 'data' / '65_Submissions_Source.xlsx',
    ROOT / 'data' / 'UOL_Internship_Placement_IP_Combined (2)(1).xlsx',
    ROOT / 'data' / 'UOL_MoU_MoU_Signing_Combined_With_Country.xlsx',
]
print('UOL Institutional Engagement Dashboard')
for p in SOURCES:
    print(('OK:      ' if p.exists() else 'MISSING: ') + str(p.relative_to(ROOT)))
if not all(p.exists() for p in SOURCES):
    print('Warning: some expected workbook files are missing; their views may be unavailable.')

class Handler(SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Cache-Control', 'no-store, max-age=0')
        super().end_headers()

handler = partial(Handler, directory=str(ROOT))
server = ThreadingHTTPServer(('127.0.0.1', 0), handler)
url = f'http://127.0.0.1:{server.server_port}/index.html'
print('Opening', url)
print('Edit Excel files in the data folder and refresh the browser to load changes.')
print('Press Ctrl+C to stop the local server.')
threading.Timer(0.5, lambda: webbrowser.open(url)).start()
try:
    server.serve_forever()
except KeyboardInterrupt:
    pass
finally:
    server.server_close()
