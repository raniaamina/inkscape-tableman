#!/usr/bin/env python3
"""
Inkscape Table Manager Extension — CRUD table management with Web UI.
Uses GTK + WebKit2 for native window rendering.
"""

import inkex
from inkex import TextElement, Rectangle, Tspan, Group
from lxml import etree
import os
import sys
import json
import threading
import http.server
import socketserver
import webbrowser
import socket
import uuid
import subprocess
import urllib.parse

# Try to load GTK and WebKit2
try:
    import gi
    gi.require_version('Gtk', '3.0')
    gi.require_version('WebKit2', '4.1')
    from gi.repository import Gtk, WebKit2, GLib
    GTK_UI_AVAILABLE = True
except (ImportError, ValueError):
    try:
        import gi
        gi.require_version('Gtk', '3.0')
        gi.require_version('WebKit2', '4.0')
        from gi.repository import Gtk, WebKit2, GLib
        GTK_UI_AVAILABLE = True
    except:
        GTK_UI_AVAILABLE = False

# Redirect stderr to log file
log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'extension_debug.log')
log_file = open(log_path, 'w')
os.dup2(log_file.fileno(), sys.stderr.fileno())

# Namespace for custom attributes
TABLEMAN_NS = 'http://raniaamina.id/tableman'
TABLEMAN_PREFIX = 'tableman'
INKSCAPE_LABEL_NS = 'http://www.inkscape.org/namespaces/inkscape'
SVG_NS = 'http://www.w3.org/2000/svg'


def get_system_fonts():
    """Get list of system fonts using fc-list."""
    fallback = ['sans-serif', 'serif', 'monospace', 'Arial', 'Helvetica',
                'Times New Roman', 'Courier New', 'Georgia', 'Verdana']
    try:
        result = subprocess.run(
            ['fc-list', ':lang=en', 'family'],
            capture_output=True, text=True, timeout=5
        )
        if result.returncode != 0:
            # Try without lang filter
            result = subprocess.run(
                ['fc-list', '', 'family'],
                capture_output=True, text=True, timeout=5
            )
        if result.returncode == 0 and result.stdout.strip():
            fonts = set()
            for line in result.stdout.strip().split('\n'):
                # fc-list outputs "Family Name" or "Family1,Family2"
                for name in line.split(','):
                    name = name.strip()
                    if name:
                        fonts.add(name)
            if fonts:
                return sorted(fonts)
    except Exception:
        pass
    return fallback


class TablemanExtension(inkex.EffectExtension):
    """Inkscape Table Manager with full CRUD via Web UI."""

    TABLEMAN_ID_PREFIX = 'tableman-'

    class WebUIHandler(http.server.SimpleHTTPRequestHandler):
        """HTTP handler with REST API for table CRUD operations."""

        def __init__(self, *args, extension_instance=None, **kwargs):
            self.extension_instance = extension_instance
            super().__init__(*args, **kwargs)

        def log_message(self, format, *args):
            pass

        def _send_json(self, data, status=200):
            self.send_response(status)
            self.send_header('Content-type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps(data).encode('utf-8'))

        def _read_json(self):
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            return json.loads(post_data.decode('utf-8'))

        def do_GET(self):
            if self.path == '/status':
                self._send_json(self.extension_instance.status_data)
            elif self.path == '/tables':
                tables = self.extension_instance.scan_tables()
                self._send_json({'tables': tables})
            elif self.path.startswith('/table/'):
                table_id = urllib.parse.unquote(self.path[7:])
                table_data = self.extension_instance.load_table(table_id)
                if table_data:
                    self._send_json(table_data)
                else:
                    self._send_json({'error': 'Table not found'}, 404)
            elif self.path == '/fonts':
                fonts = get_system_fonts()
                self._send_json({'fonts': fonts})
            else:
                super().do_GET()

        def do_POST(self):
            if self.path == '/submit':
                data = self._read_json()
                threading.Thread(
                    target=self.extension_instance.process_submit, args=(data,)
                ).start()
                self._send_json({'status': 'started'})

            elif self.path == '/delete':
                data = self._read_json()
                result = self.extension_instance.delete_table(data.get('id'))
                self._send_json(result)

            elif self.path == '/close':
                self._send_json({'status': 'closing'})
                if GTK_UI_AVAILABLE:
                    GLib.idle_add(Gtk.main_quit)

    def __init__(self):
        super().__init__()
        self.status_data = {"status": "idle", "progress": 0, "message": ""}
        self.is_processing = False

    def add_arguments(self, pars):
        pass

    # ─── SVG SCANNING ─────────────────────────────────────────

    def scan_tables(self):
        tables = []
        for elem in self.svg.iter():
            elem_id = elem.get('id', '')
            if elem_id.startswith(self.TABLEMAN_ID_PREFIX):
                label = elem.get(f'{{{INKSCAPE_LABEL_NS}}}label', elem_id)
                rows = elem.get(f'{{{TABLEMAN_NS}}}rows', '0')
                cols = elem.get(f'{{{TABLEMAN_NS}}}cols', '0')
                tables.append({
                    'id': elem_id,
                    'label': label,
                    'rows': int(rows),
                    'cols': int(cols)
                })
        return tables

    def load_table(self, table_id):
        elem = self.svg.getElementById(table_id)
        if elem is None:
            return None

        try:
            data = json.loads(elem.get(f'{{{TABLEMAN_NS}}}data', '[]'))
        except:
            data = []
        try:
            merges = json.loads(elem.get(f'{{{TABLEMAN_NS}}}merges', '[]'))
        except:
            merges = []
        try:
            cell_styles = json.loads(elem.get(f'{{{TABLEMAN_NS}}}cell-styles', '[]'))
        except:
            cell_styles = []

        return {
            'id': table_id,
            'label': elem.get(f'{{{INKSCAPE_LABEL_NS}}}label', table_id),
            'rows': int(elem.get(f'{{{TABLEMAN_NS}}}rows', '0')),
            'cols': int(elem.get(f'{{{TABLEMAN_NS}}}cols', '0')),
            'cell_width': float(elem.get(f'{{{TABLEMAN_NS}}}cell-width', '120')),
            'cell_height': float(elem.get(f'{{{TABLEMAN_NS}}}cell-height', '40')),
            'border_width': float(elem.get(f'{{{TABLEMAN_NS}}}border-width', '1')),
            'border_color': elem.get(f'{{{TABLEMAN_NS}}}border-color', '#333333'),
            'header_fill': elem.get(f'{{{TABLEMAN_NS}}}header-fill', '#bb86fc'),
            'body_fill': elem.get(f'{{{TABLEMAN_NS}}}body-fill', '#1e1e1e'),
            'font_family': elem.get(f'{{{TABLEMAN_NS}}}font-family', 'sans-serif'),
            'font_size': float(elem.get(f'{{{TABLEMAN_NS}}}font-size', '14')),
            'text_color': elem.get(f'{{{TABLEMAN_NS}}}text-color', '#ffffff'),
            'header_text_color': elem.get(f'{{{TABLEMAN_NS}}}header-text-color', '#000000'),
            'data': data,
            'merges': merges,
            'cell_styles': cell_styles
        }

    def delete_table(self, table_id):
        if not table_id:
            return {'status': 'error', 'message': 'No ID provided'}
        elem = self.svg.getElementById(table_id)
        if elem is not None:
            elem.getparent().remove(elem)
            return {'status': 'deleted'}
        return {'status': 'error', 'message': 'Table not found'}

    # ─── TABLE RENDERING ──────────────────────────────────────

    def _get_merge_at(self, merges, r, c):
        for m in merges:
            if m['row'] == r and m['col'] == c:
                return (m.get('rowspan', 1), m.get('colspan', 1))
        return None

    def _is_hidden_by_merge(self, merges, r, c):
        for m in merges:
            mr, mc = m['row'], m['col']
            rs, cs = m.get('rowspan', 1), m.get('colspan', 1)
            if mr <= r < mr + rs and mc <= c < mc + cs:
                if mr == r and mc == c:
                    return False
                return True
        return False

    def render_table(self, group, data):
        rows = int(data.get('rows', 3))
        cols = int(data.get('cols', 3))
        cw = float(data.get('cell_width', 120))
        ch = float(data.get('cell_height', 40))
        border_w = float(data.get('border_width', 1))
        border_c = data.get('border_color', '#333333')
        header_fill = data.get('header_fill', '#bb86fc')
        body_fill = data.get('body_fill', '#1e1e1e')
        font_family = data.get('font_family', 'sans-serif')
        font_size = float(data.get('font_size', 14))
        text_color = data.get('text_color', '#ffffff')
        header_text_color = data.get('header_text_color', '#000000')
        cell_data = data.get('data', [])
        merges = data.get('merges', [])
        cell_styles = data.get('cell_styles', [])

        for r in range(rows):
            for c in range(cols):
                if self._is_hidden_by_merge(merges, r, c):
                    continue

                x = c * cw
                y = r * ch
                is_header = (r == 0)

                merge_info = self._get_merge_at(merges, r, c)
                if merge_info:
                    rs, cs = merge_info
                    cell_w = cs * cw
                    cell_h = rs * ch
                else:
                    cell_w = cw
                    cell_h = ch

                # Cell rectangle
                rect = Rectangle()
                rect.set('x', str(x))
                rect.set('y', str(y))
                rect.set('width', str(cell_w))
                rect.set('height', str(cell_h))
                rect.style = {
                    'fill': header_fill if is_header else body_fill,
                    'stroke': border_c,
                    'stroke-width': str(border_w)
                }
                group.append(rect)

                # Cell text
                cell_text = ''
                if r < len(cell_data) and c < len(cell_data[r]):
                    cell_text = str(cell_data[r][c])

                if cell_text:
                    # Get per-cell style
                    cs_obj = {}
                    if r < len(cell_styles) and c < len(cell_styles[r]):
                        cs_obj = cell_styles[r][c] if isinstance(cell_styles[r][c], dict) else {}

                    text_style = {
                        'font-size': f'{font_size}px',
                        'font-family': font_family,
                        'fill': header_text_color if is_header else text_color,
                        'text-anchor': 'middle',
                    }

                    # Per-cell formatting
                    if cs_obj.get('bold'):
                        text_style['font-weight'] = 'bold'
                    if cs_obj.get('italic'):
                        text_style['font-style'] = 'italic'

                    decoration = []
                    if cs_obj.get('underline'):
                        decoration.append('underline')
                    if cs_obj.get('strikethrough'):
                        decoration.append('line-through')
                    if decoration:
                        text_style['text-decoration'] = ' '.join(decoration)

                    text_elem = TextElement()
                    text_elem.set('x', str(x + cell_w / 2))
                    text_elem.set('y', str(y + cell_h / 2 + font_size / 3))
                    text_elem.style = text_style
                    tspan = Tspan()
                    tspan.text = cell_text
                    text_elem.append(tspan)
                    group.append(text_elem)

    # ─── SUBMIT HANDLER ───────────────────────────────────────

    def process_submit(self, data):
        try:
            self.is_processing = True
            self.status_data = {"status": "processing", "progress": 10, "message": "Parsing table data..."}

            import time
            time.sleep(0.2)

            table_id = data.get('id', '')
            label = data.get('label', 'Untitled Table')
            is_new = not table_id or not self.svg.getElementById(table_id)

            if is_new:
                table_id = self.TABLEMAN_ID_PREFIX + str(uuid.uuid4())[:8]

            self.status_data.update({"progress": 30, "message": "Preparing SVG group..."})

            etree.register_namespace(TABLEMAN_PREFIX, TABLEMAN_NS)

            if is_new:
                group = Group()
                group.set('id', table_id)
                layer = self.svg.get_current_layer()

                view_center = self.svg.namedview.center
                cx = view_center[0] if view_center else 100
                cy = view_center[1] if view_center else 100

                r = int(data.get('rows', 3))
                c = int(data.get('cols', 3))
                cw = float(data.get('cell_width', 120))
                ch = float(data.get('cell_height', 40))
                group.set('transform', f'translate({cx - (c*cw)/2},{cy - (r*ch)/2})')
                layer.append(group)
            else:
                group = self.svg.getElementById(table_id)
                for child in list(group):
                    group.remove(child)

            self.status_data.update({"progress": 50, "message": "Storing metadata..."})

            group.set(f'{{{INKSCAPE_LABEL_NS}}}label', f'Tableman: {label}')
            for key in ['rows','cols','cell_width','cell_height','border_width',
                        'border_color','header_fill','body_fill','font_family',
                        'font_size','text_color','header_text_color']:
                attr = key.replace('_', '-')
                group.set(f'{{{TABLEMAN_NS}}}{attr}', str(data.get(key, '')))

            group.set(f'{{{TABLEMAN_NS}}}data', json.dumps(data.get('data', [])))
            group.set(f'{{{TABLEMAN_NS}}}merges', json.dumps(data.get('merges', [])))
            group.set(f'{{{TABLEMAN_NS}}}cell-styles', json.dumps(data.get('cell_styles', [])))

            self.status_data.update({"progress": 70, "message": "Rendering table..."})
            time.sleep(0.2)

            self.render_table(group, data)

            self.status_data.update({"progress": 90, "message": "Finalizing..."})
            time.sleep(0.2)

            action = "Created" if is_new else "Updated"
            self.status_data = {
                "status": "completed", "progress": 100,
                "message": f"{action} table '{label}' successfully!"
            }

            if GTK_UI_AVAILABLE:
                time.sleep(0.8)
                GLib.idle_add(Gtk.main_quit)

        except Exception as e:
            self.status_data = {"status": "error", "progress": 0, "message": f"Error: {str(e)}"}
        finally:
            self.is_processing = False

    # ─── WEB UI LAUNCH ────────────────────────────────────────

    def run_web_ui(self):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(('', 0))
            port = s.getsockname()[1]

        ui_dir = os.path.join(os.path.dirname(__file__), 'ui')
        handler_class = lambda *args, **kwargs: self.WebUIHandler(
            *args, extension_instance=self, directory=ui_dir, **kwargs
        )

        server = socketserver.TCPServer(("", port), handler_class)
        server_thread = threading.Thread(target=server.serve_forever)
        server_thread.daemon = True
        server_thread.start()

        url = f"http://localhost:{port}/index.html"

        if GTK_UI_AVAILABLE:
            heartbeat_id = GLib.timeout_add(100, lambda: True)
            GLib.idle_add(self._launch_gtk_window, url, server)
            Gtk.main()
            GLib.source_remove(heartbeat_id)
        else:
            webbrowser.open(url)
            while self.is_processing:
                import time
                time.sleep(0.1)

        server.server_close()

    def _launch_gtk_window(self, url, server):
        window = Gtk.Window(title="Tableman — Table Manager")
        window.set_default_size(640, 780)
        window.set_position(Gtk.WindowPosition.CENTER)
        window.set_resizable(False)

        webview = WebKit2.WebView()
        window.add(webview)
        webview.load_uri(url)

        def on_destroy(widget):
            server.shutdown()
            Gtk.main_quit()

        window.connect("destroy", on_destroy)
        window.show_all()

    def effect(self):
        self.run_web_ui()


if __name__ == '__main__':
    TablemanExtension().run()
