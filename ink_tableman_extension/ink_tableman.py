#!/usr/bin/env python3
"""
Inkscape extension Boilerplate with Web UI using GTK and WebKit2.
"""

import inkex
from inkex import TextElement, Rectangle, Ellipse, Tspan, Layer
import os
import sys
import json
import threading
import http.server
import socketserver
import webbrowser
import socket

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

# Redirect stderr to a log file to avoid Inkscape's built-in error dialog popup 
# whenever warnings or non-fatal errors are emitted by GTK/WebKit
log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'extension_debug.log')
log_file = open(log_path, 'w')
os.dup2(log_file.fileno(), sys.stderr.fileno())

class WebUIExtension(inkex.EffectExtension):
    """Base class for Inkscape extension with Web UI."""

    class WebUIHandler(http.server.SimpleHTTPRequestHandler):
        """Handler for the Web UI HTTP server."""
        
        def __init__(self, *args, extension_instance=None, **kwargs):
            self.extension_instance = extension_instance
            super().__init__(*args, **kwargs)

        def log_message(self, format, *args):
            # Silence logging
            pass

        def do_GET(self):
            if self.path == '/status':
                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps(self.extension_instance.status_data).encode('utf-8'))
            else:
                super().do_GET()

        def do_POST(self):
            if self.path == '/submit':
                content_length = int(self.headers['Content-Length'])
                post_data = self.rfile.read(content_length)
                data = json.loads(post_data.decode('utf-8'))

                # Process the data asynchronously
                threading.Thread(target=self.extension_instance.process_background, args=(data,)).start()
                
                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'status': 'started'}).encode('utf-8'))
                
            elif self.path == '/close':
                # Shut down safely
                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'status': 'closing'}).encode('utf-8'))
                
                if GTK_UI_AVAILABLE:
                    from gi.repository import GLib, Gtk
                    # Idle add quit so response finishes sending first
                    GLib.idle_add(Gtk.main_quit)

    def __init__(self):
        super().__init__()
        self.status_data = {"status": "idle", "progress": 0, "message": ""}
        self.is_processing = False
        
    def add_arguments(self, pars):
        # Additional parameters if required
        pass

    def run_web_ui(self):
        """Start a local server and open the Web UI in a native GTK window."""
        # Find a free port
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(('', 0))
            port = s.getsockname()[1]
        
        ui_dir = os.path.join(os.path.dirname(__file__), 'ui')
        
        # Create the server
        handler_class = lambda *args, **kwargs: self.WebUIHandler(
            *args, extension_instance=self, directory=ui_dir, **kwargs
        )
        
        server = socketserver.TCPServer(("", port), handler_class)
        # Start server in a thread so GTK can run its main loop
        server_thread = threading.Thread(target=server.serve_forever)
        server_thread.daemon = True
        server_thread.start()
        
        url = f"http://localhost:{port}/index.html"
        
        if GTK_UI_AVAILABLE:
            # Add a heartbeat to keep the GTK main loop "alive" during background tasks
            heartbeat_id = GLib.timeout_add(100, lambda: True)
            
            # Run GTK native window
            GLib.idle_add(self._launch_gtk_window, url, server)
            Gtk.main()
            
            # Cleanup heartbeat
            GLib.source_remove(heartbeat_id)
        else:
            # Fallback to browser
            webbrowser.open(url)
            # For fallback, we need to wait until the task completes
            while self.is_processing:
                import time
                time.sleep(0.1)
        
        server.server_close()

    def _launch_gtk_window(self, url, server):
        window = Gtk.Window(title="Basic Shape Inserter")
        # Set dimensions here
        window.set_default_size(500, 750)
        window.set_position(Gtk.WindowPosition.CENTER)
        window.set_resizable(False) # Prevent resizing and maximizing
        
        webview = WebKit2.WebView()
        window.add(webview)
        webview.load_uri(url)
        
        def on_destroy(widget):
            server.shutdown()
            Gtk.main_quit()
            
        window.connect("destroy", on_destroy)
        window.show_all()

    def process_background(self, data):
        """Asynchronous generation process called in a background thread."""
        try:
            self.is_processing = True
            
            self.status_data = {"status": "processing", "progress": 10, "message": "Parsing options..."}
            
            import time
            time.sleep(0.3)
            
            # Extract common parameters
            shape_type = data.get('shape_type', 'rect')
            fill_color = data.get('fill_color', '#cccccc')
            stroke_color = data.get('stroke_color', '#000000')
            stroke_width = data.get('stroke_width', '1')
            
            self.status_data.update({"progress": 40, "message": f"Creating {shape_type} element..."})
            
            # Style dictionary
            style = {
                'fill': fill_color,
                'stroke': stroke_color,
                'stroke-width': str(stroke_width)
            }
            
            layer = self.svg.get_current_layer()
            
            # Get center of view for placement
            view_center = self.svg.namedview.center
            cx = view_center[0] if view_center else 100
            cy = view_center[1] if view_center else 100
            
            if shape_type == 'text':
                text_content = data.get('text_content', '')
                if not text_content: text_content = "Type something..."
                font_size = data.get('font_size', '24')
                
                style['font-size'] = f"{font_size}px"
                style['font-family'] = "sans-serif"
                
                # Create text element
                text_elem = layer.add(TextElement())
                text_elem.set('x', str(cx))
                text_elem.set('y', str(cy))
                text_elem.style = style
                
                # Add tspan with content
                tspan = Tspan()
                tspan.set('x', str(cx))
                tspan.set('y', str(cy))
                tspan.text = text_content
                text_elem.append(tspan)
                
            elif shape_type == 'rect':
                width = float(data.get('width', 100))
                height = float(data.get('height', 100))
                
                rect = layer.add(Rectangle())
                rect.set('x', str(cx - width/2))
                rect.set('y', str(cy - height/2))
                rect.set('width', str(width))
                rect.set('height', str(height))
                rect.style = style
                
            elif shape_type == 'ellipse':
                width = float(data.get('width', 100))
                height = float(data.get('height', 100))
                
                ellipse = layer.add(Ellipse())
                ellipse.set('cx', str(cx))
                ellipse.set('cy', str(cy))
                ellipse.set('rx', str(width/2))
                ellipse.set('ry', str(height/2))
                ellipse.style = style

            self.status_data.update({"progress": 80, "message": "Appending to Inkscape..."})
            time.sleep(0.3)
                
            self.status_data = {"status": "completed", "progress": 100, "message": "Successfully generated content!"}
            
            # Auto-close after completion
            if GTK_UI_AVAILABLE:
                time.sleep(0.8) # Let the UI show 100% completion before closing
                from gi.repository import GLib, Gtk
                GLib.idle_add(Gtk.main_quit)
                
        except Exception as e:
            self.status_data = {"status": "error", "progress": 0, "message": f"Error: {str(e)}"}
        finally:
            self.is_processing = False

    def effect(self):
        """Main effect function."""
        # Run Web UI immediately, block until GTK loop finishes
        self.run_web_ui()

if __name__ == '__main__':
    WebUIExtension().run()
