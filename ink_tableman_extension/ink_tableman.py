#!/usr/bin/env python3
"""
Inkscape Table Manager Extension — CRUD table management with Web UI.
"""

import inkex
from inkex import TextElement, Rectangle, Tspan, Group
from lxml import etree
import os, sys, json, threading, http.server, socketserver
import webbrowser, socket, uuid, subprocess, urllib.parse

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

log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'extension_debug.log')
log_file = open(log_path, 'w')
os.dup2(log_file.fileno(), sys.stderr.fileno())

TABLEMAN_NS = 'http://raniaamina.id/tableman'
TABLEMAN_PREFIX = 'tableman'
INKSCAPE_LABEL_NS = 'http://www.inkscape.org/namespaces/inkscape'

SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'settings.json')

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f:
                return json.load(f)
        except: pass
    return {"theme": "system"}

def save_settings(s):
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(s, f)
    except: pass


def get_system_fonts():
    fallback = ['sans-serif','serif','monospace','Arial','Helvetica',
                'Times New Roman','Courier New','Georgia','Verdana',
                'Noto Sans','DejaVu Sans','Liberation Sans']
    try:
        result = subprocess.run(['fc-list','','family'],
            capture_output=True, text=True, timeout=5)
        if result.returncode == 0 and result.stdout.strip():
            fonts = set()
            for line in result.stdout.strip().split('\n'):
                for name in line.split(','):
                    name = name.strip()
                    if name: fonts.add(name)
            if fonts: return sorted(fonts)
    except: pass
    return fallback


class TablemanExtension(inkex.EffectExtension):

    TABLEMAN_ID_PREFIX = 'tableman-'

    class WebUIHandler(http.server.SimpleHTTPRequestHandler):

        def __init__(self, *args, extension_instance=None, **kwargs):
            self.ext = extension_instance
            super().__init__(*args, **kwargs)

        def log_message(self, f, *a): pass

        def _json(self, d, s=200):
            self.send_response(s)
            self.send_header('Content-type','application/json')
            self.send_header('Access-Control-Allow-Origin','*')
            self.end_headers()
            self.wfile.write(json.dumps(d).encode())

        def _body(self):
            return json.loads(self.rfile.read(int(self.headers['Content-Length'])).decode())

        def do_GET(self):
            if self.path=='/status': self._json(self.ext.status_data)
            elif self.path=='/tables': self._json({'tables':self.ext.scan_tables()})
            elif self.path.startswith('/table/'):
                d=self.ext.load_table(urllib.parse.unquote(self.path[7:]))
                self._json(d if d else {'error':'Not found'}, 200 if d else 404)
            elif self.path=='/fonts': self._json({'fonts':get_system_fonts()})
            elif self.path=='/settings': self._json(load_settings())
            else: super().do_GET()

        def do_POST(self):
            if self.path=='/submit':
                d=self._body()
                threading.Thread(target=self.ext.process_submit,args=(d,)).start()
                self._json({'status':'started'})
            elif self.path=='/delete':
                self._json(self.ext.delete_table(self._body().get('id')))
            elif self.path=='/settings':
                save_settings(self._body())
                self._json({'status':'saved'})
            elif self.path=='/close':
                self._json({'status':'closing'})
                if GTK_UI_AVAILABLE: GLib.idle_add(Gtk.main_quit)

    def __init__(self):
        super().__init__()
        self.status_data={"status":"idle","progress":0,"message":""}
        self.is_processing=False

    def add_arguments(self, pars): pass

    def scan_tables(self):
        tables=[]
        for e in self.svg.iter():
            eid=e.get('id','')
            if eid.startswith(self.TABLEMAN_ID_PREFIX):
                tables.append({
                    'id':eid,
                    'label':e.get(f'{{{INKSCAPE_LABEL_NS}}}label',eid),
                    'rows':int(e.get(f'{{{TABLEMAN_NS}}}rows','0')),
                    'cols':int(e.get(f'{{{TABLEMAN_NS}}}cols','0'))
                })
        return tables

    def load_table(self, tid):
        e=self.svg.getElementById(tid)
        if e is None: return None
        def jl(a,d='[]'):
            try: return json.loads(e.get(f'{{{TABLEMAN_NS}}}{a}',d))
            except: return json.loads(d)
        return {
            'id':tid,
            'label':e.get(f'{{{INKSCAPE_LABEL_NS}}}label',tid),
            'rows':int(e.get(f'{{{TABLEMAN_NS}}}rows','0')),
            'cols':int(e.get(f'{{{TABLEMAN_NS}}}cols','0')),
            'cell_width':float(e.get(f'{{{TABLEMAN_NS}}}cell-width','120')),
            'cell_height':float(e.get(f'{{{TABLEMAN_NS}}}cell-height','40')),
            'col_widths':jl('col-widths'),
            'row_heights':jl('row-heights'),
            'border_width':float(e.get(f'{{{TABLEMAN_NS}}}border-width','1')),
            'border_color':e.get(f'{{{TABLEMAN_NS}}}border-color','#333333'),
            'header_fill':e.get(f'{{{TABLEMAN_NS}}}header-fill','#bb86fc'),
            'body_fill':e.get(f'{{{TABLEMAN_NS}}}body-fill','#1e1e1e'),
            'banded_rows':e.get(f'{{{TABLEMAN_NS}}}banded-rows','0')=='1',
            'banded_color':e.get(f'{{{TABLEMAN_NS}}}banded-color','#2a2a2a'),
            'font_family':e.get(f'{{{TABLEMAN_NS}}}font-family','sans-serif'),
            'font_size':float(e.get(f'{{{TABLEMAN_NS}}}font-size','14')),
            'text_color':e.get(f'{{{TABLEMAN_NS}}}text-color','#ffffff'),
            'header_text_color':e.get(f'{{{TABLEMAN_NS}}}header-text-color','#000000'),
            'data':jl('data'),'merges':jl('merges'),'cell_styles':jl('cell-styles')
        }

    def delete_table(self, tid):
        if not tid: return {'status':'error','message':'No ID'}
        e=self.svg.getElementById(tid)
        if e is not None: e.getparent().remove(e); return {'status':'deleted'}
        return {'status':'error','message':'Not found'}

    def format_value(self, val, fmt, decs=2, sym=None):
        if not val or fmt == 'text': return str(val)
        raw = str(val).strip()
        try:
            import re
            # Smart parsing matching JS logic for id-ID and standard formats
            if 'Rp' in raw or (',' in raw and '.' not in raw) or raw.count('.') > 1:
                clean = raw.replace('Rp', '').replace('.', '').replace(',', '.').replace(' ', '')
                num = float(clean)
            else:
                num_str = re.sub(r'[^0-9.-]', '', raw)
                num = float(num_str)
        except: return str(val)

        try:
            locale_sep = ('.', ',') # thousands, decimal
            def fmt_num(n, d, sep=locale_sep):
                s = "{:,.{decs}f}".format(n, decs=d)
                return s.replace(',','X').replace('.',sep[1]).replace('X',sep[0])

            if fmt == 'number':
                return fmt_num(num, decs)
            elif fmt == 'percent':
                return f"{fmt_num(num, decs)}%"
            elif fmt == 'currency':
                return f"Rp{fmt_num(num, decs)}"
            elif fmt == 'currency_usd':
                s = "{:,.{decs}f}".format(num, decs=decs)
                return f"${s}"
            elif fmt == 'currency_eur':
                return f"€{fmt_num(num, decs, sep=(',', '.'))}"
            elif fmt == 'currency_custom':
                return f"{sym or '$'}{fmt_num(num, decs)}"
            elif fmt == 'currency_round':
                return f"Rp{fmt_num(num, 0)}"
            elif fmt == 'scientific':
                return "{:.{decs}e}".format(num, decs=decs)
            elif fmt.startswith('date'):
                import datetime
                try:
                    dt = datetime.datetime.fromisoformat(str(val))
                    if fmt == 'date_iso': return dt.strftime('%Y-%m-%d')
                    if fmt == 'date_us': return dt.strftime('%m/%d/%Y')
                    if fmt == 'date_long':
                        months = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember']
                        return f"{dt.day} {months[dt.month-1]} {dt.year}"
                    return dt.strftime('%d/%m/%Y')
                except: return str(val)
            return str(val)
        except: return str(val)

    def _merge_at(self, merges, r, c):
        for m in merges:
            if m['row']==r and m['col']==c:
                return m.get('rowspan',1), m.get('colspan',1)
        return None

    def _hidden(self, merges, r, c):
        for m in merges:
            mr,mc=m['row'],m['col']
            rs,cs=m.get('rowspan',1),m.get('colspan',1)
            if mr<=r<mr+rs and mc<=c<mc+cs and not(mr==r and mc==c): return True
        return False

    def render_table(self, group, data):
        rows=int(data.get('rows',3)); cols=int(data.get('cols',3))
        defW=float(data.get('cell_width',120)); defH=float(data.get('cell_height',40))

        # Per-column widths / per-row heights
        cw_arr=data.get('col_widths',[])
        rh_arr=data.get('row_heights',[])
        while len(cw_arr)<cols: cw_arr.append(defW)
        while len(rh_arr)<rows: rh_arr.append(defH)

        bw=float(data.get('border_width',1))
        bc=data.get('border_color','#333333')
        hfill=data.get('header_fill','#bb86fc')
        bfill=data.get('body_fill','#1e1e1e')
        is_banded=data.get('banded_rows',False)
        banded_color=data.get('banded_color','#2a2a2a')
        gf=data.get('font_family','sans-serif')
        gfs=float(data.get('font_size',14))
        gtc=data.get('text_color','#ffffff')
        ghtc=data.get('header_text_color','#000000')
        cdat=data.get('data',[])
        merges=data.get('merges',[])
        csty=data.get('cell_styles',[])

        # Precompute cumulative x/y
        cx=[0]
        for w in cw_arr: cx.append(cx[-1]+w)
        cy=[0]
        for h in rh_arr: cy.append(cy[-1]+h)

        for r in range(rows):
            for c in range(cols):
                if self._hidden(merges,r,c): continue

                x=cx[c]; y=cy[r]; isH=(r==0)
                mi=self._merge_at(merges,r,c)
                if mi:
                    cel_w=cx[c+mi[1]]-cx[c]
                    cel_h=cy[r+mi[0]]-cy[r]
                else:
                    cel_w=cw_arr[c]; cel_h=rh_arr[r]

                cs={}
                if r<len(csty) and c<len(csty[r]):
                    cs=csty[r][c] if isinstance(csty[r][c],dict) else {}

                fill = cs.get('fillColor')
                if not fill:
                    if isH:
                        fill = hfill
                    else:
                        if is_banded and r % 2 == 0:
                            fill = banded_color
                        else:
                            fill = bfill

                rect=Rectangle()
                rect.set('x',str(x));rect.set('y',str(y))
                rect.set('width',str(cel_w));rect.set('height',str(cel_h))
                rect.style={'stroke':bc,'stroke-width':str(bw)}
                if fill and fill != 'none':
                    rect.style['fill'] = fill
                else:
                    rect.style['fill'] = 'none'
                group.append(rect)

                ct=''
                if r<len(cdat) and c<len(cdat[r]): ct=str(cdat[r][c])
                if not ct: continue

                fmt = cs.get('format', 'auto')
                decs = int(cs.get('decimals', 2))
                sym = cs.get('currencySymbol')
                display_text = self.format_value(ct, fmt, decs, sym)
                
                font=cs.get('fontFamily') or gf
                fsize=float(cs.get('fontSize') or gfs)
                tcolor=cs.get('textColor') or (ghtc if isH else gtc)
                halign=cs.get('hAlign','center')
                valign=cs.get('vAlign','middle')
                rot_val=cs.get('rotation',0)  # can be 'stack' or number
                wrap=cs.get('wrap',False)

                amap={'left':'start','center':'middle','right':'end'}
                tanch=amap.get(halign,'middle')
                pad=4
                tx={'left':x+pad,'right':x+cel_w-pad}.get(halign,x+cel_w/2)
                ty={'top':y+fsize+pad,'bottom':y+cel_h-pad}.get(valign,y+cel_h/2+fsize/3)

                ts={'font-size':f'{fsize}px','font-family':font,'fill':tcolor,'text-anchor':tanch}
                if cs.get('bold'): ts['font-weight']='bold'
                if cs.get('italic'): ts['font-style']='italic'
                deco=[v for v in [cs.get('underline') and 'underline',cs.get('strikethrough') and 'line-through'] if v]
                if deco: ts['text-decoration']=' '.join(deco)

                te=TextElement(); te.style=ts

                is_stack = (str(rot_val) == 'stack')
                rotation = 0 if is_stack else float(rot_val or 0)

                if is_stack:
                    # Stack vertical: one character per line, centered
                    chars = list(display_text)
                    total_h = len(chars) * fsize * 1.2
                    sx = x + cel_w / 2
                    if valign == 'top':
                        sy = y + fsize + pad
                    elif valign == 'bottom':
                        sy = y + cel_h - total_h + fsize - pad
                    else:
                        sy = y + (cel_h - total_h) / 2 + fsize
                    ts['text-anchor'] = 'middle'
                    te.style = ts
                    for i, ch in enumerate(chars):
                        ts2=Tspan(); ts2.set('x',str(sx))
                        ts2.set('y',str(sy + i * fsize * 1.2))
                        ts2.text = ch; te.append(ts2)
                elif wrap:
                    avg=fsize*0.6; mx=max(1,int((cel_w-2*pad)/avg))
                    words=display_text.split(' '); lines=[]; cur=''
                    for w in words:
                        if cur and len(cur)+1+len(w)>mx: lines.append(cur); cur=w
                        else: cur=(cur+' '+w).strip()
                    if cur: lines.append(cur)
                    th=len(lines)*fsize*1.2
                    sy={'top':y+fsize+pad,'bottom':y+cel_h-th+fsize-pad}.get(valign,y+(cel_h-th)/2+fsize)
                    for i,ln in enumerate(lines):
                        ts2=Tspan(); ts2.set('x',str(tx)); ts2.set('y',str(sy+i*fsize*1.2))
                        ts2.text=ln; te.append(ts2)
                else:
                    te.set('x',str(tx)); te.set('y',str(ty))
                    ts2=Tspan(); ts2.text=display_text; te.append(ts2)

                if rotation!=0:
                    cxr=x+cel_w/2; cyr=y+cel_h/2
                    te.set('transform',f'rotate({rotation},{cxr},{cyr})')

                group.append(te)

    def process_submit(self, data):
        try:
            self.is_processing=True
            self.status_data={"status":"processing","progress":10,"message":"Parsing..."}
            import time; time.sleep(0.2)

            tid=data.get('id',''); label=data.get('label','Untitled')
            is_new=not tid or not self.svg.getElementById(tid)
            if is_new: tid=self.TABLEMAN_ID_PREFIX+str(uuid.uuid4())[:8]

            self.status_data.update({"progress":30,"message":"Preparing..."})
            etree.register_namespace(TABLEMAN_PREFIX,TABLEMAN_NS)

            if is_new:
                group=Group(); group.set('id',tid)
                layer=self.svg.get_current_layer()
                vc=self.svg.namedview.center
                cx_c=vc[0] if vc else 100; cy_c=vc[1] if vc else 100
                cw_arr=data.get('col_widths',[]); rh_arr=data.get('row_heights',[])
                tw=sum(cw_arr) if cw_arr else int(data.get('cols',3))*float(data.get('cell_width',120))
                th=sum(rh_arr) if rh_arr else int(data.get('rows',3))*float(data.get('cell_height',40))
                group.set('transform',f'translate({cx_c-tw/2},{cy_c-th/2})')
                layer.append(group)
            else:
                group=self.svg.getElementById(tid)
                for child in list(group): group.remove(child)

            self.status_data.update({"progress":50,"message":"Storing..."})
            group.set(f'{{{INKSCAPE_LABEL_NS}}}label',f'{label}')
            for k in ['rows','cols','cell_width','cell_height','border_width',
                       'border_color','header_fill','body_fill','banded_rows',
                       'banded_color','font_family','font_size','text_color','header_text_color']:
                val=data.get(k,'')
                if k=='banded_rows': val='1' if val else '0'
                group.set(f'{{{TABLEMAN_NS}}}{k.replace("_","-")}',str(val))

            group.set(f'{{{TABLEMAN_NS}}}data',json.dumps(data.get('data',[])))
            group.set(f'{{{TABLEMAN_NS}}}merges',json.dumps(data.get('merges',[])))
            group.set(f'{{{TABLEMAN_NS}}}cell-styles',json.dumps(data.get('cell_styles',[])))
            group.set(f'{{{TABLEMAN_NS}}}col-widths',json.dumps(data.get('col_widths',[])))
            group.set(f'{{{TABLEMAN_NS}}}row-heights',json.dumps(data.get('row_heights',[])))

            self.status_data.update({"progress":70,"message":"Rendering..."})
            time.sleep(0.2)
            self.render_table(group,data)

            self.status_data={"status":"completed","progress":100,
                "message":f"{'Created' if is_new else 'Updated'} '{label}'!"}
            if GTK_UI_AVAILABLE: time.sleep(0.8); GLib.idle_add(Gtk.main_quit)
        except Exception as ex:
            self.status_data={"status":"error","progress":0,"message":f"Error: {ex}"}
        finally:
            self.is_processing=False

    def run_web_ui(self):
        with socket.socket(socket.AF_INET,socket.SOCK_STREAM) as s:
            s.bind(('',0)); port=s.getsockname()[1]
        ui_dir=os.path.join(os.path.dirname(__file__),'ui')
        handler=lambda *a,**kw:self.WebUIHandler(*a,extension_instance=self,directory=ui_dir,**kw)
        server=socketserver.TCPServer(("",port),handler)
        t=threading.Thread(target=server.serve_forever); t.daemon=True; t.start()
        url=f"http://localhost:{port}/index.html"
        if GTK_UI_AVAILABLE:
            hb=GLib.timeout_add(100,lambda:True)
            GLib.idle_add(self._launch_gtk,url,server)
            Gtk.main(); GLib.source_remove(hb)
        else:
            webbrowser.open(url)
            while self.is_processing:
                import time; time.sleep(0.1)
        server.server_close()

    def _launch_gtk(self, url, server):
        w=Gtk.Window(title="Tableman — Table Manager")
        w.set_default_size(780,820)
        w.set_position(Gtk.WindowPosition.CENTER)
        w.set_resizable(True)
        wv=WebKit2.WebView(); w.add(wv); wv.load_uri(url)
        w.connect("destroy",lambda _:(server.shutdown(),Gtk.main_quit()))
        w.show_all()

    def effect(self):
        self.run_web_ui()

if __name__=='__main__':
    TablemanExtension().run()
