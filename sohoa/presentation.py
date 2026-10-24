"""Tkinter adapter. Workers communicate exclusively through an event queue."""
import json
import logging
from pathlib import Path
import queue
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from .application import merge, compare, JobContext
from .domain import HEADERS, Progress, AppError, Cancelled
from .excel import inspect, auto_mapping
from .paths import resolve_alias, resolve_data
from .school_scan import scan_schools

class App(ttk.Frame):
    def __init__(self, root, config_dir=None):
        super().__init__(root,padding=12); self.pack(fill='both',expand=True)
        self.root=root; self.events=queue.Queue(); self.stop=threading.Event(); self.busy=False
        self.sources=[]; self.mapvars={}; self.column_labels=[]
        self.config_dir=Path(config_dir) if config_dir else Path.home()/'THCS_Offline'
        self.config_dir.mkdir(exist_ok=True)
        self.settings_path=self.config_dir/'settings.json'
        try: self.settings=json.loads(self.settings_path.read_text(encoding='utf-8'))
        except (OSError,ValueError): self.settings={}
        self.output=tk.StringVar(value=self.settings.get('output',str(Path.cwd()/'output')))
        self.data=tk.StringVar(value=str(resolve_data(self.settings.get('data'))))
        self.alias=tk.StringVar(value=str(resolve_alias(self.config_dir,self.settings.get('alias'))))
        self.school=tk.StringVar(); self.header=tk.StringVar(value='1'); self.years=tk.StringVar(value='1984-2005')
        tabs=ttk.Notebook(self); tabs.pack(fill='x')
        a=ttk.Frame(tabs,padding=10); b=ttk.Frame(tabs,padding=10)
        tabs.add(a,text='1. Gộp và tách theo năm'); tabs.add(b,text='2. Đối soát và báo cáo')
        self.path_row(a,0,'Thư mục dữ liệu',self.data,lambda: self.choose_dir(self.data))
        ttk.Label(a,text='Gộp các file Excel, tách theo năm tốt nghiệp cột M: 1984.xlsx, 1985.xlsx…').grid(row=2,column=0,columnspan=3,sticky='w')
        self.merge_btn=ttk.Button(a,text='Gộp và tách',command=self.start_merge); self.merge_btn.grid(row=3,column=0,pady=8,sticky='w')
        self.scan_btn=ttk.Button(a,text='Quét danh sách tên trường để duyệt',command=self.start_scan)
        self.scan_btn.grid(row=3,column=1,pady=8,sticky='w')
        self.source_label=ttk.Label(b,text='Chưa chọn nguồn chuẩn',wraplength=650)
        self.source_label.grid(row=0,column=1,columnspan=2,sticky='w')
        ttk.Button(b,text='Chọn nguồn chuẩn',command=self.choose_sources).grid(row=0,column=0,sticky='w')
        self.path_row(b,1,'File trường gửi',self.school,self.choose_school)
        ttk.Label(b,text='Dòng tiêu đề').grid(row=2,column=0,sticky='w')
        ttk.Spinbox(b,from_=1,to=100,textvariable=self.header,width=8).grid(row=2,column=1,sticky='w')
        ttk.Button(b,text='Đọc tiêu đề / Ánh xạ',command=self.load_headers).grid(row=2,column=2)
        self.mapping_frame=ttk.LabelFrame(b,text='Ánh xạ cột (số hiệu bằng không bắt buộc)',padding=8)
        self.mapping_frame.grid(row=3,column=0,columnspan=3,sticky='ew',pady=6)
        for row,i in enumerate((4,6,12,1)):
            ttk.Label(self.mapping_frame,text=HEADERS[i]).grid(row=row,column=0,sticky='w')
            var=tk.StringVar(); combo=ttk.Combobox(self.mapping_frame,textvariable=var,state='readonly',width=65)
            combo.grid(row=row,column=1,sticky='ew'); self.mapvars[i]=(var,combo)
        ttk.Label(b,text='Năm kiểm tra').grid(row=4,column=0,sticky='w')
        ttk.Entry(b,textvariable=self.years).grid(row=4,column=1,sticky='ew')
        ttk.Label(b,text='Ví dụ: 1984-2005 hoặc 1990,1992').grid(row=4,column=2)
        self.compare_btn=ttk.Button(b,text='Xuất báo cáo đối soát',command=self.start_compare)
        self.compare_btn.grid(row=5,column=0,pady=8,sticky='w')
        self.path_row(self,1,'Thư mục kết quả',self.output,lambda:self.choose_dir(self.output),pack=True)
        self.progress=ttk.Progressbar(self,mode='indeterminate'); self.progress.pack(fill='x',pady=8)
        self.status=ttk.Label(self,text='Sẵn sàng'); self.status.pack(anchor='w')
        self.console=tk.Text(self,height=16,state='disabled',wrap='word'); self.console.pack(fill='both',expand=True,pady=6)
        self.cancel_btn=ttk.Button(self,text='Hủy tác vụ',command=self.stop.set,state='disabled'); self.cancel_btn.pack(anchor='e')
        root.protocol('WM_DELETE_WINDOW',self.close)
        root.after(100,self.poll)

    def path_row(self,parent,row,label,var,command,pack=False):
        if pack:
            frame=ttk.Frame(parent); frame.pack(fill='x',pady=8); parent=frame; row=0
        ttk.Label(parent,text=label).grid(row=row,column=0,sticky='w',padx=4,pady=4)
        ttk.Entry(parent,textvariable=var,width=70).grid(row=row,column=1,sticky='ew')
        ttk.Button(parent,text='Chọn…',command=command).grid(row=row,column=2,padx=4)
        parent.columnconfigure(1,weight=1)

    def choose_dir(self,var):
        value=filedialog.askdirectory()
        if value: var.set(value)

    def choose_file(self,var):
        value=filedialog.askopenfilename(filetypes=[('Excel / CSV','*.xlsx *.csv')])
        if value: var.set(value)

    def choose_sources(self):
        files=filedialog.askopenfilenames(filetypes=[('Excel','*.xlsx')])
        if files:
            self.sources=list(files); self.source_label.configure(text='; '.join(Path(p).name for p in files))

    def choose_school(self):
        self.choose_file(self.school)
        for var,_ in self.mapvars.values(): var.set('')

    def load_headers(self):
        if self.busy: return
        try:
            headers=inspect(self.school.get(),int(self.header.get()))
            labels=['(Không dùng)']+[f'{i+1}: {h}' for i,h in enumerate(headers)]
            mapping=auto_mapping(headers)
            saved=self.settings.get('mapping_profiles',{}).get(json.dumps(headers,ensure_ascii=False),{})
            for i,(var,combo) in self.mapvars.items():
                combo.configure(values=labels)
                index=saved.get(str(i),mapping.get(i,-1))
                var.set(labels[index+1] if isinstance(index,int) and -1<=index<len(headers) else labels[0])
            self.column_labels=labels; self.loaded_school=(self.school.get(),self.header.get()); self.loaded_headers=headers
        except Exception as exc: messagebox.showerror('Không đọc được tiêu đề',str(exc))

    def start_scan(self):
        data,output=self.data.get(),self.output.get()
        self.start(lambda ctx:scan_schools(data,output,ctx))

    def start_merge(self):
        data,output,alias=self.data.get(),self.output.get(),self.alias.get()
        if not Path(data).is_dir():
            messagebox.showerror('Thiếu thư mục dữ liệu',
                'Không tìm thấy thư mục dữ liệu. Bấm Chọn… tại mục Thư mục dữ liệu để chọn nơi chứa các file Excel năm.\n' + data)
            return
        self.start(lambda ctx:merge(data,output,alias,ctx))

    def start_compare(self):
        try:
            if getattr(self,'loaded_school',None)!=(self.school.get(),self.header.get()):
                raise AppError('Hãy đọc tiêu đề và xác nhận ánh xạ cho file đang chọn.')
            mapping={i:self.column_labels.index(var.get())-1 for i,(var,_) in self.mapvars.items() if var.get() and var.get()!='(Không dùng)'}
            years=set()
            for token in self.years.get().split(','):
                parts=token.strip().split('-')
                if len(parts)==1: years.add(int(parts[0]))
                elif len(parts)==2 and int(parts[0])<=int(parts[1]): years.update(range(int(parts[0]),int(parts[1])+1))
                else: raise ValueError('Khoảng năm không hợp lệ')
            if not years or not years.issubset(set(range(1984,2006))): raise ValueError('Năm phải thuộc 1984–2005')
            profiles=self.settings.setdefault('mapping_profiles',{})
            profiles[json.dumps(self.loaded_headers,ensure_ascii=False)]={str(k):v for k,v in mapping.items()}
            args=(list(self.sources),self.school.get(),self.output.get(),int(self.header.get()),mapping,years)
            self.start(lambda ctx:compare(*args,context=ctx))
        except Exception as exc: messagebox.showerror('Kiểm tra đầu vào',str(exc))

    def start(self,fn):
        if self.busy: return
        self.settings.update(output=self.output.get(),data=self.data.get(),alias=self.alias.get())
        self.settings_path.write_text(json.dumps(self.settings,ensure_ascii=False,indent=2),encoding='utf-8')
        self.busy=True; self.stop.clear()
        self.merge_btn.configure(state='disabled'); self.scan_btn.configure(state='disabled'); self.compare_btn.configure(state='disabled'); self.cancel_btn.configure(state='normal')
        self.progress.configure(mode='indeterminate'); self.progress.start()
        def worker():
            try:
                path=fn(JobContext(self.events.put,self.stop.is_set))
                self.events.put(('done',str(path)))
            except Cancelled as exc: self.events.put(('cancel',str(exc)))
            except Exception as exc:
                with open(self.config_dir/'errors.log','a',encoding='utf-8') as f: f.write(traceback.format_exc()+'\n')
                self.events.put(('error',str(exc)))
        threading.Thread(target=worker,daemon=True).start()

    def poll(self):
        for _ in range(200):
            try: event=self.events.get_nowait()
            except queue.Empty: break
            if isinstance(event,Progress):
                self.log(event.message); self.status.configure(text=event.message)
                if event.total:
                    self.progress.stop()
                    self.progress.configure(mode='determinate',maximum=event.total,value=event.current)
                elif str(self.progress['mode'])!='indeterminate':
                    self.progress.configure(mode='indeterminate'); self.progress.start()
            else:
                kind,message=event; self.busy=False; self.progress.stop()
                self.merge_btn.configure(state='normal'); self.scan_btn.configure(state='normal'); self.compare_btn.configure(state='normal'); self.cancel_btn.configure(state='disabled')
                label={'done':'Đã xuất kết quả','cancel':'Đã hủy','error':'Thất bại'}[kind]
                self.log(f'{label}: {message}'); self.status.configure(text=label)
                (messagebox.showerror if kind=='error' else messagebox.showinfo)(label,message)
        self.root.after(100,self.poll)

    def log(self,message):
        self.console.configure(state='normal'); self.console.insert('end',message+'\n')
        if int(self.console.index('end-1c').split('.')[0])>3000: self.console.delete('1.0','500.0')
        self.console.see('end'); self.console.configure(state='disabled')
        with open(self.config_dir/'activity.log','a',encoding='utf-8') as f: f.write(message+'\n')

    def close(self):
        if self.busy:
            self.stop.set(); self.status.configure(text='Đang yêu cầu hủy; chờ tác vụ dừng rồi đóng cửa sổ.'); return
        self.root.destroy()

def launch():
    root=tk.Tk(); root.title('Đối soát số hóa tốt nghiệp THCS — Offline')
    root.geometry('1100x820'); root.minsize(850,700)
    ttk.Style(root).theme_use('clam')
    App(root); root.mainloop()
