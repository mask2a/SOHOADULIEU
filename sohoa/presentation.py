"""Tkinter adapter. Workers communicate exclusively through an event queue."""
import json
import logging
from pathlib import Path
import queue
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from .application import merge, compare, compare_batch, JobContext
from .domain import HEADERS, Progress, AppError, Cancelled
from .excel import inspect, auto_mapping
from .paths import resolve_alias, resolve_data

class App(ttk.Frame):
    def __init__(self, root, config_dir=None):
        super().__init__(root,padding=12); self.pack(fill='both',expand=True)
        self.root=root; self.events=queue.Queue(); self.progress_events=queue.Queue(maxsize=1)
        self.results=queue.Queue()
        self.stop=threading.Event(); self.busy=False
        self.sources=[]; self.mapvars={}; self.column_labels=[]
        self.config_dir=Path(config_dir) if config_dir else Path.home()/'THCS_Offline'
        self.config_dir.mkdir(exist_ok=True)
        self.settings_path=self.config_dir/'settings.json'
        try: self.settings=json.loads(self.settings_path.read_text(encoding='utf-8'))
        except (OSError,ValueError): self.settings={}
        self.output=tk.StringVar(value=self.settings.get('output',str(Path.cwd()/'output')))
        self.data=tk.StringVar(value=str(resolve_data(self.settings.get('data'))))
        self.alias=tk.StringVar(value=str(resolve_alias(self.config_dir,self.settings.get('alias'))))
        self.school=tk.StringVar(); self.header=tk.StringVar(value='1'); self.years=tk.StringVar()
        tabs=ttk.Notebook(self); tabs.pack(fill='x'); self.tabs=tabs
        a=ttk.Frame(tabs,padding=10); b=ttk.Frame(tabs,padding=10)
        c=ttk.Frame(tabs,padding=10)
        tabs.add(a,text='1. Gộp và tách theo năm'); tabs.add(b,text='2. Đối soát và báo cáo')
        tabs.add(c,text='3. Đối soát theo lô')
        self.path_row(a,0,'Thư mục dữ liệu',self.data,lambda: self.choose_dir(self.data))
        ttk.Label(a,text='Gộp các file Excel, tách theo năm tốt nghiệp cột M: 1984.xlsx, 1985.xlsx…').grid(row=2,column=0,columnspan=3,sticky='w')
        self.merge_btn=ttk.Button(a,text='Gộp và tách theo năm',command=self.start_merge); self.merge_btn.grid(row=3,column=0,pady=8,sticky='w')
        self.source_label=ttk.Label(b,text='Chưa chọn file nguồn chuẩn',wraplength=650)
        self.source_label.grid(row=0,column=1,columnspan=2,sticky='w')
        ttk.Button(b,text='Chọn file nguồn chuẩn',command=self.choose_sources).grid(row=0,column=0,sticky='w')
        self.path_row(b,1,'File trường gửi theo năm',self.school,self.choose_school)
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
        ttk.Label(b,text='Để trống để kiểm tra mọi năm; hoặc nhập năm/khoảng năm').grid(row=4,column=2)
        self.compare_btn=ttk.Button(b,text='Xuất báo cáo đối soát',command=self.start_compare)
        self.compare_btn.grid(row=5,column=0,pady=8,sticky='w')
        self.source_dir=tk.StringVar(); self.school_dir=tk.StringVar()
        self.path_row(c,0,'Thư mục nguồn chuẩn',self.source_dir,lambda:self.choose_dir(self.source_dir))
        self.path_row(c,1,'Thư mục trường gửi',self.school_dir,lambda:self.choose_dir(self.school_dir))
        ttk.Label(c,text='Tự ghép THCS.CQ.<năm>_output.xlsx với <năm>.xlsx; chỉ kiểm tra năm trùng.').grid(
            row=2,column=0,columnspan=3,sticky='w',pady=4)
        self.compare_batch_btn=ttk.Button(c,text='Đối soát các năm trùng',command=self.start_compare_batch)
        self.compare_batch_btn.grid(row=3,column=0,pady=8,sticky='w')
        self.path_row(self,1,'Thư mục kết quả',self.output,lambda:self.choose_dir(self.output),pack=True)
        progress_row=ttk.Frame(self)
        progress_row.pack(fill='x',pady=8)
        self.progress=ttk.Progressbar(progress_row,mode='indeterminate')
        self.progress.pack(side='left',fill='x',expand=True)
        self.progress_percent=tk.StringVar(value='0.0%')
        ttk.Label(
            progress_row,textvariable=self.progress_percent,width=7,anchor='e',
            font=('Segoe UI',11,'bold'),
        ).pack(side='left',padx=(8,0))
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
        path=filedialog.askopenfilename(filetypes=[('Excel','*.xlsx')])
        if path:
            self.sources=[path]; self.source_label.configure(text=Path(path).name)

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
            years=None
            year_text=self.years.get().strip()
            if year_text:
                years=set()
                for token in year_text.split(','):
                    parts=token.strip().split('-')
                    if len(parts)==1 and parts[0].isdigit():
                        years.add(int(parts[0]))
                    elif (len(parts)==2 and all(part.isdigit() for part in parts)
                          and int(parts[0])<=int(parts[1])):
                        years.update(range(int(parts[0]),int(parts[1])+1))
                    else:
                        raise ValueError('Nhập năm hoặc khoảng năm hợp lệ, ví dụ 1983 hoặc 2026-2030')
            profiles=self.settings.setdefault('mapping_profiles',{})
            profiles[json.dumps(self.loaded_headers,ensure_ascii=False)]={str(k):v for k,v in mapping.items()}
            args=(list(self.sources),self.school.get(),self.output.get(),int(self.header.get()),mapping,years)
            self.start(lambda ctx:compare(*args,context=ctx))
        except Exception as exc: messagebox.showerror('Kiểm tra đầu vào',str(exc))

    def start_compare_batch(self):
        source_dir,school_dir,output=self.source_dir.get(),self.school_dir.get(),self.output.get()
        if not Path(source_dir).is_dir():
            messagebox.showerror('Thiếu thư mục nguồn',f'Chọn thư mục có file THCS.CQ.<năm>_output.xlsx.\n{source_dir}')
            return
        if not Path(school_dir).is_dir():
            messagebox.showerror('Thiếu thư mục trường gửi',f'Chọn thư mục có file <năm>.xlsx.\n{school_dir}')
            return
        self.start(lambda ctx:compare_batch(source_dir,school_dir,output,context=ctx))

    def start(self,fn):
        if self.busy: return
        self.settings.update(output=self.output.get(),data=self.data.get(),alias=self.alias.get())
        self.settings_path.write_text(json.dumps(self.settings,ensure_ascii=False,indent=2),encoding='utf-8')
        self.busy=True; self.stop.clear()
        self.merge_btn.configure(state='disabled')
        self.compare_btn.configure(state='disabled'); self.compare_batch_btn.configure(state='disabled')
        self.cancel_btn.configure(state='normal')
        self.progress_percent.set('0.0%')
        self.progress.configure(mode='indeterminate'); self.progress.start()
        def worker():
            try:
                path=fn(JobContext(self.emit_progress,self.stop.is_set))
                self.results.put(('done',str(path)))
            except Cancelled as exc: self.results.put(('cancel',str(exc)))
            except Exception as exc:
                with open(self.config_dir/'errors.log','a',encoding='utf-8') as f: f.write(traceback.format_exc()+'\n')
                self.results.put(('error',str(exc)))
        threading.Thread(target=worker,daemon=True).start()

    def emit_progress(self, event):
        if event.log_message:
            self.events.put(event)
        try:
            self.progress_events.put_nowait(event)
        except queue.Full:
            try: self.progress_events.get_nowait()
            except queue.Empty: pass
            self.progress_events.put_nowait(event)

    def poll(self):
        try: result=self.results.get_nowait()
        except queue.Empty: result=None
        if result is not None:
            while True:
                try: self.events.get_nowait()
                except queue.Empty: break
            while True:
                try: self.progress_events.get_nowait()
                except queue.Empty: break
            self.finish(result)
            self.root.after(100,self.poll)
            return
        for _ in range(200):
            try: event=self.events.get_nowait()
            except queue.Empty: break
            if isinstance(event,Progress) and self.busy and event.log_message:
                self.log(event.message)
        try: event=self.progress_events.get_nowait()
        except queue.Empty: event=None
        if isinstance(event,Progress) and self.busy:
            self.status.configure(text=event.message)
            if event.total:
                self.progress.stop()
                self.progress.configure(mode='determinate',maximum=event.total,value=event.current)
                self.progress_percent.set(f'{event.current/event.total:.1%}')
        try: result=self.results.get_nowait()
        except queue.Empty: result=None
        if result is not None:
            self.finish(result)
        self.root.after(100,self.poll)

    def finish(self, result):
        kind,message=result; self.busy=False; self.progress.stop()
        self.merge_btn.configure(state='normal')
        self.compare_btn.configure(state='normal'); self.compare_batch_btn.configure(state='normal')
        self.cancel_btn.configure(state='disabled')
        if kind=='done':
            self.progress_percent.set('100.0%')
        label={'done':'Đã xuất kết quả','cancel':'Đã hủy','error':'Thất bại'}[kind]
        self.log(f'{label}: {message}'); self.status.configure(text=label)
        (messagebox.showerror if kind=='error' else messagebox.showinfo)(label,message)

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
