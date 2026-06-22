
import json, os, sys, shutil, subprocess, winreg, ctypes
import tkinter as tk
from tkinter import messagebox, ttk
from pathlib import Path

def S(x): return x.encode('ascii').decode('unicode_escape')
APP_ID = 'LitSearchPro_Generic'
DISPLAY_NAME = '科研文献与实验室安全管理平台'
APP_FOLDER = '科研文献与实验室安全管理平台'
STALE_KEYS = ['LitSearchPro_Generic']
OLD_START_MENU_NAMES = ['科研文献与实验室安全管理平台']
DATA_KIND='client'
def is_admin():
    try: return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception: return False
def reg_views():
    views=[0]
    for attr in ('KEY_WOW64_64KEY','KEY_WOW64_32KEY'):
        v=getattr(winreg,attr,0)
        if v and v not in views: views.append(v)
    return views
def delete_key(root, subkey, view=0):
    try:
        with winreg.OpenKey(root, subkey, 0, winreg.KEY_READ|winreg.KEY_WRITE|view) as k:
            while True:
                try: delete_key(root, subkey+'\\'+winreg.EnumKey(k,0), view)
                except OSError: break
        winreg.DeleteKeyEx(root, subkey, view, 0) if hasattr(winreg,'DeleteKeyEx') else winreg.DeleteKey(root, subkey)
    except Exception: pass
def remove_registry():
    base='Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\'
    for key in [APP_ID]+STALE_KEYS:
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            for view in reg_views(): delete_key(root, base+key, view)
def safe_remove(path):
    p=Path(path)
    try:
        if p.is_dir(): shutil.rmtree(p,ignore_errors=True)
        elif p.exists(): p.unlink()
    except Exception: pass
def remove_shortcuts():
    bases=[Path(os.environ.get('APPDATA',''))/'Microsoft'/'Windows'/'Start Menu'/'Programs', Path(os.environ.get('ProgramData',r'C:\ProgramData'))/'Microsoft'/'Windows'/'Start Menu'/'Programs']
    for base in bases:
        for name in OLD_START_MENU_NAMES+[APP_FOLDER,DISPLAY_NAME]: safe_remove(base/name)
    for desk in [Path(os.environ.get('PUBLIC',str(Path.home())))/'Desktop', Path(os.environ.get('USERPROFILE',str(Path.home())))/'Desktop']:
        for name in [DISPLAY_NAME+'.lnk']: safe_remove(desk/name)
def load_json(path):
    try: return json.loads(Path(path).read_text(encoding='utf-8'))
    except Exception: return {}
def client_targets():
    app=Path(os.environ.get('APPDATA',str(Path.home())))/'LitSearchPro'; settings=app/'settings_v11.json'; cfg=load_json(settings)
    return {'api':settings,'settings':settings,'db':Path(os.path.expandvars(cfg.get('database_path') or str(app/'library_v11.db'))),'pdf':Path(os.path.expandvars(cfg.get('pdf_dir') or str(app/'pdfs'))),'backups':app/'backups','models':Path(os.path.expandvars(cfg.get('local_ai_model_dir') or str(app/'models')))}
def server_targets():
    app=Path(os.environ.get('APPDATA',str(Path.home())))/'LitSearchProServer'; cfgfile=app/'server_config.json'; cfg=load_json(cfgfile); data=Path(os.path.expandvars(cfg.get('data_dir') or str(app)))
    return {'config':cfgfile,'db':data/'collaboration_server.db','uploads':data/'uploads','logs':[data/'server_crash.log',data/'server_native_crash.log',app/'server_crash.log',app/'server_native_crash.log']}
def clear_api_keys():
    t=client_targets()['api']; data=load_json(t); changed=False
    for key in ['s2_key','deepseek_key','qwen_key','openai_key','gemini_key','doubao_key','wenxin_key','zotero_key','collaboration_token']:
        if data.get(key): data[key]=''; changed=True
    if changed: Path(t).write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
class Uninstaller(tk.Tk):
    def __init__(self,install_dir):
        super().__init__(); self.install_dir=Path(install_dir); self.title(S(r'\u5378\u8f7d ')+DISPLAY_NAME); self.geometry('720x560'); self.minsize(560,430); self.configure(bg='#F6F8FC')
        self.status=tk.StringVar(value=S(r'\u51c6\u5907\u5378\u8f7d')); self.detail=tk.StringVar(value=S(r'\u8bf7\u9009\u62e9\u9700\u8981\u4e00\u5e76\u5220\u9664\u7684\u6570\u636e\u3002\u9ed8\u8ba4\u4fdd\u7559\u79d1\u7814\u6570\u636e\u3002'))
        self.vars={'api':tk.BooleanVar(value=True),'settings':tk.BooleanVar(value=False),'db':tk.BooleanVar(value=False),'pdf':tk.BooleanVar(value=False),'backups':tk.BooleanVar(value=False),'models':tk.BooleanVar(value=False)} if DATA_KIND=='client' else {'config':tk.BooleanVar(value=False),'db':tk.BooleanVar(value=False),'uploads':tk.BooleanVar(value=False),'logs':tk.BooleanVar(value=True)}; self.build()
    def build(self):
        self.grid_columnconfigure(0,weight=1); self.grid_rowconfigure(1,weight=1); head=tk.Frame(self,bg='#F6F8FC'); head.grid(row=0,column=0,sticky='ew',padx=22,pady=(16,6)); tk.Label(head,text=S(r'\u5378\u8f7d ')+DISPLAY_NAME,bg='#F6F8FC',fg='#0F172A',font=('Microsoft YaHei UI',16,'bold')).pack(anchor='w'); tk.Label(head,text=S(r'\u7a0b\u5e8f\u6587\u4ef6\u4f1a\u88ab\u5220\u9664\uff1b\u6570\u636e\u5e93\u3001PDF\u3001\u670d\u52a1\u5668\u6570\u636e\u7b49\u7531\u60a8\u51b3\u5b9a\u662f\u5426\u5220\u9664\u3002'),bg='#F6F8FC',fg='#64748B',wraplength=650,justify='left').pack(anchor='w')
        card=tk.Frame(self,bg='white',highlightbackground='#DCE3EF',highlightthickness=1); card.grid(row=1,column=0,sticky='nsew',padx=22,pady=8); targets=client_targets() if DATA_KIND=='client' else server_targets()
        rows=[('api',S(r'\u5220\u9664 API Key / Token\uff08\u63a8\u8350\uff09'),S(r'\u53ea\u6e05\u7a7a\u5bc6\u94a5\uff0c\u4e0d\u5220\u9664\u6587\u732e\u6570\u636e\u5e93\u548c PDF\u3002')),('settings',S(r'\u5220\u9664\u7528\u6237\u8bbe\u7f6e'),str(targets.get('settings',''))),('db',S(r'\u5220\u9664\u6587\u732e\u6570\u636e\u5e93'),str(targets.get('db',''))),('pdf',S(r'\u5220\u9664 PDF \u4fdd\u5b58\u76ee\u5f55'),str(targets.get('pdf',''))),('backups',S(r'\u5220\u9664\u5907\u4efd\u76ee\u5f55'),str(targets.get('backups',''))),('models',S(r'\u5220\u9664\u672c\u5730 AI \u6a21\u578b\u76ee\u5f55'),str(targets.get('models','')))] if DATA_KIND=='client' else [('config',S(r'\u5220\u9664\u670d\u52a1\u5668\u914d\u7f6e'),str(targets['config'])),('db',S(r'\u5220\u9664\u670d\u52a1\u5668\u6570\u636e\u5e93/\u8d26\u53f7/\u5ba1\u6279/\u804a\u5929\u8bb0\u5f55'),str(targets['db'])),('uploads',S(r'\u5220\u9664\u670d\u52a1\u5668\u4e0a\u4f20\u6587\u4ef6'),str(targets['uploads'])),('logs',S(r'\u5220\u9664\u670d\u52a1\u5668\u65e5\u5fd7\uff08\u63a8\u8350\uff09'),'server_crash.log / server_native_crash.log')]
        for r,(key,title,desc) in enumerate(rows):
            f=tk.Frame(card,bg='white'); f.grid(row=r,column=0,sticky='ew',padx=16,pady=5); f.grid_columnconfigure(0,weight=1); tk.Checkbutton(f,text=title,variable=self.vars[key],bg='white',activebackground='white',font=('Microsoft YaHei UI',10,'bold')).grid(row=0,column=0,sticky='w'); tk.Label(f,text=desc,bg='white',fg='#64748B',wraplength=620,justify='left').grid(row=1,column=0,sticky='w',padx=24)
        bottom=tk.Frame(self,bg='#F6F8FC'); bottom.grid(row=2,column=0,sticky='ew',padx=22,pady=(4,14)); bottom.grid_columnconfigure(0,weight=1); self.progress=ttk.Progressbar(bottom,mode='determinate',maximum=100); self.progress.grid(row=0,column=0,columnspan=3,sticky='ew',pady=(0,6)); tk.Label(bottom,textvariable=self.status,bg='#F6F8FC',fg='#2563EB',font=('Microsoft YaHei UI',10,'bold')).grid(row=1,column=0,sticky='w'); ttk.Button(bottom,text=S(r'\u53d6\u6d88'),command=self.destroy).grid(row=1,column=1,padx=8); ttk.Button(bottom,text=S(r'\u5378\u8f7d'),command=self.do_uninstall).grid(row=1,column=2)
    def step(self,i,total,status,detail=''):
        self.progress['value']=int(i/total*100); self.status.set(status); self.detail.set(detail); self.update_idletasks()
    def do_uninstall(self):
        if not is_admin(): messagebox.showwarning(S(r'\u6743\u9650\u4e0d\u8db3'),S(r'\u8bf7\u4ee5\u7ba1\u7406\u5458\u6743\u9650\u8fd0\u884c\u5378\u8f7d\u7a0b\u5e8f\u3002')); return
        if not messagebox.askyesno(S(r'\u786e\u8ba4\u5378\u8f7d'),S(r'\u786e\u5b9a\u6309\u7167\u5f53\u524d\u9009\u62e9\u5378\u8f7d\u5417\uff1f')): return
        total=8; self.step(1,total,S(r'\u6b63\u5728\u5220\u9664\u5feb\u6377\u65b9\u5f0f'),S(r'\u5f00\u59cb\u83dc\u5355\u4e0e\u684c\u9762\u5165\u53e3')); remove_shortcuts(); targets=client_targets() if DATA_KIND=='client' else server_targets(); idx=2
        if DATA_KIND=='client':
            if self.vars['api'].get(): self.step(idx,total,S(r'\u6b63\u5728\u6e05\u9664 API Key'),str(targets['api'])); clear_api_keys()
            else: self.step(idx,total,S(r'\u4fdd\u7559 API Key'),S(r'\u7528\u6237\u9009\u62e9\u4fdd\u7559'))
            idx+=1
            for key in ['settings','db','pdf','backups','models']:
                self.step(idx,total,(S(r'\u6b63\u5728\u5220\u9664 ') if self.vars[key].get() else S(r'\u4fdd\u7559 '))+key,str(targets[key]));
                if self.vars[key].get(): safe_remove(targets[key])
                idx+=1
        else:
            for key in ['config','db','uploads','logs']:
                self.step(idx,total,(S(r'\u6b63\u5728\u5220\u9664 ') if self.vars[key].get() else S(r'\u4fdd\u7559 '))+key,str(targets[key]));
                if self.vars[key].get():
                    if key=='logs': [safe_remove(p) for p in targets[key]]
                    else: safe_remove(targets[key])
                idx+=1
        self.step(total-1,total,S(r'\u6b63\u5728\u5220\u9664 Windows \u5378\u8f7d\u9879'),APP_ID); remove_registry(); self.step(total,total,S(r'\u6b63\u5728\u5220\u9664\u7a0b\u5e8f\u76ee\u5f55'),str(self.install_dir)); bat=Path(os.environ.get('TEMP',str(self.install_dir)))/(APP_ID+'_cleanup.cmd'); bat.write_text('@echo off\r\ntimeout /t 2 /nobreak >nul\r\nrmdir /s /q "'+str(self.install_dir)+'"\r\ndel "%~f0"\r\n',encoding='mbcs',errors='ignore'); subprocess.Popen(['cmd','/c',str(bat)],creationflags=0x08000000); messagebox.showinfo(S(r'\u5378\u8f7d'),S(r'\u5378\u8f7d\u5df2\u5f00\u59cb\uff0c\u7a0b\u5e8f\u76ee\u5f55\u5c06\u81ea\u52a8\u5220\u9664\u3002')); self.destroy()
def main(): Uninstaller(Path(sys.argv[1]) if len(sys.argv)>1 else Path(__file__).resolve().parent).mainloop()
if __name__=='__main__': main()
