
import os, sys, shutil, subprocess, winreg, ctypes
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

def S(x): return x.encode('ascii').decode('unicode_escape')
APP_ID = 'LitSearchPro_Generic'
DISPLAY_NAME = '科研文献与实验室安全管理平台'
APP_FOLDER = '科研文献与实验室安全管理平台'
MAIN_EXE = 'LitSearchPro_Generic_v22.1.21.exe'
PAYLOAD_UNINSTALL_EXE = 'LitSearchPro_Generic_v22.1.21_Uninstall.exe'
INSTALLED_UNINSTALL_EXE = 'Uninstall.exe'
CHANGELOG = 'CHANGELOG.md'
ICON = 'generic_logo.ico'
PUBLISHER = 'LitSearchPro Contributors'
FALLBACK_DIR = r'C:\Program Files\LitSearchPro Generic'
STALE_KEYS = ['LitSearchPro_Generic']
OLD_START_MENU_NAMES = ['科研文献与实验室安全管理平台']

def resource_path(name): return Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent)) / name
def is_admin():
    try: return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception: return False
def reg_views():
    views=[0]
    for attr in ('KEY_WOW64_64KEY','KEY_WOW64_32KEY'):
        v=getattr(winreg, attr, 0)
        if v and v not in views: views.append(v)
    return views
def delete_key(root, subkey, view=0):
    try:
        with winreg.OpenKey(root, subkey, 0, winreg.KEY_READ | winreg.KEY_WRITE | view) as k:
            while True:
                try: delete_key(root, subkey+'\\'+winreg.EnumKey(k,0), view)
                except OSError: break
        winreg.DeleteKeyEx(root, subkey, view, 0) if hasattr(winreg,'DeleteKeyEx') else winreg.DeleteKey(root, subkey)
    except Exception: pass
def read_reg(root, subkey, name, view=0):
    try:
        with winreg.OpenKey(root, subkey, 0, winreg.KEY_READ | view) as k: return winreg.QueryValueEx(k,name)[0]
    except Exception: return ''
def uninstall_base(): return r'Software\Microsoft\Windows\CurrentVersion\Uninstall'
def find_install_dir():
    for key in [APP_ID]+STALE_KEYS:
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            for view in reg_views():
                loc=read_reg(root, uninstall_base()+'\\'+key, 'InstallLocation', view)
                if loc: return Path(loc)
    return Path(FALLBACK_DIR)
def cleanup_registry():
    for key in STALE_KEYS:
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            for view in reg_views(): delete_key(root, uninstall_base()+'\\'+key, view)
def write_uninstall_entry(install_dir):
    cleanup_registry(); view=getattr(winreg,'KEY_WOW64_64KEY',0); sub=uninstall_base()+'\\'+APP_ID
    with winreg.CreateKeyEx(winreg.HKEY_LOCAL_MACHINE, sub, 0, winreg.KEY_WRITE | view) as k:
        vals={'DisplayName':DISPLAY_NAME,'DisplayVersion':'22.1.21','Publisher':PUBLISHER,'InstallLocation':str(install_dir),'DisplayIcon':str(install_dir/ICON),'UninstallString':f'"{install_dir/INSTALLED_UNINSTALL_EXE}"','QuietUninstallString':f'"{install_dir/INSTALLED_UNINSTALL_EXE}" /quiet'}
        for n,v in vals.items(): winreg.SetValueEx(k,n,0,winreg.REG_SZ,v)
        size=sum(p.stat().st_size for p in install_dir.rglob('*') if p.is_file())//1024
        winreg.SetValueEx(k,'EstimatedSize',0,winreg.REG_DWORD,int(size)); winreg.SetValueEx(k,'NoModify',0,winreg.REG_DWORD,1); winreg.SetValueEx(k,'NoRepair',0,winreg.REG_DWORD,1)
    for view in reg_views(): delete_key(winreg.HKEY_CURRENT_USER, sub, view)
def safe_remove(path):
    p=Path(path)
    try:
        if p.is_dir(): shutil.rmtree(p, ignore_errors=True)
        elif p.exists(): p.unlink()
    except Exception: pass
def shortcut_script(link, target, workdir, icon):
    vbs=Path(os.environ.get('TEMP',str(Path.home()))) / (APP_ID+'_shortcut.vbs')
    content='Set ws = WScript.CreateObject("WScript.Shell")\nSet s = ws.CreateShortcut("%s")\ns.TargetPath = "%s"\ns.WorkingDirectory = "%s"\ns.IconLocation = "%s"\ns.Save\n' % (link,target,workdir,icon)
    vbs.write_text(content,encoding='mbcs',errors='ignore'); subprocess.run(['cscript','//nologo',str(vbs)],check=False,creationflags=0x08000000)
    try: vbs.unlink()
    except Exception: pass
def remove_old_shortcuts():
    bases=[Path(os.environ.get('APPDATA',''))/'Microsoft'/'Windows'/'Start Menu'/'Programs', Path(os.environ.get('ProgramData',r'C:\ProgramData'))/'Microsoft'/'Windows'/'Start Menu'/'Programs']
    for base in bases:
        for name in OLD_START_MENU_NAMES+[APP_FOLDER,DISPLAY_NAME]: safe_remove(base/name)
def create_shortcuts(install_dir, desktop=True):
    remove_old_shortcuts(); common=Path(os.environ.get('ProgramData',r'C:\ProgramData'))/'Microsoft'/'Windows'/'Start Menu'/'Programs'/DISPLAY_NAME; common.mkdir(parents=True,exist_ok=True)
    shortcut_script(common/(DISPLAY_NAME+'.lnk'), install_dir/MAIN_EXE, install_dir, install_dir/ICON)
    shortcut_script(common/(S(r'\u5378\u8f7d ')+DISPLAY_NAME+'.lnk'), install_dir/INSTALLED_UNINSTALL_EXE, install_dir, install_dir/ICON)
    if desktop:
        desk=Path(os.environ.get('PUBLIC',str(Path.home())))/'Desktop'
        try: desk.mkdir(parents=True,exist_ok=True)
        except Exception: desk=Path(os.environ.get('USERPROFILE',str(Path.home())))/'Desktop'
        shortcut_script(desk/(DISPLAY_NAME+'.lnk'), install_dir/MAIN_EXE, install_dir, install_dir/ICON)
class Installer(tk.Tk):
    def __init__(self):
        super().__init__(); self.title(DISPLAY_NAME+' v22.1.21 '+S(r'\u5b89\u88c5\u7a0b\u5e8f')); self.geometry('760x520'); self.minsize(560,420); self.configure(bg='#F6F8FC')
        try: self.iconbitmap(str(resource_path(ICON)))
        except Exception: pass
        self.install_dir=tk.StringVar(value=str(find_install_dir())); self.desktop=tk.BooleanVar(value=True); self.status=tk.StringVar(value=S(r'\u51c6\u5907\u5b89\u88c5')); self.detail=tk.StringVar(value=S(r'\u5b89\u88c5\u7a0b\u5e8f\u4f1a\u7ee7\u627f\u65e7\u7248\u8def\u5f84\uff0c\u5e76\u5728\u63a7\u5236\u9762\u677f\u4e2d\u5347\u7ea7\u4e3a v22.1.21\u3002')); self.build()
    def build(self):
        self.grid_columnconfigure(0,weight=1); self.grid_rowconfigure(1,weight=1)
        head=tk.Frame(self,bg='#F6F8FC'); head.grid(row=0,column=0,sticky='ew',padx=22,pady=(18,8))
        tk.Label(head,text=DISPLAY_NAME+' v22.1.21',bg='#F6F8FC',fg='#0F172A',font=('Microsoft YaHei UI',17,'bold')).pack(anchor='w')
        tk.Label(head,text=S(r'\u5347\u7ea7\u5b89\u88c5\u5305\uff1a\u81ea\u52a8\u8bfb\u53d6\u65e7\u7248\u5b89\u88c5\u8def\u5f84\uff0c\u63a7\u5236\u9762\u677f/\u8bbe\u7f6e\u4e2d\u53ea\u4fdd\u7559\u540c\u4e00\u4e2a\u8f6f\u4ef6\u6761\u76ee\u3002'),bg='#F6F8FC',fg='#64748B',wraplength=700,justify='left').pack(anchor='w',pady=(4,0))
        card=tk.Frame(self,bg='white',highlightbackground='#DCE3EF',highlightthickness=1); card.grid(row=1,column=0,sticky='nsew',padx=22,pady=8); card.grid_columnconfigure(0,weight=1)
        tk.Label(card,text=S(r'\u5b89\u88c5\u8def\u5f84'),bg='white',fg='#0F172A',font=('Microsoft YaHei UI',11,'bold')).grid(row=0,column=0,sticky='w',padx=16,pady=(16,6))
        row=tk.Frame(card,bg='white'); row.grid(row=1,column=0,sticky='ew',padx=16); row.grid_columnconfigure(0,weight=1)
        tk.Entry(row,textvariable=self.install_dir,relief=tk.SOLID,bd=1,font=('Microsoft YaHei UI',10)).grid(row=0,column=0,sticky='ew',ipady=6); ttk.Button(row,text=S(r'\u6d4f\u89c8'),command=self.browse).grid(row=0,column=1,padx=(8,0))
        tk.Checkbutton(card,text=S(r'\u521b\u5efa\u684c\u9762\u5feb\u6377\u65b9\u5f0f'),variable=self.desktop,bg='white',activebackground='white').grid(row=2,column=0,sticky='w',padx=16,pady=12)
        tk.Label(card,textvariable=self.detail,bg='white',fg='#64748B',wraplength=680,justify='left').grid(row=3,column=0,sticky='ew',padx=16,pady=(0,12))
        bottom=tk.Frame(self,bg='#F6F8FC'); bottom.grid(row=2,column=0,sticky='ew',padx=22,pady=(4,14)); bottom.grid_columnconfigure(0,weight=1)
        self.progress=ttk.Progressbar(bottom,mode='determinate',maximum=100); self.progress.grid(row=0,column=0,columnspan=3,sticky='ew',pady=(0,6))
        tk.Label(bottom,textvariable=self.status,bg='#F6F8FC',fg='#2563EB',font=('Microsoft YaHei UI',10,'bold')).grid(row=1,column=0,sticky='w'); ttk.Button(bottom,text=S(r'\u53d6\u6d88'),command=self.destroy).grid(row=1,column=1,padx=8); ttk.Button(bottom,text=S(r'\u5b89\u88c5 / \u5347\u7ea7'),command=self.install).grid(row=1,column=2)
    def browse(self):
        chosen=filedialog.askdirectory(initialdir=str(Path(self.install_dir.get()).parent),title=S(r'\u9009\u62e9\u5b89\u88c5\u76ee\u5f55'))
        if chosen: self.install_dir.set(str(Path(chosen)))
    def step(self,i,total,status,detail=''):
        self.progress['value']=int(i/total*100); self.status.set(status); self.detail.set(detail); self.update_idletasks()
    def copy(self,src,dst,i,total):
        self.step(i,total,S(r'\u6b63\u5728\u590d\u5236\uff1a')+dst.name,str(src)+'  ->  '+str(dst)); dst.parent.mkdir(parents=True,exist_ok=True); shutil.copy2(src,dst)
    def install(self):
        if not is_admin(): messagebox.showwarning(S(r'\u6743\u9650\u4e0d\u8db3'),S(r'\u8bf7\u4ee5\u7ba1\u7406\u5458\u6743\u9650\u8fd0\u884c\u5b89\u88c5\u5305\uff0c\u5426\u5219\u65e0\u6cd5\u5347\u7ea7\u63a7\u5236\u9762\u677f\u4e2d\u7684\u65e7\u7248\u5378\u8f7d\u9879\u3002')); return
        try:
            target=Path(self.install_dir.get()).expanduser().resolve(); total=9; self.step(1,total,S(r'\u6b63\u5728\u51c6\u5907\u5b89\u88c5\u76ee\u5f55'),str(target)); target.mkdir(parents=True,exist_ok=True)
            self.copy(resource_path(MAIN_EXE),target/MAIN_EXE,2,total); self.copy(resource_path(PAYLOAD_UNINSTALL_EXE),target/INSTALLED_UNINSTALL_EXE,3,total); self.copy(resource_path(CHANGELOG),target/CHANGELOG,4,total); self.copy(resource_path(ICON),target/ICON,5,total)
            self.step(6,total,S(r'\u6b63\u5728\u6e05\u7406\u65e7\u7248\u5feb\u6377\u65b9\u5f0f'),S(r'\u5f00\u59cb\u83dc\u5355\u4e0e\u684c\u9762\u65e7\u5165\u53e3')); remove_old_shortcuts(); self.step(7,total,S(r'\u6b63\u5728\u521b\u5efa\u5f00\u59cb\u83dc\u5355\u5165\u53e3'),DISPLAY_NAME); create_shortcuts(target,self.desktop.get()); self.step(8,total,S(r'\u6b63\u5728\u5199\u5165 Windows \u5378\u8f7d\u4fe1\u606f'),S(r'\u6ce8\u518c\u8868\u9879\uff1a')+APP_ID+S(r'\uff0c\u7248\u672c\uff1a22.1.21')); write_uninstall_entry(target)
            self.step(9,total,S(r'\u5b89\u88c5\u5b8c\u6210'),S(r'\u53ef\u4ee5\u5728\u63a7\u5236\u9762\u677f\u3001Windows 11 \u8bbe\u7f6e\u548c\u5f00\u59cb\u83dc\u5355\u4e2d\u627e\u5230\u8be5\u8f6f\u4ef6\u3002')); messagebox.showinfo(S(r'\u5b89\u88c5\u5b8c\u6210'),DISPLAY_NAME+' v22.1.21 '+S(r'\u5df2\u5b89\u88c5/\u5347\u7ea7\u5b8c\u6210\u3002\n\n')+str(target)); self.destroy()
        except Exception as e: self.status.set(S(r'\u5b89\u88c5\u5931\u8d25')); self.detail.set(str(e)); messagebox.showerror(S(r'\u5b89\u88c5\u5931\u8d25'),str(e))
if __name__=='__main__': Installer().mainloop()


