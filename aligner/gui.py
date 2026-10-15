"""Tkinter Windows 桌面工作台。后台线程处理文件，主线程更新界面。"""
from pathlib import Path
import json
import os
import sys
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from .core import load_tables, choose, choose_reference, detect, from_text, align
from .export import export, audit, log_dir

# 桌面层的交互与线程约定
# Tkinter 控件只在主线程创建和修改，工作线程只计算数据。
# 文件导入和导出可能受磁盘速度影响，因此不能阻塞事件循环。
# background 使用 after 将结果传回主线程，不跨线程改 StringVar。
# busy 标记阻止用户重复启动重叠任务，保护当前状态的一致性。
# 任何后台错误都经同一出口返回消息框，主界面保持可用。
# 来源预览、对齐结果、参考顺序分别展示同一组真实数据。
# 界面没有预填匹配数量，也没有读取预制结果来演示功能。
# 模拟样例按钮只负责填入路径，不代替文件读取和运算。
# 表格树只展示编号与审计信息，完整业务列在输出 Excel 中保留。
# 四个统计值直接来自 align 的实际返回值，不另外估算。
# 多候选是已匹配的子集，因此四个统计值不能直接相加。
# 修改输入或重新导入后应失效旧结果，防止误导出上一轮。
# 文件路径供本机使用，不发送给任何远端服务。
# 粘贴文本单独保存在 Table 中，不写入临时上传目录。
# 未匹配策略在开始计算时固定，结果中记录本次采用的策略。
# 导出依据最近一次成功计算的结果，而不是界面上的排序猜测。
# 状态栏提供处理进度与保存反馈，不要求用户查看终端日志。
# 默认窗口适配常见桌面尺寸，最小尺寸保护控制区可见性。
# Treeview 明确限制最小显示行数，确保底部导出控件不被挤走。
# 用户可直接输入保存位置，减少系统保存对话框带来的打断。
# 保存扩展名跟随来源真实格式，避免用户手工填写错扩展名。
# Windows 打包程序使用可执行文件位置寻找附带模拟样例。
# 源码运行使用工程位置寻找样例，不依赖当前 Python 安装目录。
# 图纸数据在后台内存与本地文件之间流转，应用不构造网络请求。
# 长时间耐久运行通过独立脚本复用核心，避免 UI 操作计时混淆。
# 标题区显示本地处理事实，不展示未经验证的效率提升百分比。
# 颜色只作辅助，状态同时以中文文字表达，便于审计和复制。
# 参考候选显示 Excel 原行号，与导出的审计记录保持一致。
# 所有交互入口最终复用 core/export，避免桌面与 CLI 算法分叉。

class App:
    def __init__(self, root):
        # 三个状态对象分别代表参考、来源与最近的计算结果。
        # 尚未导入时使用 None，不能把空表当成成功导入。
        # 公共模拟数据不包含参赛人电话或比赛附件原始业务资料。
        self.root = root
        self.ref = self.source = self.result = None
        self.root.title('序衡 | 本地报表智能排序对齐')
        self.root.geometry('1280x820+30+30')
        self.root.minsize(1100, 740)
        self.root.configure(bg='#eff3f8')
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('.', font=('Microsoft YaHei UI', 11))
        style.configure('TButton', padding=(15, 9))
        style.configure('Accent.TButton', background='#186658', foreground='white')
        style.configure('Treeview', rowheight=34, font=('Microsoft YaHei UI', 11))
        style.configure('Treeview.Heading', font=('Microsoft YaHei UI', 10, 'bold'), padding=8)
        side = tk.Frame(root, bg='#102f37', width=205)
        side.pack(side='left', fill='y')
        side.pack_propagate(False)
        tk.Label(side, text='序 衡', font=('Microsoft YaHei UI', 28, 'bold'), bg='#102f37', fg='white').pack(anchor='w', padx=25, pady=(35, 4))
        tk.Label(side, text='报表智能排序对齐', bg='#102f37', fg='#a1c1c5', font=('Microsoft YaHei UI', 11)).pack(anchor='w', padx=25)
        for n, title, sub in [('01','导入资料','Excel 或文本粘贴'),('02','自动识别','编号列与工作表'),('03','匹配排序','逐行查看匹配依据'),('04','导出成果','原格式表格与审计记录')]:
            tk.Label(side, text=n+'  '+title, bg='#102f37', fg='white', font=('Microsoft YaHei UI', 13, 'bold')).pack(anchor='w', padx=25, pady=(30, 5))
            tk.Label(side, text=sub, bg='#102f37', fg='#a1c1c5', font=('Microsoft YaHei UI', 10)).pack(anchor='w', padx=25)
        tk.Label(side, text='本地处理\n无需账号 · 不上传资料\nv1.0.0', justify='left', bg='#102f37', fg='#a1c1c5', font=('Microsoft YaHei UI', 10)).pack(side='bottom', anchor='w', padx=25, pady=30)
        body = tk.Frame(root, bg='#eff3f8')
        body.pack(fill='both', expand=True, padx=25, pady=22)
        tk.Label(body, text='把不同顺序的报表，对齐到同一个基准', bg='#eff3f8', fg='#142f38', font=('Microsoft YaHei UI', 20, 'bold')).pack(anchor='w')
        tk.Label(body, text='自动识别编号列，保留完整业务行；多候选与未匹配项都有据可查。', bg='#eff3f8', fg='#586a76').pack(anchor='w', pady=(6, 16))
        inputs = tk.Frame(body, bg='white', padx=16, pady=12)
        inputs.pack(fill='x')
        self.paths = []
        for i, label in enumerate(['参考表格', '待排序表格']):
            tk.Label(inputs, text=label, bg='white', width=10, anchor='w').grid(row=i, column=0, pady=5)
            var = tk.StringVar()
            self.paths.append(var)
            ttk.Entry(inputs, textvariable=var).grid(row=i, column=1, sticky='ew', padx=10)
            ttk.Button(inputs, text='选择文件', command=lambda i=i: self.browse(i)).grid(row=i, column=2)
            ttk.Button(inputs, text='粘贴文本', command=lambda i=i: self.paste(i)).grid(row=i, column=3, padx=(8,0))
        inputs.columnconfigure(1, weight=1)
        bar = tk.Frame(body, bg='#eff3f8')
        bar.pack(fill='x', pady=12)
        ttk.Button(bar, text='填入模拟样例', command=self.sample).pack(side='left')
        ttk.Button(bar, text='导入并识别', command=self.load).pack(side='left', padx=8)
        self.policy = tk.StringVar(value='未匹配置尾')
        ttk.Combobox(bar, textvariable=self.policy, values=['未匹配置尾','未匹配保留原位'], width=18, state='readonly').pack(side='left', padx=8)
        ttk.Button(bar, text='开始对齐', style='Accent.TButton', command=self.process).pack(side='right')
        self.info = tk.StringVar(value='等待导入  /  支持 .xls、.xlsx、制表符文本和 CSV 文本')
        tk.Label(body, textvariable=self.info, bg='#e1ebe9', fg='#16584e', anchor='w', padx=12, pady=10, wraplength=950).pack(fill='x')
        stats = tk.Frame(body, bg='#eff3f8')
        stats.pack(fill='x', pady=12)
        self.stats = []
        for title in ['数据行数','匹配成功','未匹配','多候选需核对']:
            panel = tk.Frame(stats, bg='white', padx=14, pady=8)
            panel.pack(side='left', fill='x', expand=True, padx=(0,8))
            value = tk.StringVar(value='—')
            self.stats.append(value)
            tk.Label(panel, text=title, bg='white', fg='#617580', font=('Microsoft YaHei UI',10)).pack(anchor='w')
            tk.Label(panel, textvariable=value, bg='white', fg='#143c43', font=('Microsoft YaHei UI',23,'bold')).pack(anchor='w')
        tabs = ttk.Notebook(body)
        tabs.pack(fill='both', expand=True)
        self.trees = []
        for title in ['待排序原表', '对齐结果', '参考顺序']:
            frame = ttk.Frame(tabs)
            tabs.add(frame, text=title)
            tree = ttk.Treeview(frame, columns=('output','origin','key','status','ref'), show='headings', height=6)
            for c, heading, width in [('output','当前行',80),('origin','原始行',80),('key','识别编号',330),('status','状态',135),('ref','参考行位置',170)]:
                tree.heading(c,text=heading)
                tree.column(c,width=width,anchor='w' if c=='key' else 'center',stretch=c=='key')
            sb = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
            tree.configure(yscrollcommand=sb.set)
            sb.pack(side='right',fill='y')
            tree.pack(fill='both',expand=True)
            tree.tag_configure('unmatched',foreground='#a24b24',background='#fff2e5')
            tree.tag_configure('ambiguous',foreground='#8b6619',background='#fff9db')
            self.trees.append(tree)
        self.tabs = tabs
        bottom = tk.Frame(body,bg='#eff3f8')
        bottom.pack(fill='x',pady=(12,0))
        self.status = tk.StringVar(value='准备就绪。模拟样例含重复编号、部分匹配和未匹配情况。')
        tk.Label(bottom,textvariable=self.status,bg='#eff3f8',fg='#526771',wraplength=435,justify='left',font=('Microsoft YaHei UI',9)).pack(side='left')
        ttk.Button(bottom,text='导出成果',style='Accent.TButton',command=self.save).pack(side='right')
        self.output_path = tk.StringVar(value='output/对齐成果.xlsx')
        ttk.Entry(bottom,textvariable=self.output_path,width=26).pack(side='right',padx=8)
        tk.Label(bottom,text='保存到',bg='#eff3f8',fg='#526771',font=('Microsoft YaHei UI',9)).pack(side='right')
        self.busy = False

    def background(self, task, done):
        # task 返回普通 Python 值，done 承担所有界面更新。
        # 异常文本先复制到局部变量，避免异常作用域清理后丢失。
        # daemon 线程不会在窗口关闭后独占程序退出流程。
        # 成功与失败都必须复位 busy，使用户能够继续操作。
        if self.busy:
            return
        self.busy = True
        self.status.set('正在本地处理，请稍候…')
        def worker():
            try:
                value = task()
                self.root.after(0, lambda: finish(value))
            except Exception as e:
                text = str(e)
                self.root.after(0, lambda: fail(text))
        def finish(value):
            self.busy = False
            done(value)
        def fail(text):
            self.busy = False
            self.status.set('处理未完成：' + text)
            messagebox.showerror('请检查输入',text)
        threading.Thread(target=worker,daemon=True).start()

    def invalidate(self):
        # 清除统计和结果页，而不是保留看似仍然有效的上一轮数字。
        # 来源和参考预览随后由 refresh_inputs 使用新输入刷新。
        # 旧 result 置空后，导出入口会要求重新运行对齐。
        self.result = None
        for v in self.stats:
            v.set('—')
        self.trees[1].delete(*self.trees[1].get_children())

    def browse(self, i):
        p = filedialog.askopenfilename(filetypes=[('Excel 工作簿','*.xlsx *.xls')])
        if p:
            self.paths[i].set(p)

    def sample(self):
        # frozen 是 PyInstaller 的运行标记，源码运行时不存在。
        # 优先使用相对路径，使演示和源码包不暴露本机用户目录。
        # 这里只填路径，导入按钮才真正读取文件。
        base = Path(sys.executable).parent if getattr(sys,'frozen',False) else Path(__file__).resolve().parent.parent
        for i,name in enumerate(['模拟参考.xlsx','模拟待排序.xlsx']):
            self.paths[i].set(str(Path('samples') / name) if Path('samples').is_dir() else str(base / 'samples' / name))
        self.status.set('已填入明确标注的模拟业务数据，请点击“导入并识别”。')

    def paste(self, i):
        # i 为零代表参考，为一代表来源，两个输入可以混合使用。
        # 弹窗先解析并识别，成功后才替换上一份有效输入。
        # 解析失败保留编辑内容，用户能够修正后再次提交。
        win = tk.Toplevel(self.root)
        win.title('粘贴参考表格' if i == 0 else '粘贴待排序表格')
        win.geometry('760x450')
        tk.Label(win,text='直接粘贴 Excel 区域（制表符分列），也支持 CSV 和单列编号。').pack(pady=12)
        text = tk.Text(win,font=('Microsoft YaHei UI',12),wrap='none')
        text.pack(fill='both',expand=True,padx=15,pady=8)
        def apply():
            try:
                table = detect(from_text(text.get('1.0','end')))
                if i == 0:
                    self.ref = table
                else:
                    self.source = table
                self.paths[i].set('已粘贴文本')
                self.invalidate()
                self.refresh_inputs()
                win.destroy()
            except Exception as e:
                messagebox.showerror('无法读取',str(e),parent=win)
        ttk.Button(win,text='使用这份文本',command=apply).pack(pady=12)
        text.focus_set()

    def load(self):
        # 先识别来源，再利用来源编号选择参考工作表。
        # 两份输入都成功后才一起提交给界面，避免半更新。
        # 已粘贴文本不会被错误地解释为本地文件路径。
        paths = [v.get() for v in self.paths]
        def task():
            source = self.source if paths[1] == '已粘贴文本' and self.source else choose(load_tables(paths[1]))
            ref = self.ref if paths[0] == '已粘贴文本' and self.ref else choose_reference(load_tables(paths[0]),source)
            return ref,source
        def done(pair):
            self.ref,self.source = pair
            self.invalidate()
            self.refresh_inputs()
            self.status.set('导入完成。可查看原始顺序和参考顺序，再点击“开始对齐”。')
        self.background(task,done)

    def refresh_inputs(self):
        if self.ref and self.source:
            self.info.set(f'参考：{self.ref.sheet} · 第 {self.ref.col} 列 · 表头第 {self.ref.header} 行    |    待排序：{self.source.sheet} · 第 {self.source.col} 列 · 表头第 {self.source.header} 行')
        for table,tree in [(self.source,self.trees[0]),(self.ref,self.trees[2])]:
            tree.delete(*tree.get_children())
            if table:
                for r in table.records():
                    tree.insert('','end',values=(r,r,table.value(r,table.col),'待处理' if table is self.source else '参考',r if table is self.ref else ''))

    def process(self):
        # 导入与计算分开，让用户有机会先检查编号列和原顺序。
        # 中文下拉选项只在这里映射为稳定的内部策略标识。
        # 处理完成后自动进入结果页，关键结果立即可见。
        if not self.ref or not self.source:
            messagebox.showinfo('先导入资料','请选择两张表并点击“导入并识别”，或分别粘贴文本。')
            return
        policy = 'fixed' if '保留' in self.policy.get() else 'tail'
        started = time.perf_counter()
        def done(result):
            self.result = result
            elapsed = time.perf_counter() - started
            for var,key in zip(self.stats,['total','matched','unmatched','ambiguous']):
                var.set(str(result[key]))
            tree = self.trees[1]
            tree.delete(*tree.get_children())
            for d in result['rows']:
                tag = 'unmatched' if d['status']=='未匹配' else 'ambiguous' if d['status']=='多候选' else ''
                tree.insert('','end',values=(d['output_row'],d['original_row'],d['key'],d['status'],', '.join(map(str,d['reference_rows']))),tags=(tag,))
            self.tabs.select(1)
            self.status.set(f'已完成对齐。多候选按最早参考行排序；原始行号始终保留。')
            audit({'action':'align','reference':self.ref.name,'source':self.source.name,'elapsed_s':round(elapsed,6),**{k:result[k] for k in ['total','matched','unmatched','ambiguous']}})
        self.background(lambda:align(self.ref,self.source,policy),done)

    def save(self):
        # 只有成功计算过的结果才允许导出，防止创建空成果。
        # 目标目录由导出层创建，GUI 不预先写入业务文件。
        # 保存成功仅更新状态栏，便于连续处理下一份报表。
        if not self.result:
            messagebox.showinfo('先运行对齐','请先完成对齐后再导出。')
            return
        suffix = Path(self.source.path).suffix.lower() if self.source.path else '.xlsx'
        p = self.output_path.get().strip()
        if p and Path(p).suffix.lower() != suffix:
            p = str(Path(p).with_suffix(suffix))
            self.output_path.set(p)
        if p:
            def done(output):
                self.status.set('已导出 '+output.name+' 及同名 audit.csv、audit.json；原文件未修改。')
            self.background(lambda:export(self.source,self.result,p,self.ref.path),done)

def run():
    root = tk.Tk()
    App(root)
    root.mainloop()
