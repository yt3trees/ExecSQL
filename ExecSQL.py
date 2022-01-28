from logging import disable
import sys
import os
from turtle import width
import pyodbc
import pyperclip as pc
import tkinter as tk
from tkinter import ttk
from tkinter import font
from tkinter import messagebox
from tkinter import filedialog
import json
import pprint
import subprocess
import win32gui # pip install pywin32
import ctypes
import chardet
import re
import textwrap
import csv
import datetime

VERSION = 'v1.0.0'

# 高DPI対応
import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(True)
except:
    pass

# scriptPath = os.getcwd()
scriptPath = os.path.dirname(str(sys.argv[0]))

PARAM = scriptPath + "\ExecSQL.config.json"
print ('PARAM:',PARAM)

ARGV = sys.argv
# ARGV.append('C:\\work\\dev\\python\\テスト\\testSelect.sql') #デバッグ用
# ARGV.append('C:\\work\\dev\\python\\テスト\\testError.sql') #デバッグ用
# ARGV.append('C:\\work\\dev\\python\\テスト\\test4.sql') #デバッグ用
# ARGV.append('C:\\work\\dev\\python\\テスト\\testSelect2.sql') #デバッグ用
# ARGV.append('C:\\work\\dev\\python\\テスト\\testSelect3.sql') #デバッグ用
# ARGV.append('C:\\work\\dev\\python\\テスト\\test2.sql') #デバッグ用
# ARGV.append('C:\\work\\dev\\python\\テスト\\testViewDelCreate.sql') #デバッグ用
if len(ARGV) <= 1:
    messagebox.showerror('Error!', 'ファイルが指定されていません')
    sys.exit() # 引数がない場合は終了
ARGV.pop(0) # sys.argv[0](自分自身)を配列から除外
print('ARGV:', ARGV)

# .sqlファイル以外が含まれていた場合は終了
for f in ARGV:
    if os.path.splitext(f)[1] != '.sql':
        messagebox.showerror('Error!', '.sqlファイルのみを選択してください' + f)
        sys.exit()

class Application(tk.Frame):
    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        # self.pack(expand=1, fill=tk.BOTH, anchor=tk.NW)
        # self.master.overrideredirect(True)
        self.font = font.Font(self,family='Microsoft Sans Serif', size=11)
        self.fontSub = font.Font(self,family='Microsoft Sans Serif', size=10)
        ww = self.master.winfo_screenwidth()
        wh = self.master.winfo_screenheight()
        self.master.title('ExecSQL')

        # jsonファイルの読み込み
        jsonValues = self.load_json()
        # jsonの値でテーマを変更
        if jsonValues.get('Theme') == 'dark':
            self.bgColor = '#373739'
            self.fgColor = '#F0EEEC'
            self.grayColor = '#4a4a4c'
        elif jsonValues.get('Theme') == 'light':
            self.bgColor = '#e2e2e2'
            self.fgColor = '#373739'
            self.grayColor = '#F0EEEC'
        else:
            message = 'Themeの指定は"dark"か"light"を指定してください\n' + PARAM + 'を確認してください'
            messagebox.showerror('Error!', message)
            sys.exit()
        self.jsonServer = jsonValues.get('Server').split('*') # 区切り文字で配列に格納
        # jsonDatabase = jsonValues.get('Database')
        jsonUser = jsonValues.get('User')
        jsonPassword = jsonValues.get('Password')
        self.jsonOnePoint = jsonValues.get('ColorOnePoint')
        self.jsonDriver = jsonValues.get('Driver')

        self.master.configure(bg=self.bgColor) # 背景色

        # ラベル定義
        labelServer = tk.Label(self.master, text = 'Server', anchor=tk.CENTER, width=10, font=self.font, bg=self.bgColor,fg=self.fgColor)
        labelServer.grid(row = 0, column = 0)
        labelDatabase = tk.Label(self.master, text = 'Database', anchor=tk.CENTER, width=10, font=self.font, bg=self.bgColor,fg=self.fgColor)
        labelDatabase.grid(row = 1, column = 0)
        labelUser = tk.Label(self.master, text = 'User', anchor=tk.CENTER, width=10, font=self.font, bg=self.bgColor,fg=self.fgColor)
        labelUser.grid(row = 2, column = 0)
        labelPassword = tk.Label(self.master, text = 'Password', anchor=tk.CENTER, width=10, font=self.font, bg=self.bgColor,fg=self.fgColor)
        labelPassword.grid(row = 3, column = 0)

        # スタイル定義(コンボボックス用)
        style = ttk.Style()
        style.theme_create('combostyle', parent='alt',
                         settings = {'TCombobox':
                                     {'configure':
                                      {'selectbackground': self.grayColor if jsonValues.get('Theme') == 'dark' else self.fgColor, # テーマがdarkならgrayColor、lightならfgColor
                                       'fieldbackground': self.grayColor,
                                       'background': self.grayColor,
                                       'foreground': self.fgColor,
                                       'relief': 'flat',
                                       }}}
                         )
        style.theme_use('combostyle')
        self.jsonServer.reverse()

        # テキストボックス定義
        self.entServer = ttk.Combobox(self.master, width = 39, font=self.font, style='TCombobox', values=self.jsonServer)
        self.entServer.current(0)
        # self.entServer.bind('<<ComboboxSelected>>', lambda e:self.get_database('update'))
        self.entServer.grid(row = 0, column = 1, sticky="W", columnspan=2)
        self.entUser = tk.Entry(self.master, width = 40, font=self.font, bg=self.grayColor, fg=self.fgColor, relief='flat',disabledbackground=self.bgColor, disabledforeground=self.grayColor)
        self.entUser.insert(0, jsonUser)
        self.entUser.grid(row = 2, column = 1, sticky="W", columnspan=2)
        self.entPassword = tk.Entry(self.master, width = 30, font=self.font, bg=self.grayColor, fg=self.fgColor, show='*', relief='flat',disabledbackground=self.bgColor, disabledforeground=self.grayColor)
        self.entPassword.insert(0, jsonPassword)
        self.entPassword.grid(row = 3, column = 1, sticky="W")
        self.get_database('')
        self.entDatabase = ttk.Combobox(self.master, width = 39, font=self.font, style='TCombobox', values='master', state='readonly')
        self.entDatabase.current(0)
        self.entDatabase.bind('<Button-1>', lambda e:self.get_database('update')) # データベースの初期値はmaster固定でクリック時に取得する
        self.entDatabase.grid(row = 1, column = 1, sticky="W", columnspan=2)

        # ボタン定義
        buttExec = tk.Button(self.master, text="Run", width=10, height=2, bg=self.bgColor, fg=self.jsonOnePoint, activebackground=self.grayColor, relief='raised', font=self.font)#5EA15F
        buttExec.bind("<Button-1>", lambda e: self.after(1, self.push_exec_query))
        buttExec.grid(row=0, column=3, rowspan=2, sticky='w')
        buttClose = tk.Button(self.master, text="Close", width=10, height=2, font=self.font, bg=self.grayColor, fg=self.fgColor, relief='raised')
        buttClose.bind("<Button-1>", lambda e: master.destroy())
        buttClose.grid(row=2, column=3, rowspan=2, sticky='w')

        # チェックボックス定義
        self.winauth = tk.BooleanVar()
        self.winauth.set(False)
        def disable_enable_entry():
            '''チェック時にユーザー名とパスワードを無効化する'''
            if self.winauth.get() == True:
                self.entUser.config(state='disabled')
                self.entPassword.config(state='disabled')
            else:
                self.entUser.config(state='normal')
                self.entPassword.config(state='normal')
        checkWinauth = tk.Checkbutton(self.master, text = 'Win Auth', variable = self.winauth, bg=self.bgColor, fg=self.fgColor,\
             activebackground=self.bgColor,activeforeground=self.fgColor, relief='flat', font=self.font, selectcolor=self.bgColor,\
             command=disable_enable_entry)
        checkWinauth.grid(row = 3, column = 2, sticky="W", padx=5)
        def enter_sc_check(event):
            event.widget['selectcolor'] = self.grayColor
        def leave_sc_check(event):
            event.widget['selectcolor'] = self.bgColor
        checkWinauth.bind("<Enter>", enter_sc_check) # マウスカーソルが重なったら
        checkWinauth.bind("<Leave>", leave_sc_check) # マウスカーソルが離れたら

        # ファイルリスト
        i:int = 0
        var = []
        variable = []
        fileTmp = []
        for file in ARGV:
            if len(file) > 50:
                fileAfter = '...\\' + os.path.basename(os.path.dirname(file)) + '\\' + os.path.basename(file)
            else:
                fileAfter = file
            fileAfter = str(i + 1) + ': ' + fileAfter
            fileTmp.append(file)
            var.insert(i, file)
            variable.insert(i,tk.Label(self.master, text = fileAfter, width=51, anchor='w', font=self.fontSub, bg=self.bgColor, fg=self.fgColor))
            variable[i].grid(row = i+4, column = 0, columnspan=3, sticky="W", padx=5)
            i = i + 1
        def callback(event):
            idx = event.widget['text'][0:1]
            print(fileTmp[int(idx)-1])
            hWnd = win32gui.FindWindow(None,os.path.basename(os.path.dirname(fileTmp[int(idx)-1])))
            title = win32gui.GetWindowText(hWnd)
            if title == os.path.basename(os.path.dirname(fileTmp[int(idx)-1])): # 既にウィンドウが存在している場合はそのウィンドウをアクティブにする
                if win32gui.IsIconic(hWnd):
                    win32gui.ShowWindow(hWnd,1)
                ctypes.windll.user32.SetForegroundWindow(hWnd)
            else:
                subprocess.Popen(['explorer',os.path.dirname(event.widget['text'])]) # ウィンドウが存在しない場合は新規ウィンドウを開く

        # ファイルリストラベルのbind定義
        def enter_fg(event):
            event.widget['fg'] = self.jsonOnePoint
        def leave_fg(event):
            event.widget['fg'] = self.fgColor
        s:int = 0
        for r in ARGV:
            variable[s].bind("<Button-1>", callback) # クリックでファイルを開く
            variable[s].bind("<Enter>", enter_fg) # マウスカーソルが重なったら
            variable[s].bind("<Leave>", leave_fg) # マウスカーソルが離れたら
            s += 1

        # ショートカット定義
        self.master.bind("<Escape>", lambda e: master.destroy()) # 閉じる
        self.master.bind("<F5>", lambda e: self.push_exec_query()) # 実行
        # F1でバージョン
        self.master.bind("<F1>", lambda e: messagebox.showinfo('Version', 'Version: ' + VERSION))

        self.master.update_idletasks()
        lw = self.master.winfo_width()
        lh = self.master.winfo_height()
        self.master.geometry("+"+str(int(ww/2-lw/2))+"+"+str(int(wh/2.5-lh/2.5))) # 画面中央に表示

    def push_exec_query(self):
        '''
        exec_queryを実行する
        after関数を使用することでボタンが押しっぱなしになるのを防ぐ
        '''
        self.exec_query()

    def exec_query(self):
        '''
        SQLを実行する
        '''
        for file in ARGV:
            if os.path.isfile(file) == True:
                pass
            else:
                messagebox.showerror('Error!', 'ファイルが見つかりません')
                return
        try:
            server = self.entServer.get()
            database = self.entDatabase.get()
            user = self.entUser.get()
            password = self.entPassword.get()
            if self.winauth.get() == True:
                trusted_connection = 'yes' # Windows認証
            elif self.winauth.get() == False:
                trusted_connection = 'no'

            connect= pyodbc.connect('DRIVER={'+ self.jsonDriver +'}'+#SQL Server or ODBC Driver 17 for SQL Server
                                    ';SERVER=' + server +
                                    ';DATABASE=' + database +
                                    ';uid=' + user +
                                    ';pwd=' + password +
                                    ';Trusted_Connection=' + trusted_connection) # Windows認証
            connect.setencoding('shift-jis')

            # ファイルごとにSQLを実行する
            i:int = 4
            errorCount:int = 0
            RESULT = []
            COLUMN = []
            nonSlct = 0
            n = 0
            for file in ARGV:
                try:
                    print('************************************************')
                    print('\n★実行対象:', file)

                    # file  = open(file, 'r', encoding=encode['encoding'])
                    with open(file, 'rb') as f:
                        b = f.read()
                    encode = chardet.detect(b)
                    if encode['encoding'] == 'Windows-1252':
                        query = b.decode('shift-jis')
                    else:
                        query = b.decode(encode['encoding'])

                    # /* */で囲まれたコメントを検索
                    findComments = re.compile(r'/\*.*?\*/', flags=re.DOTALL)

                    findSelect = re.compile(r'^\s*SELECT\s*?' # SELECTのみの行を検索
                                        , re.IGNORECASE # 大文字・小文字を区別しない
                                        | re.MULTILINE) # 複数行にマッチさせる
                    selectCheck = [comment for comment in findComments.findall(query) if findSelect.search(comment)] # SELECTを含むコメントを検索
                    # /* */で囲まれたSELECTをコメントアウト
                    for comment in selectCheck:
                        query = query.replace(comment, re.sub(findSelect, '-- SELECT', comment))
                    selectCheck = []
                    selectCheck += [slct for slct in findSelect.findall(query)] # 正規表現にマッチする文字列をリストに格納
                    print ('SELECTを', len(selectCheck), '個見つけました')

                    findCreate = re.compile(r'^\s*CREATE\s*?' # CREATEのみの行を検索
                                        , re.IGNORECASE # 大文字・小文字を区別しない
                                        | re.MULTILINE) # 複数行にマッチさせる
                    createCheck = [comment for comment in findComments.findall(query) if findCreate.search(comment)] # CREATEを含むコメントを検索
                    # /* */で囲まれたCREATEをコメントアウト
                    for comment in createCheck:
                        query = query.replace(comment, re.sub(findCreate, '-- CREATE', comment))
                    createCheck = []
                    createCheck += [crt for crt in findCreate.findall(query)] # 正規表現にマッチする文字列をリストに格納
                    print ('CREATEを', len(createCheck), '個見つけました')

                    findUpdate = re.compile(r'^\s*UPDATE\s*?' # UPDATEのみの行を検索
                                        , re.IGNORECASE # 大文字・小文字を区別しない
                                        | re.MULTILINE) # 複数行にマッチさせる
                    updateCheck = [comment for comment in findComments.findall(query) if findUpdate.search(comment)] # UPDATEを含むコメントを検索
                    # /* */で囲まれたUPDATEをコメントアウト
                    for comment in updateCheck:
                        query = query.replace(comment, re.sub(findUpdate, '-- UPDATE', comment))
                    updateCheck = []
                    updateCheck += [updt for updt in findUpdate.findall(query)] # 正規表現にマッチする文字列をリストに格納
                    print ('UPDATEを', len(updateCheck), '個見つけました')

                    # クエリからGOコマンドを削除する
                    findGo = re.compile(r'^\s*GO\s*$' # GOのみの行を検索
                                        , re.IGNORECASE # 大文字・小文字を区別しない
                                        | re.MULTILINE) # 複数行にマッチさせる
                    goCheck = [comment for comment in findComments.findall(query) if findGo.search(comment)] # GOコマンドを含むコメントを検索
                    # /* */で囲まれたGOをコメントアウト
                    for comment in goCheck:
                        query = query.replace(comment, re.sub(findGo, '-- GO', comment))
                    print ('GOコマンドを', len(goCheck), '個見つけました')
                    # GOコマンドを分割してリストに格納
                    query = [part for part in findGo.split(query) if part]

                    # SELECT以外のクエリが含まれる場合はSELECT結果を表示しない
                    if len(createCheck) > 0 or len(updateCheck) > 0:
                        selectCheck = []

                    # # クエリからUSEコマンドを削除する ※pyodbcではUSEコマンドを含めるとSELECTできないため
                    # findUse = re.compile(r'(?!\/\*.*\n*).*^\s*USE\s*.*(?!\n*.*?\*\/)' # USEで始まる行を検索(/* */で囲まれている行は除く)
                    #                     , re.IGNORECASE # 大文字・小文字を区別しない
                    #                     | re.MULTILINE) # 複数行にマッチさせる
                    # useCheck = [use for use in findUse.findall(query)] # 正規表現にマッチする文字列をリストに格納
                    # print('ゆーざう', useCheck)
                    # print ('USEコマンドを', len(useCheck), '個削除しました')
                    # # USEが存在する場合は接続先データベースを変更する
                    # if len(useCheck) == 1:
                    #     # connect.close()
                    #     # database = useCheck[0].split()[1]
                    #     # database = database.replace('[', '')
                    #     # database = database.replace(']', '')
                    #     # print('USEコマンドが存在するため接続先データベースを', database, 'に変更しました')
                    #     # connect= pyodbc.connect('DRIVER={'+ self.jsonDriver +'}'+
                    #     #             ';SERVER=' + server +
                    #     #             ';DATABASE=' + database +
                    #     #             ';uid=' + user +
                    #     #             ';pwd=' + password +
                    #     #             ';Trusted_Connection=' + trusted_connection) # Windows認証
                    #     #     connect.setencoding('shift-jis')
                    #     # if len(useCheck) > 1:
                    #     #     msg = 'USEコマンドが複数存在します。1つのみにしてください。\n' + file
                    #     #     messagebox.showerror('Error!', msg)
                    #     #     return
                    #     # for us in useCheck:
                    #     #     us = '-- ' + us
                    #     #     query = re.sub(findUse, us, query) # USEをコメントアウト

                    print ('\n-----After replacement query FROM-----', sep = '')
                    print('分割数:', len(query))
                    for q in query:
                        print('↓-----↓')
                        print(q)
                    print ('\n-----After replacement query TO-----\n', sep = '')

                        # query = file.read()
                        # query= 'SELECT * FROM testTable'
                    # file.close()
                    cursor = connect.cursor()
#                     query = 'CREATE VIEW [dbo].[testView]\
# AS\n\
# SELECT                         dbo.testTable.*\n\
# FROM                            dbo.testTable'
                    for command in query:
                        exe = cursor.execute(command)
                    # exe = cursor.execute(query)
                    if len(selectCheck) > 0:
                        rows = cursor.fetchall() #TASK:SELECT対応
                    else:
                        connect.commit()
                        nonSlct += 1

                    if len(selectCheck) > 0:
                        n = i
                        n = n - nonSlct

                    if len(selectCheck) > 0:
                        # 取得データ整形・格納
                        result = []
                        for r in rows:
                            txt = []
                            for d in range(len(r)):
                                # NULLのデータは実行しない
                                if r[d] is None:
                                    txt.append('')
                                else:
                                    txt.append(str.strip(str(r[d])))
                            result.append(txt)
                            # result.append((str.strip(r[0]), str.strip(r[1])))
                        # pprint.pprint( list(rows) )
                        # print(list(rows))
                        RESULT.insert(n-4, result)
                        print('結果:',result)

                        #カラム名取得
                        columns = []
                        for column in exe.description:
                            columns.append(column[0])
                        COLUMN.insert(n-4, columns)
                        print('カラム:', columns)

                        def sub_window(rows, columns):
                            # print('ボタン押下columns',columns)
                            # print('ボタン押下result',rows)
                            def save_csv():
                                dtime =  datetime.date.today().strftime('%Y%m%d')
                                filename = dtime + '_' + os.path.splitext(os.path.basename(file))[0] + '.csv'
                                ftype = [('CSV File', '.csv')]
                                fname = filedialog.asksaveasfilename(initialfile=filename, filetypes=ftype)
                                if fname:
                                    with open(fname, 'w', newline='') as f:
                                        data = []
                                        data.append(columns)
                                        for row in rows:
                                            data.append(row)
                                        writer = csv.writer(f)
                                        writer.writerows(data)
                                    print('保存しました')
                                    messagebox.showinfo('Complete', '保存しました')
                                else:
                                    print('キャンセルしました')

                            sub_win = tk.Toplevel()
                            sub_win.title('Result')
                            ww = sub_win.winfo_screenwidth()
                            wh = sub_win.winfo_screenheight()
                            style = ttk.Style()
                            style.configure('Treeview.Heading', font = self.font, foreground=self.fgColor, background=self.bgColor, relief=tk.RAISED)
                            style.configure('Treeview', font = self.font, rowheight=30, foreground=self.fgColor, background=self.bgColor,fieldbackground=self.bgColor, relief=tk.FLAT)
                            sub_win.configure(bg=self.bgColor)

                            # データ件数が30件以上の場合、スクロールバーを表示し30件の縦幅にする
                            if len(rows) >= 30:
                                rowInt = 30
                                scrollbar = tk.Scrollbar(sub_win, orient=tk.VERTICAL)
                                scrollbar.set(0.2,0.3)
                                # scrollbar.grid(row=0, column=2, sticky=tk.N+tk.S)
                                scrollbar.place(relheight=1.0,relwidth=0.02,relx=0.98,rely=0.0)
                                # tree = ttk.Treeview(sub_win, columns=columns, height=30, yscrollcommand=scrollbar.set)
                                tree = ttk.Treeview(sub_win, columns=columns, height=rowInt, yscrollcommand=scrollbar.set)
                                scrollbar.config(command=tree.yview)
                            else:
                                rowInt = len(rows)
                                tree = ttk.Treeview(sub_win, columns=columns, height=rowInt)

                            # 横スクロールバーは列数に関係なく表示
                            xscrollbar = tk.Scrollbar(sub_win, orient=tk.HORIZONTAL)
                            xscrollbar.place(relheight=0.07, relwidth=1.0, relx=0.0, rely=0.83)
                            xscrollbar.config(command=tree.xview)
                            tree.configure(xscrollcommand=xscrollbar.set)

                            # タグ定義(色設定用)
                            tree.tag_configure('color', background=self.bgColor , foreground=self.fgColor)
                            tree.tag_configure('color2', background=self.grayColor , foreground=self.fgColor)

                            # カラム定義
                            tree.column('#0',width=0, stretch='no')
                            for c in columns:
                                tree.column(c, anchor='center')
                            tree.heading('#0',text='')
                            for c in columns:
                                tree.heading(c, text=c,anchor='center', command=lambda c=c: sortby(tree, c, 0))

                            def sortby(tree, col, descending):
                                l = [(tree.set(k, col), k) for k in tree.get_children('')]
                                l.sort(reverse=descending)
                                # アイテムの並び替え
                                for index, (val, k) in enumerate(l):
                                    tree.move(k, '', index)
                                tree.heading(col, command=lambda col=col: sortby(tree, col, int(not descending)))

                            # レコードの追加
                            for i in range(len(rows)):
                                # 1列ごとにセルの背景色を変える
                                if i % 2 == 0:
                                    tree.insert('', 'end', values=rows[i], tags='color2')
                                else:
                                    tree.insert('', 'end', values=rows[i], tags='color')

                            # tree.grid(row=0, column=0, columnspan=2, sticky='nsew')
                            tree.place(relheight=1.0,relwidth=0.98,relx=0.0,rely=0.0)
                            tree.columnconfigure(0, weight=1)
                            tree.rowconfigure(0, weight=1)
                            tree.grid_propagate(False)

                            bt = tk.Button(sub_win, text='Close', command=sub_win.destroy, width=10, height=1\
                                            , bg=self.grayColor, fg=self.fgColor, activebackground=self.grayColor, relief='raised', font=self.font)
                            # bt.grid(row=2, column=1, sticky='nsew')
                            bt.place(relheight=0.10,relwidth=0.50,relx=0.5,rely=0.9)
                            csvButt = tk.Button(sub_win, text='Save as CSV', command=save_csv, width=10, height=1\
                                            , bg=self.bgColor, fg=self.fgColor, activebackground=self.grayColor, relief='raised', font=self.font)
                            # csvButt.grid(row=2, column=0, sticky='nsew')
                            csvButt.place(relheight=0.10,relwidth=0.50,relx=0.0,rely=0.9)

                            sub_win.update_idletasks()
                            # lw = sub_win.winfo_width()
                            # lh = sub_win.winfo_height()
                            lw = 1500
                            lh = 400
                            sub_win.geometry(str(lw)+'x'+str(lh)+"+"+str(int(ww/2-lw/2))+"+"+str(int(wh/2.5-lh/2.5))) # 画面中央に表示

                    if len(selectCheck) == 0:
                        tk.Label(self.master, text = 'Success', width=10, anchor='w', font=self.fontSub, bg=self.bgColor,fg=self.fgColor)\
                        .grid(row = i, column = 3, padx=5, sticky="W")
                    else:
                        def callback(event):
                            d = str(event.widget['text'])[-1:]
                            print('\n------------テーブル表示ボタン押下------------', event.widget['text'])
                            sub_window(RESULT[int(d)], COLUMN[int(d)])
                        def enter_fg(event):
                            event.widget['fg'] = self.jsonOnePoint
                        def leave_fg(event):
                            event.widget['fg'] = self.fgColor
                        print('処理中行i', n)
                        variableA = []
                        variableA.insert(n, tk.Label(self.master, text = 'Show Results     ' + ' ' + str(n-4), padx=8, width=10, anchor='w'\
                                                        , font=self.fontSub, bg=self.grayColor,fg=self.fgColor, relief='ridge'))
                        variableA[0].grid(row = i, column = 3,  sticky="w")
                        variableA[0].bind("<Button-1>", callback)
                        variableA[0].bind("<Enter>", enter_fg) # マウスカーソルが重なったら
                        variableA[0].bind("<Leave>", leave_fg) # マウスカーソルが離れたら

                    i += 1
                    cursor.close()
                except Exception as e:
                    print(e)
                    errorCount += 1
                    RESULT.insert(i-4, ['No Data'])
                    COLUMN.insert(i-4, ['No Data'])
                    nonSlct += 1
                    # if ''.join(e.args).startswith('No results.  Previous SQL was not a query.'):#TASK:SELECT対応
                    #     tk.Label(self.master, text = 'Success', width=10, anchor='w', font=self.fontSub, bg=self.bgColor,fg=self.fgColor).grid(row = i, column = 3,  sticky="W")#9fa05e
                    #     i += 1
                    # else:
                    # エラーメッセージ表示 一定の文字数を超えたら改行を入れる
                    res = textwrap.wrap(str(e), width=65)
                    res = ('\n'.join(res))
                    tk.Label(self.master, text = res, anchor='w', justify='left', font=self.fontSub, bg=self.bgColor,fg='#a05e5e').grid(row = i, column = 3,  sticky="W")
                    i += 1

            connect.close()

            # サーバコンボボックスの履歴制御(5件まで)
            if (self.entServer.get() in self.jsonServer) == False: # サーバ履歴に同じ値が無ければ追加
                self.jsonServer.pop(4) # 一番最後を削除
                self.jsonServer.insert(0,self.entServer.get())
                self.entServer['values'] = self.jsonServer # コンボボックスを更新
            else: # サーバ履歴に同じ値があれば入れ替え
                selectIndex = self.jsonServer.index(self.entServer.get())
                self.jsonServer.pop(selectIndex)
                self.jsonServer.insert(0,self.entServer.get()) # 一旦配列から削除して0番目に追加
                self.entServer['values'] = self.jsonServer # コンボボックスを更新

            # if errorCount == 0:
            #     messagebox.showinfo('Success', '実行完了しました。')
            # else:
            #     messagebox.showinfo('Error', '失敗したクエリがあります。')

            self.save_json()

        except Exception as e:
            print(e)
            messagebox.showerror('Error!', e)

    def save_json(self):
        '''
        jsonファイルに保存
        '''
        # json用文字列調整
        tmpServer = []
        for i in reversed(self.jsonServer): # 配列を逆にして格納
            tmpServer.append(i)
        jServer = '*'.join(tmpServer) # 配列を区切り文字で連結
        # jDatabase = self.entDatabase.get()
        jUser = self.entUser.get()
        jPassword = self.entPassword.get()
        list = {
        'Server': jServer,
            # 'Database': jDatabase,
            'User': jUser,
            'Password': jPassword,
            'ColorOnePoint': self.load_json().get('ColorOnePoint'),
            'Theme': self.load_json().get('Theme'),
            'Driver': self.load_json().get('Driver')
        }
        # jsonファイルに保存
        try:
            if os.path.exists(PARAM):
                with open(PARAM, 'w') as f:
                    json.dump(list, f, indent=2, ensure_ascii=False)
        except Exception as e:
            e = "jsonファイルの書き込みに失敗しました。\n" + str(e)
            print(e)
            messagebox.showerror('Error!', e)

    def load_json(self):
        '''
        jsonファイルの読み込み
        '''
        try:
            if os.path.exists(PARAM):
                jsonOpen = open(PARAM,"r")
                jsonLoad = json.load(jsonOpen)
                return jsonLoad
        except Exception as e:
            e = "jsonファイルの読み込みに失敗しました。\n" + str(e)
            print(e)
            messagebox.showerror('Error!', e)

    def get_database(self, param):
        '''
        サーバ指定でデータベースを取得
        '''
        if param == 'update':
            try:
                server = self.entServer.get()
                user = self.entUser.get()
                password = self.entPassword.get()
                if self.winauth.get() == True:
                    trusted_connection = 'yes' # Windows認証
                elif self.winauth.get() == False:
                    trusted_connection = 'no'

                connect= pyodbc.connect('DRIVER={' + self.jsonDriver + '}'+
                                        ';SERVER=' + server +
                                        ';DATABASE=' + 'master' + # masterにはログインできる前提
                                        ';uid=' + user +
                                        ';pwd=' + password +
                                        ';Trusted_Connection=' + trusted_connection) # Windows認証
                connect.setencoding('shift-jis')

                query= 'SELECT name FROM sys.databases' # データベース名取得
                cursor = connect.cursor()
                cursor.execute(query)
                rows = cursor.fetchall()
                cursor.close()

                # 文字列整形
                self.dbList = []
                for r in rows:
                    s = str(r)
                    self.dbList.append(s[2:-4])

                print('ConnectingDB:', self.dbList)

                self.entDatabase['values'] = self.dbList # コンボボックスを更新
                self.entDatabase.current(0) # 初期値を設定
            except Exception as e:
                messagebox.showerror('Error!', e)
        else:
            pass

def main():
    root = tk.Tk()
    root.attributes('-alpha', 0.0) # ウィンドウちらつき対策 1度目の生成を透明化
    app = Application(master=root)
    root.attributes('-alpha', 1.0) # 完全にウィンドウ生成完了したらアクティブ化
    app.mainloop()

if __name__ == '__main__':
    main()