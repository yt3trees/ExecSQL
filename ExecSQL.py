from logging import disable
import sys
import os
import pyodbc
import pyperclip as pc
import tkinter as tk
from tkinter import ttk
from tkinter import font
from tkinter import messagebox
import json
import pprint
import subprocess
import win32gui # pip install pywin32
import ctypes
import chardet
import re
import textwrap

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
# ARGV.append('C:\\work\\dev\\python\\テスト\\test4.sql') #デバッグ用
# ARGV.append('C:\\work\\dev\\python\\テスト\\testError.sql') #デバッグ用
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
            self.bgColor = '#e0e0e5'
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
        buttExec.grid(row=0, column=3, rowspan=2, padx=5, sticky='w')
        buttClose = tk.Button(self.master, text="Close", width=10, height=2, font=self.font, bg=self.grayColor, fg=self.fgColor, relief='raised')
        buttClose.bind("<Button-1>", lambda e: master.destroy())
        buttClose.grid(row=2, column=3, rowspan=2, padx=5, sticky='w')

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
        for file in ARGV:
            var.insert(i,file)
            variable.insert(i,tk.Label(self.master, text = file, width=51, anchor='w', font=self.fontSub, bg=self.bgColor, fg=self.fgColor))
            variable[i].grid(row = i+4, column = 0, columnspan=3,  sticky="W")
            i = i + 1
        def callback(event):
            hWnd = win32gui.FindWindow(None,os.path.basename(os.path.dirname(event.widget['text'])))
            title = win32gui.GetWindowText(hWnd)
            if title == os.path.basename(os.path.dirname(event.widget['text'])): # 既にウィンドウが存在している場合はそのウィンドウをアクティブにする
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
            for file in ARGV:
                try:
                    print('実行対象:', file)

                    # file  = open(file, 'r', encoding=encode['encoding'])
                    with open(file, 'rb') as f:
                        b = f.read()
                    encode = chardet.detect(b)
                    if encode['encoding'] == 'Windows-1252':
                        query = b.decode('shift-jis')
                    else:
                        query = b.decode(encode['encoding'])

                    # クエリからGOコマンドを削除する
                    findGo = re.compile(r'(?!\/\*.*\n*).*^\s*GO\s*$.*(?!\n*.*?\*\/)' # GOのみの行を検索(/* */で囲まれている行は除く)
                                        , re.IGNORECASE # 大文字・小文字を区別しない
                                        | re.MULTILINE) # 複数行にマッチさせる
                    goCheck = [go for go in findGo.findall(query)] # 正規表現にマッチする文字列をリストに格納
                    print ('GOコマンドを', len(goCheck), '個削除しました')
                    query = re.sub(findGo, '-- GO', query) # GOをコメントアウト
                    print ('\n-----After replacement query FROM-----\n', query, '\n-----After replacement query TO-----\n', sep = '')

                    # query = file.read()
                    # query= 'SELECT * FROM testTable'
                    # file.close()
                    cursor = connect.cursor()
                    cursor.execute(query)
                    connect.commit()
                    # rows = cursor.fetchall() #TASK:SELECT対応
                    # pprint.pprint( rows )
                    # print(rows)
                    tk.Label(self.master, text = 'Success', width=10, anchor='w', font=self.fontSub, bg=self.bgColor,fg=self.fgColor).grid(row = i, column = 3,  sticky="W")#9fa05e
                    i += 1
                except Exception as e:
                    print(e)
                    errorCount += 1
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

                connect= pyodbc.connect('DRIVER={+' + self.jsonDriver + '}'+
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