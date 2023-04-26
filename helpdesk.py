# -- coding: utf-8 --
from tkinter import * 
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
import tkinter as tk
import sqlite3
import pandas as pd


# db 연결
conn = sqlite3.connect('data.db')
c = conn.cursor()

'''
생성되는 테이블의 사용 용도
column1 = 접수상태
column2 = 부서
column3 = 성명
column4 = 신청내용
column5 = 기타입력
'''
c.execute('''CREATE TABLE IF NOT EXISTS App_DB
            (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            column1 TEXT,
            column2 TEXT,
            column3 TEXT,
            column4 TEXT,
            column5 TEXT,
            column6 TEXT,
            column7 TEXT
            )''')
conn.commit()

# 기본설정
path = '맑은 고딕'
font_time = path, 20, 'bold'
font = path, 11
font_bold = path, 11, 'bold'
font_mini = path, 10
title_info = '접수대장'
geometry = '1800x1000+400+50'
search_color = 'khaki'

# tkinter GUI 생성
root = tk.Tk()
root.title(title_info) 
root.geometry(geometry)
root.state('zoomed')
# root.iconbitmap('logo.ico')
root.columnconfigure(0, weight=1)
root.rowconfigure(2, weight=1)

# Frame 1 - Date / Time 
date_frame = tk.Frame(root)
date_frame.grid(row=0, column=0, padx=280, pady=20, sticky='sw')


# Notebook - TAB 1 / TAB 2 
notebook = ttk.Notebook(root)
notebook.grid()

# TAB 1 =============================================================
frame1 = Frame(notebook)
notebook.add(frame1, text="신청접수")

# TAB 1 - input_frame
input_frame = tk.Frame(frame1)
input_frame.grid(row=0, column=0, padx=10, pady=10, sticky='sw')

lab1 = '부서'
lab2 = '성명'
lab3 = '신청내용'
lab4 = '기타입력'

lab1 = Label(input_frame, text=lab1, font=font_bold)
lab1.grid(row=0, column=0, padx=40, sticky=W)

ent1 = Entry(input_frame, font=font, relief='flat', highlightthickness=1)
ent1.config(highlightbackground='gray')
ent1.grid(row=1, column=0, padx=40)
ent1.focus() # ent1에 프롬프트 설정

lab2 = Label(input_frame, text=lab2, font=font_bold)
lab2.grid(row=0, column=1, padx=30, sticky=W)

ent2 = Entry(input_frame, font=font, relief='flat', highlightthickness=1)
ent2.config(highlightbackground='gray')
ent2.grid(row=1, column=1, padx=30)

lab3 = Label(input_frame, text=lab3, font=font_bold)
lab3.grid(row=0, column=2, padx=30, sticky=W)

ent3 = Entry(input_frame, width=100, font=font, relief='flat', highlightthickness=1)
ent3.config(highlightbackground='gray')
ent3.grid(row=1, column=2, padx=30)

lab4 = Label(input_frame, text=lab4, font=font_bold)
lab4.grid(row=0, column=3, padx=30 ,sticky=W)

ent4 = Entry(input_frame, width=40, font=font, relief='flat', highlightthickness=1)
ent4.config(highlightbackground='gray')
ent4.grid(row=1, column=3, padx=30)

# TAB 1 - Option Frame
option_frame = tk.Frame(frame1)
option_frame.grid(row=1, column=0, padx=10, pady=10, sticky='sw')

def work_data_count():
    def count():
        data_count = len(tree.get_children())
        data_count_label.config(text=f' 신청접수 COUNT : {data_count} 건 ')    
    '''여백 설정을 위한 라벨 추가'''
    temp_label = Label(option_frame, width=199)
    temp_label.grid(row=0, column=2)
    data_count_label = Label(option_frame, width=30, height=1, font=('맑은 고딕', 15, 'bold'))
    data_count_label.grid(row=0, column=5,)
    count() # count 실행


# Treeview Style
style = ttk.Style()
style.theme_use('clam')
style.configure('Treeview.Heading', font=font)
style.configure('Treeview',
                    background = '#D3D3D3',
                    foreground = 'black',
                    fieldbackground = '#D3D3D3',
                    rowheight=30,
                    font=font)    

# Treeview1 생성 
column1 = '번호'
column2 = '접수상태'
column3 = '부서'
column4 = '성명'
column5 = '신청내용'
column6 = '기타입력'

tree = ttk.Treeview(frame1)
tree.grid(row=2, column=0, columnspan=2, pady=10)

yscrollcommand = ttk.Scrollbar(frame1, orient=VERTICAL, command=tree.yview)
yscrollcommand.grid(row=2, column=2, pady=10)
yscrollcommand.grid_configure(sticky='ns')

tree.tag_configure('Treeview.Heading', **{'font': font})
tree.configure(height=22)
tree.configure(yscrollcommand=yscrollcommand.set)
tree.columnconfigure(0, weight=1)
tree.rowconfigure(0, weight=1)
tree.grid_configure(sticky='nsew')

# Treeview 열 설정
tree['columns'] = (column1, 
                    column2, 
                    column3, 
                    column4, 
                    column5, 
                    column6)
tree['show'] = 'headings'

tree.column(column1, stretch=NO, minwidth=0, width=0) # 폭 조정으로 column1을 숨김
tree.column(column2, width=100, anchor="center")
tree.column(column3, width=200, anchor="center")
tree.column(column4, width=150, anchor="center")
tree.column(column5, width=1000)
tree.column(column6, width=400)

tree.heading(column1, text='번호')
tree.heading(column2, text='접수상태')
tree.heading(column3, text='부서')
tree.heading(column4, text='성명')
tree.heading(column5, text='신청내용')
tree.heading(column6, text='기타입력')

# Button Frame
btn_frame = tk.Frame(frame1)
btn_frame.grid(row=5, column=0, padx=30, pady=10, sticky='sw')

# Finished Button
def finished(event=None):
    selection = tree.selection()
    selection = selection[0]
    data = tree.item(selection)['values']
    row_id = data[0] # 첫 번째 값인 id를 가져옴
    c.execute("UPDATE App_DB SET column1 = ? WHERE id = ?", ("완료", row_id))
    conn.commit()
    update_all_table()
    messagebox.showinfo('상태','전체_DB')
finished_btn = tk.Button(btn_frame, width=10, text = '작업완료', command=finished, font=font_bold, relief='groove')
finished_btn.grid(row=0, column=1, padx=20, pady=15, sticky='sw')

# csv export
def csv_export():
    file_path = filedialog.asksaveasfilename(defaultextension='.csv') # 탐색기 실행
    if file_path: # if : 선택한 파일 경로가 있을 경우
        df = pd.read_sql_query("SELECT * from App_DB", conn) # DB에서 데이터 추출하고
        df.to_csv(file_path, index=False, encoding='utf-8-sig') # csv 파일을 utf-8로 저장
        messagebox.showinfo("완료", "파일 저장이 완료되었습니다.") # 완료 메시지
export_button = tk.Button(btn_frame, width=10, text='내보내기', command=csv_export, font=font_bold, relief='groove')
export_button.grid(row=0, column=3, padx=20, pady=15, sticky='sw')

# csv import
def csv_import():
    file_path = filedialog.askopenfilename(defaultextension='.csv') # 탐색기 실행
    if file_path: # if : 선택한 파일 경로가 있을 경우
        df = pd.read_csv(file_path) # csv 데이터를 가져와서
        for row in df.itertuples(): # DB에 데이터 추가
            c.execute("INSERT INTO App_DB (column1, column2, column3, column4, column5) VALUES (?, ?, ?, ?, ?)",
                        ('신청접수', row.column2, row.column3, row.column4, row.column5))
        conn.commit() # DB 접속 종료
        update_all_table() # 트리뷰 업데이트
        messagebox.showinfo("완료", "파일 불러오기가 완료되었습니다.") # 완료 메시지
import_button = tk.Button(btn_frame, width=10, text='가져오기', command=csv_import, font=font_bold, relief='groove')
import_button.grid(row=0, column=2, padx=20, pady=15, sticky='sw')

# # Text Frame
# text_frame = tk.Frame(frame1)
# text_frame.grid(row=5, column=1, padx=0, pady=0,)

# # Manual Label
# manual = tk.Label(text_frame, text='| 탭이동 : Ctrl + 방향키 | 검색하기 : Ctrl + F | 수정하기 : F2 | 새로고침 : F5 | 신청접수와 작업완료 전환 : F8 |', font=font_mini)
# manual.grid(padx=50, pady=0, sticky='sw')



# TAB 2 =============================================================
frame2 = Frame(notebook)
notebook.add(frame2, text="전체_DB")

# Treeview2 생성 
column1 = '번호'
column2 = '접수상태'
column3 = '부서'
column4 = '성명'
column5 = '신청내용'
column6 = '기타입력'

tree2 = ttk.Treeview(frame2)
tree2.grid(row=2, column=0, columnspan=2, pady=10)

yscrollcommand2 = ttk.Scrollbar(frame2, orient=VERTICAL, command=tree2.yview)
yscrollcommand2.grid(row=2, column=2, pady=10, sticky="ns")

tree2.tag_configure('Treeview.Heading', **{'font': font}) # 스타일 적용
tree2.configure(height=26)
tree2.configure(yscrollcommand=yscrollcommand2.set) # 높이 조정
tree2.columnconfigure(0, weight=1) # column weight 설정
tree2.rowconfigure(0, weight=1) # row weight 설정
tree2.grid_configure(sticky="nsew") # 가운데 정렬

tree2['columns'] = (column1, 
                    column2, 
                    column3, 
                    column4, 
                    column5, 
                    column6)
tree2['show'] = 'headings'

tree2.column(column1, stretch=NO, minwidth=0, width=0) # 폭 조정으로 column1을 숨김
tree2.column(column2, width=100, anchor="center")
tree2.column(column3, width=200, anchor="center")
tree2.column(column4, width=150, anchor="center")
tree2.column(column5, width=1000)
tree2.column(column6, width=400)

tree2.heading(column1, text='번호')
tree2.heading(column2, text='접수상태')
tree2.heading(column3, text='부서')
tree2.heading(column4, text='성명')
tree2.heading(column5, text='신청내용')
tree2.heading(column6, text='기타입력')

# 완료된 항목을 신청접수로 되돌리기
def running(event=None):
    selection = tree2.selection()
    selection = selection[0]
    data = tree2.item(selection)['values']
    row_id = data[0] # 첫 번째 값인 id를 가져옴
    c.execute("UPDATE App_DB SET column1 = ? WHERE id = ?", ("신청접수", row_id))
    conn.commit()
    update_all_table()
    messagebox.showinfo('상태','신청접수')

# 완료된 작업 취소하기
cancel_btn = tk.Button(frame2, text='완료작업 취소하기', command=running, width=20, height=1, font=font_bold, relief='groove')
cancel_btn.grid(row=3, column=0, padx=50, pady=15, sticky='sw')
# # 단축키 설명
# manual2 = tk.Label(frame2, text='| 탭이동 : Ctrl + 방향키 | 검색하기 : Ctrl + F | 수정하기 : F2 | 새로고침 : F5 | 신청접수와 작업완료 전환 : F8 |', font=font_mini)
# manual2.grid(padx=50, pady=20, sticky='sw')



# def =============================================================

# TAB 클릭 이벤트
def on_tab_click(event):
    x, y = event.x, event.y
    element = event.widget.identify(x, y)
    if "label" in element:
        tab_index = event.widget.index("@%d,%d" % (x, y))
        switch_tab(tab_index=tab_index)
notebook.bind("<Button-1>", on_tab_click)

# Ctrl + Right / Left 키를 이용한 TAB 전환 
current_tab = -1
def switch_tab(event=None, tab_index=None):
    global current_tab
    if tab_index is not None:
        next_tab = tab_index
    elif current_tab == -1:
        current_tab = notebook.index(notebook.select())
        return
    else:
        num_tabs = notebook.index("end")  # 탭의 개수
        if event.keysym == "Right":
            next_tab = current_tab + 1 if current_tab < num_tabs - 1 else current_tab
            # 다음 탭으로 이동 (마지막 탭일 경우 탭 유지)
        elif event.keysym == "Left":
            next_tab = current_tab - 1 if current_tab > 0 else current_tab
            # 이전 탭으로 이동 (첫 번째 탭일 경우 탭 유지)
        else:
            return
    notebook.select(next_tab)
    if next_tab == 0: 
        print('탭이동0') # 탭 이동시 추가 기능 입력 가능
    elif next_tab == 1:
        print('탭이동1') # 탭 이동시 추가 기능 입력 가능
    current_tab = next_tab
# 탭 전환 바인딩
root.bind("<Control-Right>", switch_tab)
root.bind("<Control-Left>", switch_tab)

# Header 클릭을 통한 데이터 정렬
def treeview_sort_column(tv, col, reverse):
    """Treeview 열 정렬 함수"""
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    l.sort(reverse=reverse)
    # 정렬할 데이터가 숫자일 경우
    # l.sort(reverse=reverse, key=lambda x: int(x[0]))
    # 정렬할 데이터가 날짜일 경우
    # l.sort(reverse=reverse, key=lambda x: datetime.datetime.strptime(x[0], '%Y-%m-%d'))
    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)
    # 다시 한번 클릭 시 정렬 순서를 바꾸기 위해 열 상태를 기록
    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))
# Tree1과 Tree2의 열 헤더를 클릭 시 정렬
for col in tree['columns']:
    tree.heading(col, text=col, command=lambda _col=col: treeview_sort_column(tree, _col, False))
for col in tree2['columns']:
    tree2.heading(col, text=col, command=lambda _col=col: treeview_sort_column(tree2, _col, False))

# Treeview 더블 클릭을 통한 데이터 수정
def on_double_click(event):
    # 첫 번째 열(헤더)에서 더블클릭이 일어나면 함수를 끝냄
    if event.widget.identify_region(event.x, event.y) == 'heading':
        print('heading 더블 클릭 차단')
        return
    
    # 선택한 행의 값들 가져오기
    selection = tree.selection()
    if len(selection) == 0:
        return
    values = tree.item(selection[0], 'values')
    
    # Toplevel 창 열기
    top = Toplevel(root)
    top.title('수정')
    top.geometry('+800+300')
    # top.iconbitmap('logo.ico')
    # Frame
    new_ent_frame = Frame(top)
    new_ent_frame.pack(padx=15)
    
    # 라벨 이름 설정
    new_lab1 = '부서'
    new_lab2 = '성명'
    new_lab3 = '신청내용'
    new_lab4 = '기타입력'
    
    new_lab1 = Label(new_ent_frame, text=new_lab1, font=font_bold)
    new_lab1.grid(row=0, column=0, sticky=W)
    new_ent1 = Entry(new_ent_frame, font=font, width=50, relief='flat', highlightthickness=1)
    new_ent1.config(highlightbackground='gray')
    new_ent1.grid(row=0, column=1, padx=10, pady=10, sticky=W)
    new_ent1.focus() # new_ent1에 프롬프트 설정
    
    new_lab2 = Label(new_ent_frame, text=new_lab2, font=font_bold)
    new_lab2.grid(row=1, column=0, sticky=W)
    new_ent2 = Entry(new_ent_frame, font=font, width=50, relief='flat', highlightthickness=1)
    new_ent2.config(highlightbackground='gray')
    new_ent2.grid(row=1, column=1, padx=10, pady=10, sticky=W)
    
    new_lab3 = Label(new_ent_frame, text=new_lab3, font=font_bold)
    new_lab3.grid(row=2, column=0, sticky=W)
    new_ent3 = Entry(new_ent_frame, font=font, width=100, relief='flat', highlightthickness=1)
    new_ent3.config(highlightbackground='gray')
    new_ent3.grid(row=2, column=1, padx=10, pady=10, sticky=W)
    
    new_lab4 = Label(new_ent_frame, text=new_lab4, font=font_bold)
    new_lab4.grid(row=3, column=0, sticky=W)
    new_ent4 = Entry(new_ent_frame, font=font, width=50, relief='flat', highlightthickness=1)
    new_ent4.config(highlightbackground='gray')
    new_ent4.grid(row=3, column=1, padx=10, pady=10, sticky=W)

    # 기존 DB 데이터를 ent에 입력
    new_ent1.insert(0, values[2])  # column3
    new_ent2.insert(0, values[3])  # column4
    new_ent3.insert(0, values[4])  # column5
    new_ent4.insert(0, values[5])  # column6

    # Save
    def edit_item(event=None): # 이벤트 동작될 수 있음
        if messagebox.askyesno("EDIT", "변경된 내용은 복구할 수 없습니다.\n저장하려면 YES를 누르세요."):
            # 변경된 값 가져오기
            new_value1 = new_ent1.get()
            new_value2 = new_ent2.get()
            new_value3 = new_ent3.get()
            new_value4 = new_ent4.get()
            # DB 업데이트 [변경 내용이 있는 항목만 업데이트]
            row_id = values[0]
            if new_value1:
                c.execute("UPDATE App_DB SET column2 = ? WHERE id = ?", (new_value1, row_id))
            if new_value2:
                c.execute("UPDATE App_DB SET column3 = ? WHERE id = ?", (new_value2, row_id))
            if new_value3:
                c.execute("UPDATE App_DB SET column4 = ? WHERE id = ?", (new_value3, row_id))
            if new_value4:
                c.execute("UPDATE App_DB SET column5 = ? WHERE id = ?", (new_value4, row_id))
            conn.commit()
            # Toplevel 창 닫기
            top.destroy()
            # Treeview 업데이트
            update_all_table()
    # 저장버튼
    save_button = Button(top, text='변경하기', command=edit_item, font=font_bold, relief='groove')
    save_button.pack(padx=15, pady=15)
    
    # bind
    new_ent1.bind('<Return>', next_entry)
    new_ent2.bind('<Return>', next_entry)
    new_ent3.bind('<Return>', next_entry)
    new_ent1.focus()
    new_ent4.bind('<Return>', edit_item)
    top.bind('<Escape>', lambda event: top.destroy())
    
    # 열과 높이 설정
    top.columnconfigure(0, weight=1)
    top.rowconfigure(0, weight=1)

# Treeview 검색 창 ========================================================================
def search_ctrl_f(event):
    # Toplevel 창 열기
    search_popup = Toplevel(root)
    search_popup.title('검색')
    search_popup.geometry('+800+300')
    # search_popup.iconbitmap('logo.ico')
    search_var = StringVar()
    search_var.trace('w', lambda name, index, mode: search()) # Entry에 입력할 때마다 search 함수 호출
    search_entry = Entry(search_popup, textvariable=search_var, font=font, relief='flat', highlightthickness=1)
    search_entry.config(highlightbackground='gray', width=30)
    search_entry.grid(row=0, column=1, padx=10, pady=10)
    search_entry.focus()
    # 전체 검색
    def search():
        search1()
        search2()
    # 신청접수 검색
    def search1(event=None):
        search_term = search_var.get()
        if search_term:
            # 이전 검색 결과 초기화
            for row in tree.get_children():
                tree.item(row, tags=())
            # 검색어와 일치하는 row만 선택
            found_rows = []
            for row in tree.get_children():
                item = tree.item(row)
                values = item['values']
                if search_term.lower() in str(values).lower():
                    found_rows.append(row)
                    tree.item(row, tags=('found',))
            # 일치하는 row만 보여주도록 설정
            tree.tag_configure('found', background=search_color)
        else:
            # 검색어가 없으면 모든 row 표시
            for row in tree.get_children():
                tree.item(row, tags=())
    # 전체DB 검색
    def search2(event=None):
        search_term = search_var.get()
        if search_term:
            # 이전 검색 결과 초기화
            for row in tree2.get_children():
                tree2.item(row, tags=())
            # 검색어와 일치하는 row만 선택
            found_rows = []
            for row in tree2.get_children():
                item = tree2.item(row)
                values = item['values']
                if search_term.lower() in str(values).lower():
                    found_rows.append(row)
                    tree2.item(row, tags=('found',))
            # 일치하는 row만 보여주도록 설정
            tree2.tag_configure('found', background=search_color)
        else:
            # 검색어가 없으면 모든 row 표시
            for row in tree2.get_children():
                tree2.item(row, tags=())
    # 검색창 닫기
    def destroy_search_popup(event=True):
        search_popup.destroy()
    search_entry.bind('<Return>', destroy_search_popup)
    search_entry.bind('<Escape>', destroy_search_popup)
root.bind('<Control-f>', search_ctrl_f)

# 다음 entry로 이동
def next_entry(event):
    widget = event.widget
    widget.tk_focusNext().focus()
    return "break"

# ent1로 이동하기
def move_ent1():
    ent1.focus_set()

# 이벤트 처리 함수 (단축키 설정 등)
def key_enter(event): # 저장
    messagebox.showinfo('ADD','추가 되었습니다.')
    save_data()

def key_f5(event): # 새로고침
    tree.delete(*tree.get_children())
    update_all_table()
    messagebox.showinfo('UPDATE','새로 고침.')

def key_delete1(event): # 삭제
    if messagebox.askyesno("DELETE", "삭제하시겠습니까?"):
        delete_data1()

def key_delete2(event): # 삭제
    if messagebox.askyesno("DELETE", "삭제하시겠습니까?"):
        delete_data2()

# tree와 tree2의 DB 업데이트
def update_all_table():
    update_receipt_table()  # tree의 업데이트
    update_complete_table() # tree2의 업데이트
    work_data_count() # tree의 접수량 라벨 업데이트













# 신청접수 영역 DB 업데이트
def update_receipt_table(): 
    # Treeview 모든 데이터 삭제
    tree.delete(*tree.get_children()) 
    with sqlite3.connect('data.db') as conn:
        c = conn.cursor()
        # 신청접수인 row 값만 가져옴
        c.execute("SELECT * FROM App_DB WHERE column1=?", ('신청접수',))
        rows = c.fetchall()
        # column6, column5 기준으로 정렬
        rows_sorted = sorted(rows, key=lambda x: (x[5], x[4]))
        # 정렬된 데이터 Treeview에 삽입
        for row in rows_sorted:
            # Treeview에 삽입
            tree.insert("", "end", values=row) 
















# # 신청접수 영역 DB 업데이트
# def update_receipt_table(): 
#     # Treeview 모든 데이터 삭제
#     tree.delete(*tree.get_children()) 
#     with sqlite3.connect('data.db') as conn:
#         c = conn.cursor()
#         # 신청접수인 row 값만 가져옴
#         c.execute("SELECT * FROM App_DB WHERE column1=?", ('신청접수',))
#         rows = c.fetchall()
#         # column6 기준으로 정렬
#         rows_sorted = sorted(rows, key=lambda x: x[5])
#         # 정렬된 데이터 Treeview에 삽입
#         for row in rows_sorted:
#             # Treeview에 삽입
#             tree.insert("", "end", values=row) 

# 전체DB 영역 DB 업데이트
def update_complete_table(): 
    # Treeview 모든 데이터 삭제
    tree2.delete(*tree2.get_children()) 
    # DB에서 모든 데이터 가져오기
    rows = c.execute("SELECT * FROM App_DB").fetchall()
    # column6 기준으로 정렬
    rows_sorted = sorted(rows, key=lambda x: x[5])
    # 정렬된 데이터 Treeview에 삽입
    for row in rows_sorted:
        # Treeview에 삽입
        tree2.insert("", "end", values=row) 

def save_data():
    # 엔트리 데이터
    team = ent1.get()
    name = ent2.get()
    helpdesk = ent3.get()
    etc = ent4.get()
    # DB에 데이터 추가
    c.execute("SELECT MAX(id) FROM App_DB")
    last_id = c.fetchone()[0]
    if last_id is None:
        last_id = 0
    new_id = last_id + 1
    c.execute("INSERT INTO App_DB (id, column1, column2, column3, column4, column5) VALUES (?, ?, ?, ?, ?, ?)",
                (new_id, '신청접수', team, name, helpdesk, etc))
    conn.commit()
    # 트리뷰에 데이터 추가
    tree.insert('', 'end', text=new_id, values=(team, name, helpdesk, etc))
    # 입력란 초기화 하기
    ent1.delete(0, tk.END)
    ent2.delete(0, tk.END)
    ent3.delete(0, tk.END)
    ent4.delete(0, tk.END)
    # 테이블 업데이트
    update_all_table()
    # ent1로 프롬프트 이동
    move_ent1()

# 삭제 함수 (Treeview와 DB 모두 삭제)
def delete_data1():
    selection = tree.selection()
    print(selection)
    # 선택된 항목이 없는 경우, 오류 표시 후 함수 종료
    if not selection:
        messagebox.showerror("Error", "삭제할 항목을 선택해주세요.")
        return
    # 선택된 모든 항목에 대해 반복하여 삭제
    for index, item in enumerate(selection):
        data = tree.item(item)['values']
        row_id = data[0] # 첫 번째 값인 id를 가져옴
        c.execute("DELETE FROM App_DB WHERE id=?", (row_id,)) # row_id에 해당하는 row 데이터 전체 삭제
        tree.delete(item)
    # 선택된 모든 항목에 대해 id를 수정
    for i, item in enumerate(tree.get_children()):
        new_id = i+1
        old_id = tree.item(item)['values'][0]
        c.execute("UPDATE App_DB SET id=? WHERE id=?", (new_id, old_id))
    conn.commit()
    # 테이블 갱신
    update_all_table()

# 삭제 함수 (Treeview2와 DB 모두 삭제)
def delete_data2():
    selection = tree2.selection()
    print(selection)
    # 선택된 항목이 없는 경우, 오류 표시 후 함수 종료
    if not selection:
        messagebox.showerror("Error", "삭제할 항목을 선택해주세요.")
        return
    # 선택된 모든 항목에 대해 반복하여 삭제
    for index, item in enumerate(selection):
        data = tree2.item(item)['values']
        row_id = data[0] # 첫 번째 값인 id를 가져옴
        c.execute("DELETE FROM App_DB WHERE id=?", (row_id,)) # row_id에 해당하는 row 데이터 전체 삭제
        tree2.delete(item)
    # 선택된 모든 항목에 대해 id를 수정
    for i, item in enumerate(tree2.get_children()):
        new_id = i+1
        old_id = tree2.item(item)['values'][0]
        c.execute("UPDATE App_DB SET id=? WHERE id=?", (new_id, old_id))
    conn.commit()
    # 테이블 갱신
    update_all_table()

# 바인드 설정 =====================================================================================================
ent1.bind('<Return>', next_entry)
ent2.bind('<Return>', next_entry)
ent3.bind('<Return>', next_entry)
ent4.bind('<Return>', key_enter)

tree.bind('<Double-1>', on_double_click)
tree.bind('<F2>', on_double_click)
tree.bind('<F8>', finished)
tree.bind('<Delete>', key_delete1)

tree2.bind('<F8>', running)
tree2.bind('<Delete>', key_delete2)

root.bind('<F5>', key_f5)

# Treeview 현재 대기자 수
tree.bind('<<TreeviewSelect>>', lambda event: work_data_count())

# 데이터 테이블 초기화
update_all_table()

# 프로그램 실행
root.mainloop()

# db 연결 종료
conn.close()
