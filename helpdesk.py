from tkinter import * 
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
from tkinter.ttk import Combobox
import tkinter as tk
import sqlite3
import pandas as pd

# =============================
# 설정
bold_font = '맑은 고딕', 11, 'bold'
font = '맑은 고딕', 11
mini_font = '맑은 고딕', 8
title_info = 'Helpdesk'
geometry = '1800x1000+400+50'
search_color = 'khaki'
click_color = 'LightCyan4'
# =============================

# db 연결
conn = sqlite3.connect('data.db')
c = conn.cursor()
# ==================================================================
# 테이블 생성
# column1 = 완료유무
# column2 = 부서
# column3 = 성명
# column4 = 신청내용
# column5 = 비고

c.execute('''CREATE TABLE IF NOT EXISTS App_DB
            (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            column1 TEXT,
            column2 TEXT,
            column3 TEXT,
            column4 TEXT,
            column5 TEXT,
            column6 TEXT
            )''')
conn.commit()

# tkinter GUI 생성
root = tk.Tk()
root.title(title_info) #타이틀
root.geometry(geometry) #해상도와 시작위치
root.state('zoomed') #최대화
# root column과 row의 weight 설정
root.columnconfigure(0, weight=1)
root.rowconfigure(2, weight=1)

# 'nw': 상단, 'ne': 오른쪽, 'sw': 왼쪽, 'se': 하단
# 최상단에 날짜/시간
date_frame = tk.Frame(root)
date_frame.grid(pady=30)

#==================================================================
notebook = ttk.Notebook(root)
notebook.grid()

frame1 = Frame(notebook)
notebook.add(frame1, text="신청접수")

frame2 = Frame(notebook)
notebook.add(frame2, text="작업완료")

# TAB 전환 함수 (Frame1과 Frame2의 전환 단축키)================================================
# TAB을 전환할 때 각 TAB에 있는 Treeview를 업데이트 해준다.
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
            # 다음 탭의 인덱스 (마지막 탭일 경우 현재 탭 유지)
        elif event.keysym == "Left":
            next_tab = current_tab - 1 if current_tab > 0 else current_tab
            # 이전 탭의 인덱스 (첫 번째 탭일 경우 현재 탭 유지)
        else:
            return
    notebook.select(next_tab)
    if next_tab == 0:
        update_table1()
    elif next_tab == 1:
        update_table()
    current_tab = next_tab
    
# 탭 전환 바인딩
root.bind("<Control-Right>", switch_tab)
root.bind("<Control-Left>", switch_tab)

# 마우스 클릭 이벤트
def on_tab_click(event):
    x, y = event.x, event.y
    element = event.widget.identify(x, y)
    if "label" in element:
        tab_index = event.widget.index("@%d,%d" % (x, y))
        switch_tab(tab_index=tab_index)
# 마우스 바인딩
notebook.bind("<Button-1>", on_tab_click)



#==================================================================
input_frame = tk.Frame(frame1)
input_frame.grid(row=0, column=0, padx=10, pady=10, sticky='sw')

option_frame = tk.Frame(frame1)
option_frame.grid(row=1, column=0, padx=10, pady=10, sticky='sw')

tree = ttk.Treeview(frame1)
tree.grid(row=2, column=0, columnspan=2, pady=10)

yscrollcommand = ttk.Scrollbar(frame1, orient=VERTICAL, command=tree.yview)
yscrollcommand.grid(row=2, column=2, pady=10)
yscrollcommand.grid_configure(sticky="ns")

btn_frame = tk.Frame(frame1)
btn_frame.grid(row=5, column=0, padx=30, pady=10, sticky='sw')

text_frame = tk.Frame(frame1)
text_frame.grid(row=6, column=0, padx=0, pady=10, sticky='sw')



#==================================================================
tree2 = ttk.Treeview(frame2)
tree2.grid(row=2, column=0, columnspan=2, pady=0)

yscrollcommand2 = ttk.Scrollbar(frame2, orient=VERTICAL, command=tree.yview)
yscrollcommand2.grid(row=2, column=2, pady=0)
yscrollcommand2.grid_configure(sticky="ns")



#=============================================================
# 라벨 이름 설정
lab1 = '부서'
lab2 = '성명'
lab3 = '신청내용'
lab4 = '기타입력'

lab1 = tk.Label(input_frame, text=lab1, font=font)
lab1.grid(row=0, column=0, padx=20, sticky=tk.W)
ent1 = tk.Entry(input_frame, font=font)
ent1.grid(row=1, column=0, padx=20)
ent1.focus() # ent1에 프롬프트 설정

lab2 = tk.Label(input_frame, text=lab2, font=font)
lab2.grid(row=0, column=1, padx=20 ,sticky=tk.W)
ent2 = tk.Entry(input_frame, font=font)
ent2.grid(row=1, column=1, padx=20)

lab3 = tk.Label(input_frame, text=lab3, font=font)
lab3.grid(row=0, column=2, padx=20 ,sticky=tk.W)
ent3 = tk.Entry(input_frame, width=60, font=font)
ent3.grid(row=1, column=2, padx=20)

lab4 = tk.Label(input_frame, text=lab4, font=font)
lab4.grid(row=0, column=3, padx=20 ,sticky=tk.W)
ent4 = tk.Entry(input_frame, width=30, font=font)
ent4.grid(row=1, column=3, padx=20)

#검색창===============================================================================
search_var = StringVar()
search_var.trace('w', lambda name, index, mode: search()) # Entry에 입력할 때마다 search 함수 호출

search_label = tk.Label(option_frame, text='검색하기', font=bold_font) 
search_label.grid(row=0, column=1, padx=20, sticky=tk.W)

search_entry = ttk.Entry(option_frame, textvariable=search_var)
search_entry.grid(row=1, column=1, padx=20)

# 엔트리 데이터
team = ent1.get()
name = ent2.get()
helpdesk = ent3.get()
etc = ent4.get()

# 트리뷰 스타일 지정 =============================================================
style = ttk.Style()
style.theme_use('clam')
style.configure('Treeview.Heading', font=font)
style.configure('Treeview', background = '#D3D3D3', foreground = 'black', rowheight=25, fieldbackground = '#D3D3D3', font=font)
# style.map('Treeview', background = [('selected', click_color)])

# Treeview1 생성 =============================================================
column1 = '번호'
column2 = '접수상태'
column3 = '부서'
column4 = '성명'
column5 = '신청내용'
column6 = '기타입력'

tree.tag_configure('Treeview.Heading', **{'font': font}) # 스타일 적용
tree.configure(height=20)
tree.configure(yscrollcommand=yscrollcommand.set)
tree.columnconfigure(0, weight=1) # column weight 설정
tree.rowconfigure(0, weight=1) # row weight 설정
tree.grid_configure(sticky="nsew") # 가운데 정렬

# Treeview 열 설정
tree['columns'] = (column1, 
                    column2, 
                    column3, 
                    column4, 
                    column5, 
                    column6)
tree['show'] = 'headings'

tree.column(column1, width=100, anchor="center")
tree.column(column2, width=150, anchor="center")
tree.column(column3, width=200, anchor="center")
tree.column(column4, width=150, anchor="center")
tree.column(column5, width=600)
tree.column(column6, width=200)

tree.heading(column1, text='번호')
tree.heading(column2, text='접수상태')
tree.heading(column3, text='부서')
tree.heading(column4, text='성명')
tree.heading(column5, text='신청내용')
tree.heading(column6, text='기타입력')

# Treeview2 생성 =============================================================
column1 = '번호'
column2 = '접수상태'
column3 = '부서'
column4 = '성명'
column5 = '신청내용'
column6 = '기타입력'

tree2.tag_configure('Treeview.Heading', **{'font': font}) # 스타일 적용
tree2.configure(height=30)
tree2.configure(yscrollcommand=yscrollcommand.set)
tree2.columnconfigure(0, weight=1) # column weight 설정
tree2.rowconfigure(0, weight=1) # row weight 설정
tree2.grid_configure(sticky="nsew") # 가운데 정렬

# Treeview 생성 =============================================================
# Treeview 열 설정
tree2['columns'] = (column1, 
                    column2, 
                    column3, 
                    column4, 
                    column5, 
                    column6)
tree2['show'] = 'headings'

tree2.column(column1, width=100, anchor="center")
tree2.column(column2, width=150, anchor="center")
tree2.column(column3, width=200, anchor="center")
tree2.column(column4, width=150, anchor="center")
tree2.column(column5, width=600)
tree2.column(column6, width=200)

tree2.heading(column1, text='번호')
tree2.heading(column2, text='접수상태')
tree2.heading(column3, text='부서')
tree2.heading(column4, text='성명')
tree2.heading(column5, text='신청내용')
tree2.heading(column6, text='기타입력')


# 검색 Entry 연결 =======================================================
def search(event=None):
    search_term = search_var.get()
    if search_term:
        # 이전 검색 결과 초기화
        for row in tree.get_children():
            tree.item(row, tags=())
        # 검색어와 일치하는 row만 선택
        for row in tree.get_children():
            item = tree.item(row)
            values = item['values']
            if search_term.lower() in str(values).lower():
                tree.item(row, tags=('found',))
        # 일치하는 row만 보여주도록 설정
        tree.tag_configure('found', background=search_color)
    else:
        # 검색어가 없으면 모든 row 표시
        for row in tree.get_children():
            tree.item(row, tags=())


# 오늘 날짜 | 현재시간 설정 
import time
weekdays_kr = ['월', '화', '수', '목', '금', '토', '일']
date_label = tk.Label(date_frame, font=('맑은 고딕', 20, 'bold'))
time_label = tk.Label(date_frame, font=('맑은 고딕', 20, 'bold'))
def update_time():
    current_time = time.strftime('현재시각 : %H시 %M분 |')
    time_label.config(text=current_time)
    time_label.after(1000, update_time)
def update_date():
    current_date = time.strftime('| %Y년 %m월 %d일 ') + weekdays_kr[time.localtime().tm_wday] + '요일 | '
    date_label.config(text=current_date)
    date_label.after(1000, update_date)
def show():
    date_label.grid(row=0, column=0, padx=0)
    time_label.grid(row=0, column=1, padx=0)
show()
update_date()
update_time()


# Treeview 더블 클릭 함수 ========================================================================
def on_double_click(event):
    # 선택한 행의 값들 가져오기
    selection = tree.selection()
    if len(selection) == 0:
        return
    values = tree.item(selection[0], 'values')

    # Toplevel 창 열기
    top = tk.Toplevel(root)
    top.title('')
    top.geometry('+800+300')
    # Frame
    new_ent_frame = tk.Frame(top)
    new_ent_frame.pack(padx=15)
    
    # 라벨 이름 설정
    new_lab1 = '부서'
    new_lab2 = '성명'
    new_lab3 = '신청내용'
    new_lab4 = '기타입력'
    
    new_lab1 = tk.Label(new_ent_frame, text=new_lab1, font=font)
    new_lab1.grid(row=0, column=0, sticky=tk.W)
    new_ent1 = tk.Entry(new_ent_frame, font=font)
    new_ent1.grid(row=1, column=0)
    new_ent1.focus() # new_ent1에 프롬프트 설정

    new_lab2 = tk.Label(new_ent_frame, text=new_lab2, font=font)
    new_lab2.grid(row=2, column=0, sticky=tk.W)
    new_ent2 = tk.Entry(new_ent_frame, font=font)
    new_ent2.grid(row=3, column=0)
    
    new_lab3 = tk.Label(new_ent_frame, text=new_lab3, font=font)
    new_lab3.grid(row=4, column=0, sticky=tk.W)
    new_ent3 = tk.Entry(new_ent_frame, font=font)
    new_ent3.grid(row=5, column=0)
    
    new_lab4 = tk.Label(new_ent_frame, text=new_lab4, font=font)
    new_lab4.grid(row=6, column=0, sticky=tk.W)
    new_ent4 = tk.Entry(new_ent_frame, font=font)
    new_ent4.grid(row=7, column=0)

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
            update_table1()

    # 저장버튼
    save_button = tk.Button(top, text='변경하기', command=edit_item)
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


# ===============================================================================================================================





# 데이터 헤더를 눌러 정렬하기
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

# 열 헤더를 클릭 시 정렬
for col in tree['columns']:
    tree.heading(col, text=col, command=lambda _col=col: treeview_sort_column(tree, _col, False))











# =============================================================================================================================
# 다음 entry로 이동
def next_entry(event):
    widget = event.widget
    widget.tk_focusNext().focus()
    return "break"

# ent1로 이동하기
def move_ent1():
    ent1.focus_set()

# 이벤트 처리 함수 (단축키 설정 등) ======================================================
def key_enter(event): # 저장
    messagebox.showinfo('ADD','추가 되었습니다.')
    save_data()

def key_f5(event): # 새로고침
    tree.delete(*tree.get_children())
    update_table1()
    messagebox.showinfo('UPDATE','새로 고침.')
    
def search_delete(event): # 검색 단축키 설정
    search_entry.delete(0, tk.END)
    search_entry.focus()
    
def key_delete1(event): # 삭제
    if messagebox.askyesno("DELETE", "삭제하시겠습니까?"):
        delete_data1()

def key_delete2(event): # 삭제
    if messagebox.askyesno("DELETE", "삭제하시겠습니까?"):
        delete_data2()

# 탭이 선택될 때 이벤트를 처리하는 함수
def on_tab_selected(event):
    selected_tab = event.widget.select()
    tab_text = event.widget.tab(selected_tab, "text")
    if tab_text == '전체목록': # 탭1이 선택되면 update_table 함수 호출
        update_table()
        print('전체목록')
# 탭 위젯에 탭 선택 이벤트 바인딩
notebook.bind("<<NotebookTabChanged>>", on_tab_selected)


# # Treeview 새로고침 ======================================================
def update_table(): # 모든 데이터를 보여줌, TAB2로 이동 했을 때 동작하도록
    tree2.delete(*tree2.get_children()) # Treeview 모든 데이터 삭제
    for row in c.execute("SELECT * FROM App_DB"): # DB에서 모든 데이터 가져오기
            tree2.insert("", "end", values=row) # Treeview에 삽입


def update_table1():
    # 이전 데이터 삭제
    tree.delete(*tree.get_children())
    
    with sqlite3.connect('data.db') as conn:
        c = conn.cursor()
        
        # 신청접수인 row 값만 가져옴
        c.execute("SELECT * FROM App_DB WHERE column1=?", ('신청접수',))
        rows = c.fetchall()
        
        # 가져온 데이터를 treeview에 추가
        for row in rows:
            tree.insert('', 'end', values=row)

# 저장 버튼 함수 ======================================================================
def save_data():
    save_button.config(state=tk.DISABLED) # 버튼 동작 후 비활성화
    root.after(1000, lambda: save_button.config(state=tk.NORMAL)) # 1000의 시간이 지난 후 다시 활성화
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
    update_table1()
    # ent1로 프롬프트 이동
    move_ent1()


# # 삭제 함수 (Treeview와 DB 모두 삭제)
# def delete_data():
#     selection = tree.selection()
#     # 선택된 항목이 없는 경우, 오류 표시 후 함수 종료
#     if not selection:
#         messagebox.showerror("Error", "삭제할 항목을 선택해주세요.")
#         return
#     # 선택된 모든 항목에 대해 반복하여 삭제
#     for item in selection: # for 문이 없으면 여러 개를 클릭해도 하나만 삭제됨
#         data = tree.item(item)['values']
#         row_id = data[0] # 첫 번째 값인 id를 가져옴
#         c.execute("DELETE FROM App_DB WHERE id=?", (row_id,)) # row_id에 해당하는 row 데이터 전체 삭제
#         conn.commit()
#         tree.delete(item)
#     # 테이블 갱신
#     update_table1()




# 삭제 함수 (Treeview2와 DB 모두 삭제)
def delete_data1():
    selection = tree.selection()
    # 선택된 항목이 없는 경우, 오류 표시 후 함수 종료
    if not selection:
        messagebox.showerror("Error", "삭제할 항목을 선택해주세요.")
        return
    # 선택된 모든 항목에 대해 반복하여 삭제
    for item in selection: # for 문이 없으면 여러 개를 클릭해도 하나만 삭제됨
        data = tree.item(item)['values']
        row_id = data[0] # 첫 번째 값인 id를 가져옴
        c.execute("DELETE FROM App_DB WHERE id=?", (row_id,)) # row_id에 해당하는 row 데이터 전체 삭제
        # 삭제한 row_id 이후의 모든 id값을 1씩 감소시켜줌
        c.execute("UPDATE App_DB SET id=id-1 WHERE id > ?", (row_id,))
        conn.commit()
        tree.delete(item)
    # 테이블 갱신
    update_table()
    update_table1()

# 삭제 함수 (Treeview2와 DB 모두 삭제)
def delete_data2():
    selection = tree2.selection()
    # 선택된 항목이 없는 경우, 오류 표시 후 함수 종료
    if not selection:
        messagebox.showerror("Error", "삭제할 항목을 선택해주세요.")
        return
    # 선택된 모든 항목에 대해 반복하여 삭제
    for item in selection: # for 문이 없으면 여러 개를 클릭해도 하나만 삭제됨
        data = tree2.item(item)['values']
        row_id = data[0] # 첫 번째 값인 id를 가져옴
        c.execute("DELETE FROM App_DB WHERE id=?", (row_id,)) # row_id에 해당하는 row 데이터 전체 삭제
        # 삭제한 row_id 이후의 모든 id값을 1씩 감소시켜줌
        c.execute("UPDATE App_DB SET id=id-1 WHERE id > ?", (row_id,))
        conn.commit()
        tree2.delete(item)
    # 테이블 갱신
    update_table()
    update_table1()






# 상태변경 : 신청접수
def running(event=None):
    selection = tree.selection()
    selection = selection[0]
    data = tree.item(selection)['values']
    row_id = data[0] # 첫 번째 값인 id를 가져옴
    c.execute("UPDATE App_DB SET column1 = ? WHERE id = ?", ("신청접수", row_id))
    conn.commit()
    update_table1()
    messagebox.showinfo('상태','신청접수')

# 상태변경 : 작업완료
def finished(event=None):
    selection = tree.selection()
    selection = selection[0]
    data = tree.item(selection)['values']
    row_id = data[0] # 첫 번째 값인 id를 가져옴
    c.execute("UPDATE App_DB SET column1 = ? WHERE id = ?", ("완료", row_id))
    conn.commit()
    update_table1()
    messagebox.showinfo('상태','작업완료')

# csv파일 내보내기
def csv_export():
    # 파일 탐색기 창 열기
    file_path = filedialog.asksaveasfilename(defaultextension='.csv')
    # 선택한 파일 경로가 있을 경우
    if file_path:
        # DB에서 데이터 추출
        df = pd.read_sql_query("SELECT * from App_DB", conn)
        # CSV 파일로 저장
        df.to_csv(file_path, index=False, encoding='utf-8-sig') # csv 파일을 utf-8로 저장
        # 저장 완료 메시지 박스 출력
        messagebox.showinfo("완료", "파일 저장이 완료되었습니다.")

# csv파일 불러오기
def csv_import():
    # 파일 탐색기 창 열기
    file_path = filedialog.askopenfilename(defaultextension='.csv')
    # 선택한 파일 경로가 있을 경우
    if file_path:
        # CSV 파일에서 데이터 읽어오기
        df = pd.read_csv(file_path)
        # DB에 데이터 추가
        for row in df.itertuples():
            c.execute("INSERT INTO App_DB (column1, column2, column3, column4, column5) VALUES (?, ?, ?, ?, ?)",
                        ('신청접수', row.column2, row.column3, row.column4, row.column5))
        conn.commit()
        # 트리뷰 업데이트
        update_table1()

        # 불러오기 완료 메시지 박스 출력
        messagebox.showinfo("완료", "파일 불러오기가 완료되었습니다.")

# 시프트 상/하 키로 범위 지정 :: 동작하는지 확인이 필요함
def on_tree_select(event):
    # 범위 지정 확인
    if event.state == 1:
        # 이전 선택된 아이템과 현재 선택된 아이템의 인덱스 얻기
        cur_item = tree.selection()[0]
        prev_item = tree.focus()

        # 이전 아이템과 현재 아이템의 인덱스를 비교하여 범위 지정
        prev_index = tree.index(prev_item)
        cur_index = tree.index(cur_item)
        if prev_index < cur_index:
            items = tree.get_children('', start=prev_index, end=cur_index)
        else:
            items = tree.get_children('', start=cur_index, end=prev_index)

        # 범위 지정된 아이템들 하이라이트 표시
        for item in items:
            tree.selection_add(item)

    else:
        # 범위 지정이 아니면 현재 선택된 아이템만 하이라이트 표시
        cur_item = tree.focus()
        tree.selection_set(cur_item)


# 버튼 프레임 =====================================================================================================
# 완료처리
finished_btn = tk.Button(btn_frame, width=10, text = '작업완료', command=finished, font=font)
finished_btn.grid(row=0, column=1, padx=20, pady=5)

# 신청접수
running_btn = tk.Button(btn_frame, width=10, text = '신청접수', command=running, font=font)
running_btn.grid(row=0, column=2, padx=20, pady=5)

# CSV 가져오기
save_button = tk.Button(btn_frame, width=10, text='가져오기', command=csv_import, font=font)
save_button.grid(row=0, column=3, padx=20, pady=5, sticky='e')

# CSV 내보내기
save_button = tk.Button(btn_frame, width=10, text='내보내기', command=csv_export, font=font)
save_button.grid(row=0, column=4, padx=20, pady=5, sticky='e')

# 단축키 설명
manual = tk.Label(text_frame, text='| 탭이동 : Ctrl + 방향키 | 검색하기 : Ctrl + F | 수정하기 : F2 | 새로고침 : F5 | 취소처리 : F7 | 완료처리 : F8 |', font=mini_font)
manual.pack(padx=50)

# 바인드 설정 =====================================================================================================
ent1.bind('<Return>', next_entry)
ent2.bind('<Return>', next_entry)
ent3.bind('<Return>', next_entry)
ent4.bind('<Return>', key_enter)

search_entry.bind('<Return>', search)


tree.bind('<Double-1>', on_double_click)
tree.bind('<F2>', on_double_click)
tree.bind('<F7>', running)
tree.bind('<F8>', finished)

tree.bind('<Delete>', key_delete1)
tree2.bind('<Delete>', key_delete2)


root.bind('<Control-f>', search_delete)

# 데이터 테이블 초기화
update_table1()

# 프로그램 실행
root.mainloop()

# db 연결 종료
conn.close()