import os
import json
import pandas as pd
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import requests
from bs4 import BeautifulSoup

# CSV와 JSON 파일의 경로 설정
csv_file = 'members_directory.csv'
json_file = 'members_directory.json'

# 파일이 없으면 새로 생성하는 함수


def create_empty_files():
    # CSV 파일 생성
    if not os.path.exists(csv_file):
        with open(csv_file, 'w', encoding='utf-8-sig') as f:
            pass  # 빈 파일 생성

    # JSON 파일 생성
    if not os.path.exists(json_file):
        with open(json_file, 'w', encoding='utf-8-sig') as f:
            json.dump([], f)  # 빈 리스트로 초기화하여 JSON 파일 생성


# main.py 실행 시 파일 생성 확인
create_empty_files()

# 크롤링 함수


def crawl_and_save(presbytery_name, url):
    response = requests.get(url)
    if response.status_code == 200:
        html = response.text
        soup = BeautifulSoup(html, 'html.parser')
        data = []

        members = soup.find_all('div', class_='et_pb_blurb_content')

        for member in members:
            name_tag = member.find('h4', class_='et_pb_module_header')
            name = name_tag.text.strip() if name_tag else 'N/A'

            description = member.find('div', class_='et_pb_blurb_description')
            if description:
                details = description.find_all('i', class_='fa fa-check')
                church = details[0].next_sibling.strip() if len(
                    details) > 0 and details[0].next_sibling else 'N/A'
                address = details[1].next_sibling.strip() if len(
                    details) > 1 and details[1].next_sibling else 'N/A'

                phone_numbers = [a.text.strip() for a in description.find_all(
                    'a', href=True) if 'tel:' in a['href']]
                if not phone_numbers:
                    phone_numbers = ['N/A']

                email_tag = description.find('a', href=True, string=True)
                email = email_tag['href'].replace('mailto:', '').strip(
                ) if email_tag and 'mailto:' in email_tag['href'] else 'N/A'
            else:
                church = 'N/A'
                address = 'N/A'
                phone_numbers = ['N/A']
                email = 'N/A'

            # 'Church'가 'N/A'인 데이터를 제외
            if church != 'N/A':
                data.append({
                    'Name': name,
                    'Church': church,
                    'Address': address,
                    'Phone': ', '.join(phone_numbers),
                    'Email': email
                })

        if data:  # 유효한 데이터가 있을 경우에만 저장
            # CSV 파일로 저장
            with open(csv_file, 'a', encoding='utf-8-sig', newline='') as f:
                f.write(f"\n-- {presbytery_name} --\n")
                df = pd.DataFrame(data)
                df.to_csv(f, index=False, mode='a',
                          header=True, encoding='utf-8-sig')

            # JSON 파일로 저장 (덮어쓰기)
            json_data = {presbytery_name: data}
            with open(json_file, 'w', encoding='utf-8-sig') as f:
                json.dump(json_data, f, ensure_ascii=False, indent=4)

            messagebox.showinfo(
                "완료", f"{presbytery_name} 시찰회의 회원 주소록이 CSV와 JSON 형식으로 성공적으로 저장되었습니다.")

            # 테이블에 데이터를 추가 (중복 방지)
            add_data_to_table(presbytery_name, data)

        else:
            messagebox.showwarning(
                "경고", "유효한 데이터가 없습니다. 'Church' 항목이 모두 'N/A'입니다.")

        url_entry.delete(0, tk.END)  # 입력필드 초기화
        presbytery_entry.focus()  # 커서를 시찰명 입력필드로 이동
    else:
        messagebox.showerror("오류", f"웹 페이지를 가져오는 데 실패했습니다. 상태 코드: {
                             response.status_code}")

# 중복을 확인하여 테이블에 데이터를 추가하는 함수


def add_data_to_table(presbytery_name, new_data):
    existing_entries = {tuple(table.item(
        item)["values"]) for item in table.get_children()}  # 리스트를 튜플로 변환

    # 시찰명 추가
    table.insert("", "end", values=(
        f"[{presbytery_name}]", "", "", "", ""), tags=('presbytery',))

    for item in new_data:
        new_entry = (item['Name'], item['Church'],
                     item['Address'], item['Phone'], item['Email'])
        if new_entry not in existing_entries:
            table.insert("", "end", values=new_entry)

# CSV 데이터를 JSON으로 변환하여 저장하는 함수


def import_csv_to_json():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        try:
            df = pd.read_csv(file_path, encoding='utf-8-sig')
            data = df.to_dict(orient='records')
            with open(json_file, 'w', encoding='utf-8-sig') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("완료", "CSV 데이터를 JSON 파일에 성공적으로 저장했습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"CSV 파일을 읽는 중 오류가 발생했습니다: {e}")

# 테이블 데이터를 CSV로 내보내는 함수


def export_to_csv():
    file_path = filedialog.asksaveasfilename(
        defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if file_path:
        try:
            columns = ["Name", "Church", "Address", "Phone", "Email"]
            data = [table.item(item)["values"]
                    for item in table.get_children()]
            df = pd.DataFrame(data, columns=columns)
            df.to_csv(file_path, index=False, encoding='utf-8-sig')
            messagebox.showinfo("완료", "데이터를 CSV 파일로 성공적으로 내보냈습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"CSV 파일로 내보내는 중 오류가 발생했습니다: {e}")

# 메뉴바 생성 함수


def create_menubar(root):
    menubar = tk.Menu(root)

    # 파일 메뉴
    file_menu = tk.Menu(menubar, tearoff=0)
    file_menu.add_command(label="엑셀 가져오기", command=import_csv_to_json)
    file_menu.add_command(label="엑셀 내보내기", command=export_to_csv)
    menubar.add_cascade(label="파일", menu=file_menu)

    # 사용설명서 메뉴
    help_menu = tk.Menu(menubar, tearoff=0)
    help_menu.add_command(label="필독내용", command=lambda: messagebox.showinfo(
        "필독내용", "본 프로그램은 남경기노회 서기 행정을 위하여 특별히 제작되었기에 사용방법은 홈페이지 관리자에게 문의바랍니다."))
    menubar.add_cascade(label="사용설명서", menu=help_menu)

    root.config(menu=menubar)

# on_submit 함수 추가


def on_submit():
    presbytery_name = presbytery_entry.get()
    url = url_entry.get()
    if presbytery_name and url:
        crawl_and_save(presbytery_name, url)
    else:
        messagebox.showwarning("경고", "시찰명과 URL을 모두 입력하세요.")


# tkinter GUI 설정
root = tk.Tk()
root.title("남경기노회 회원주소록")

# 메뉴바 생성
create_menubar(root)

# 메뉴바 스타일 적용
style = ttk.Style()
style.configure("TMenu", font=("맑은 고딕", 12),
                background="#595456", foreground="#ffffff")
root.option_add("*TMenu.Font", "맑은 고딕 12")
root.option_add("*TMenu.foreground", "#ffffff")
root.option_add("*TMenu.background", "#595456")
root.option_add("*TMenu.padding", "5")

# 모니터 중앙에 창 위치시키기
window_width = 800
window_height = 600
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
position_top = int(screen_height / 2 - window_height / 2)
position_right = int(screen_width / 2 - window_width / 2)
root.geometry(f'{window_width}x{
              window_height}+{position_right}+{position_top}')

# 상단 섹션 프레임
top_frame = tk.Frame(root)
top_frame.pack(pady=10, padx=10, fill="x")

# 왼쪽 섹션 (시찰명, URL 입력 필드)
left_frame = tk.Frame(top_frame)
left_frame.pack(side="left", padx=10)

# 제목 레이블 추가
title_label = tk.Label(left_frame, text="남경기노회주소록만들기",
                       font=("맑은 고딕", 15), anchor="center")
title_label.grid(row=0, columnspan=2, pady=10)

presbytery_label = tk.Label(left_frame, text="시찰명:", font=(
    "맑은 고딕", 12), fg="#595456")  # 폰트 컬러 변경
presbytery_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")

presbytery_entry = tk.Entry(left_frame, width=50, fg="#595456")
presbytery_entry.insert(0, "시찰명을 입력하세요.")
presbytery_entry.grid(row=1, column=1, padx=5, pady=5)

url_label = tk.Label(left_frame, text="URL:", font=(
    "맑은 고딕", 12), fg="#595456")  # 폰트 컬러 변경
url_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")

url_entry = tk.Entry(left_frame, width=50, fg="#595456")
url_entry.insert(0, "URL을 입력하세요.")
url_entry.grid(row=2, column=1, padx=5, pady=5)
url_entry.bind('<Return>', lambda event: on_submit())  # 엔터키로도 실행되도록 설정

# 오른쪽 섹션 (로고 및 텍스트)
right_frame = tk.Frame(top_frame)
right_frame.pack(side="right", padx=10)

# 로고 이미지 삽입 (로고 이미지 파일의 경로를 지정해야 함)
logo_image = Image.open("logo.png")  # 'logo.png' 파일 경로로 변경
logo_image = logo_image.resize((100, 100))  # 필요한 경우 크기 조정
logo_image = ImageTk.PhotoImage(logo_image)
logo_label = tk.Label(right_frame, image=logo_image)
logo_label.pack(pady=10)

# 텍스트 삽입
text_label = tk.Label(right_frame, text="대한예수교장로회 남경기노회",
                      font=("맑은 고딕", 15), anchor="center")
text_label.pack(pady=5)

# 테이블 리스트 섹션
table_frame = tk.Frame(root)
table_frame.pack(pady=10, padx=10, fill="both", expand=True)

# 스크롤바 추가
scroll_y = tk.Scrollbar(table_frame, orient="vertical")
scroll_y.pack(side="right", fill="y")
scroll_x = tk.Scrollbar(table_frame, orient="horizontal")
scroll_x.pack(side="bottom", fill="x")

# 테이블 생성
columns = ("Name", "Church", "Address", "Phone", "Email")
table = ttk.Treeview(table_frame, columns=columns, show="headings",
                     yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
table.heading("Name", text="Name")
table.heading("Church", text="Church")
table.heading("Address", text="Address")
table.heading("Phone", text="Phone")
table.heading("Email", text="Email")
table.pack(fill="both", expand=True)

# 테이블 헤더 스타일 변경
style.configure("Treeview.Heading", background="#595456",
                foreground="#595456", font=("맑은 고딕", 10, "bold"))

# 시찰명 폰트컬러 변경
table.tag_configure('presbytery', foreground="#f78e48")

# 스크롤바 연결
scroll_y.config(command=table.yview)
scroll_x.config(command=table.xview)

# 하단 메세지
quote_label = tk.Label(root, text="made by 나종춘(2024)",
                       font=("맑은 고딕", 10), anchor="e", fg="#595456")
quote_label.pack(side="bottom", pady=10, padx=10, anchor="e")

# 메인 루프 시작
root.mainloop()
