import xlwings as xw
import tkinter as tk
import json
import os
from tkinter import messagebox
from PIL import Image, ImageTk

# --- 🎯 추가된 기능 1: 즐겨찾기 데이터 저장/불러오기 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FAV_FILE = os.path.join(BASE_DIR, "favorites.json")

def load_favorites():
    """저장된 즐겨찾기 목록을 불러옵니다."""
    if os.path.exists(FAV_FILE):
        try:
            with open(FAV_FILE, 'r', encoding='utf-8') as f:
                return set(json.load(f)) # 빠른 검색과 중복 방지를 위해 Set 자료형 사용
        except:
            return set()
    return set()

def save_favorites(fav_set):
    """즐겨찾기 목록을 파일로 저장합니다."""
    with open(FAV_FILE, 'w', encoding='utf-8') as f:
        json.dump(list(fav_set), f, ensure_ascii=False)

# 전역 변수로 즐겨찾기 목록 관리
favorites = load_favorites()
# --------------------------------------------------------

def get_sheet_names():
    try:
        wb = xw.books.active 
        return [sheet.name for sheet in wb.sheets]
    except Exception as e:
        return []

def goto_selected_sheet(event):
    listbox = event.widget
    selection = listbox.curselection()
    if not selection:
        return 

    index = selection[0]
    raw_text = listbox.get(index)
    
    # 🎯 중요: 실제 시트로 이동할 때는 "⭐ " 기호를 떼고 넘겨줘야 합니다.
    sheet_name = raw_text.replace("⭐ ", "")
    
    try:
        wb = xw.books.active
        wb.sheets[sheet_name].activate() 
    except Exception as e:
        print(f"시트 이동 중 에러 발생: {e}")

def show_gui():
    root = tk.Tk()
    root.title("Sangbin's LAB")
    root.geometry("300x400")
    root.attributes('-topmost', True) 
    
    # 1. 헤더 프레임 (타이틀과 로고)
    header_frame = tk.Frame(root)
    header_frame.pack(fill=tk.X, padx=15, pady=10)

    # 2. 🎯 텍스트 타이틀 대신 'vivid.png' 이미지 불러오기 (왼쪽 배치)
    vivid_path = os.path.join(BASE_DIR, "vivid1.png")
    try:
        pil_vivid = Image.open(vivid_path)
        # 가로로 긴 텍스트 이미지이므로 가로 200, 세로 40 정도로 최대 크기 제한
        pil_vivid.thumbnail((210, 50), Image.Resampling.LANCZOS)
        img_vivid = ImageTk.PhotoImage(pil_vivid)
        
        lbl_vivid = tk.Label(header_frame, image=img_vivid)
        lbl_vivid.image = img_vivid # 가비지 컬렉터 방지
        lbl_vivid.pack(side=tk.LEFT, pady=(5,0)) 
    except Exception as e:
        print("Vivid 타이틀 이미지를 찾을 수 없습니다:", e)

    # 3. 로고 이미지 불러오기 (오른쪽 배치)
    logo_path = os.path.join(BASE_DIR, "logo.png")
    try:
        pil_logo = Image.open(logo_path)
        pil_logo.thumbnail((50, 50), Image.Resampling.LANCZOS)
        img_logo = ImageTk.PhotoImage(pil_logo)
        
        lbl_logo = tk.Label(header_frame, image=img_logo)
        lbl_logo.image = img_logo # 가비지 컬렉터 방지
        lbl_logo.pack(side=tk.RIGHT)
        
    except Exception as e:
        error_details = f"이미지를 불러오는 데 실패했습니다.\n\n[확인된 경로]\n{logo_path}\n\n[에러 내용]\n{e}"
        messagebox.showwarning("로고 로드 에러", error_details)
    
    # --- 🎯 4. 검색 영역 프레임 분리 (Search 글자와 입력창을 한 줄에 배치) ---
    search_frame = tk.Frame(root)
    search_frame.pack(fill=tk.X, padx=15, pady=(0, 10))
    
    # "Search" 글자를 입력창 왼쪽에 배치
    lbl_search = tk.Label(search_frame, text="Search", font=("Arial", 15, "bold"))
    lbl_search.pack(side=tk.LEFT, padx=(0, 5)) # 입력창과의 간격을 위해 오른쪽(padx) 여백 5 추가
    
    # 검색 입력창 설정 (expand=True를 통해 남은 가로 공간을 모두 채움)
    search_var = tk.StringVar()
    search_entry = tk.Entry(search_frame, textvariable=search_var, font=("Arial", 12))
    search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
    search_entry.focus()
    # -------------------------------------------------------------------------

    # 5. 리스트 영역
    list_frame = tk.Frame(root)
    list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

    scrollbar = tk.Scrollbar(list_frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=("Arial", 12), selectmode=tk.SINGLE)
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=listbox.yview)

    # --- 우클릭 팝업 메뉴 만들기 ---
    popup_menu = tk.Menu(root, tearoff=0)
    
    def toggle_favorite():
        selection = listbox.curselection()
        if not selection: return
        
        sheet_name = listbox.get(selection[0]).replace("⭐ ", "")
        
        if sheet_name in favorites:
            favorites.remove(sheet_name)
        else:
            favorites.add(sheet_name)
            
        save_favorites(favorites)
        update_listbox()

    popup_menu.add_command(label="⭐ 즐겨찾기 등록/해제", command=toggle_favorite)

    def show_context_menu(event):
        nearest_index = listbox.nearest(event.y)
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(nearest_index)
        listbox.activate(nearest_index)
        popup_menu.tk_popup(event.x_root, event.y_root)

    # --- 리스트 업데이트 로직 ---
    def update_listbox(*args):
        try:
            current_sheets = get_sheet_names()
            search_term = search_var.get().lower()
            
            if search_term:
                filtered = [s for s in current_sheets if search_term in s.lower()]
            else:
                filtered = current_sheets
                
            fav_list = [s for s in filtered if s in favorites]
            normal_list = [s for s in filtered if s not in favorites]
            
            display_list = ["⭐ " + s for s in fav_list] + normal_list
            
            gui_sheets = list(listbox.get(0, tk.END))
            if display_list != gui_sheets:
                listbox.delete(0, tk.END)
                for item in display_list:
                    listbox.insert(tk.END, item)
                    
                if display_list:
                    listbox.selection_set(0)
        except:
            pass
            
    # 이벤트 바인딩
    search_var.trace_add('write', update_listbox)
    listbox.bind('<Double-Button-1>', goto_selected_sheet) 
    listbox.bind('<Return>', goto_selected_sheet)
    search_entry.bind('<Return>', lambda event: goto_selected_sheet(tk.Event()) if listbox.curselection() else None)

    listbox.bind('<Button-2>', show_context_menu)
    listbox.bind('<Button-3>', show_context_menu)

    # 폴링 타이머
    def polling_timer():
        update_listbox()
        root.after(1000, polling_timer)
        
    polling_timer()
    root.mainloop()

def main():
    show_gui()

if __name__ == "__main__":
    xw.Book("testwb.xlsm").set_mock_caller()
    main()
