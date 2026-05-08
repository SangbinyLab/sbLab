import xlwings as xw
import tkinter as tk
import json
import os

# --- 🎯 추가된 기능 1: 즐겨찾기 데이터 저장/불러오기 ---

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 그 폴더 경로 뒤에 'favorites.json'을 정확하게 붙여줍니다.
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
    root.title("Sheet Navigator")
    root.geometry("300x400")
    root.attributes('-topmost', True) 
    
    lbl_title = tk.Label(root, text="📑 시트 검색 및 즐겨찾기", font=("Arial", 14, "bold"))
    lbl_title.pack(pady=10)

    search_var = tk.StringVar()
    search_entry = tk.Entry(root, textvariable=search_var, font=("Arial", 12))
    search_entry.pack(fill=tk.X, padx=15, pady=(0, 10))
    search_entry.focus()

    list_frame = tk.Frame(root)
    list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

    scrollbar = tk.Scrollbar(list_frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=("Arial", 12), selectmode=tk.SINGLE)
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=listbox.yview)

    # --- 🎯 추가된 기능 2: 우클릭 팝업 메뉴 만들기 ---
    popup_menu = tk.Menu(root, tearoff=0)
    
    def toggle_favorite():
        selection = listbox.curselection()
        if not selection: return
        
        # 선택된 항목의 순수 시트 이름 추출
        sheet_name = listbox.get(selection[0]).replace("⭐ ", "")
        
        # 즐겨찾기 토글 (있으면 빼고, 없으면 넣기)
        if sheet_name in favorites:
            favorites.remove(sheet_name)
        else:
            favorites.add(sheet_name)
            
        save_favorites(favorites) # 파일에 즉시 저장
        update_listbox()          # 화면 즉시 새로고침

    popup_menu.add_command(label="⭐ 즐겨찾기 등록/해제", command=toggle_favorite)

    def show_context_menu(event):
        """우클릭 시 메뉴를 마우스 위치에 띄우는 함수"""
        # 우클릭한 위치의 리스트 항목을 자동으로 선택 상태로 만듭니다.
        nearest_index = listbox.nearest(event.y)
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(nearest_index)
        listbox.activate(nearest_index)
        
        # 팝업 메뉴 띄우기
        popup_menu.tk_popup(event.x_root, event.y_root)
    # ------------------------------------------------

    def update_listbox(*args):
        try:
            current_sheets = get_sheet_names()
            search_term = search_var.get().lower()
            
            if search_term:
                filtered = [s for s in current_sheets if search_term in s.lower()]
            else:
                filtered = current_sheets
                
            # 🎯 추가된 기능 3: 리스트 정렬 (즐겨찾기를 위로, 나머지를 아래로)
            fav_list = [s for s in filtered if s in favorites]
            normal_list = [s for s in filtered if s not in favorites]
            
            # 최종적으로 GUI에 표시할 텍스트 리스트 생성
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
            
    search_var.trace_add('write', update_listbox)
    listbox.bind('<Double-Button-1>', goto_selected_sheet) 
    listbox.bind('<Return>', goto_selected_sheet)
    search_entry.bind('<Return>', lambda event: goto_selected_sheet(tk.Event()) if listbox.curselection() else None)

    # 맥 OS는 우클릭이 <Button-2> 또는 <Button-3>으로 인식될 수 있어 둘 다 연결해 둡니다.
    listbox.bind('<Button-2>', show_context_menu)
    listbox.bind('<Button-3>', show_context_menu)

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
