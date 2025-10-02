import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import re
import sqlite3
import os


# QuestionFrame 클래스 (PanedWindow 및 UI 레이아웃 수정)
class QuestionFrame(ttk.Frame):
    """자소서 문항 하나에 대한 입력 필드와 글자수 측정 기능을 제공하는 프레임"""

    def __init__(self, parent, question_number, initial_title=None, initial_data=None):
        super().__init__(parent, padding="10")
        self.question_number = question_number

        default_title = initial_title if initial_title else f"문항 {self.question_number}"
        self.title_var = tk.StringVar(value=default_title)

        self.create_widgets()

        if initial_data:
            self._load_initial_data(initial_data)

        self.answer_text.bind('<KeyRelease>', self.update_char_count)
        self.update_char_count()

    def _load_initial_data(self, data):
        """저장된 데이터를 기반으로 위젯의 내용을 채웁니다."""
        self.question_text.delete("1.0", tk.END)
        self.question_text.insert("1.0", data.get("질문", ""))

        self.type_entry.delete(0, tk.END)
        self.type_entry.insert(0, data.get("문항유형", ""))

        self.answer_text.delete("1.0", tk.END)
        self.answer_text.insert("1.0", data.get("답변", ""))

    def update_question_number(self, new_number):
        """문항 번호와 UI 제목을 업데이트합니다."""
        self.question_number = new_number

        current_title = self.title_var.get()
        match = re.fullmatch(r"문항 (\d+)", current_title)

        if match:
            self.title_var.set(f"문항 {self.question_number}")

    def update_title(self, new_title):
        """실제로 문항 제목을 업데이트하고 탭 이름도 업데이트합니다."""
        if new_title:
            self.title_var.set(new_title)

            try:
                notebook = self.master
                for tab_id in notebook.tabs():
                    if notebook.nametowidget(tab_id) is self:
                        notebook.tab(tab_id, text=new_title)
                        break
            except Exception:
                pass

    def open_title_edit_popup(self):
        """제목을 수정하는 팝업 창을 엽니다."""
        popup = tk.Toplevel(self)
        popup.title("문항 제목 수정")
        popup.transient(self.master.master)
        popup.grab_set()
        popup.geometry("350x150")

        popup_frame = ttk.Frame(popup, padding="15")
        popup_frame.pack(expand=True, fill="both")

        ttk.Label(popup_frame, text="새 문항 제목:", font=('Arial', 10, 'bold')).pack(pady=(0, 5), anchor="w")

        new_title_var = tk.StringVar(value=self.title_var.get())
        title_entry = ttk.Entry(popup_frame, textvariable=new_title_var, width=40, font=('Arial', 10))
        title_entry.pack(pady=(0, 10), fill="x", expand=True)
        title_entry.focus_set()
        title_entry.icursor(tk.END)

        def on_confirm(event=None):
            new_title = new_title_var.get().strip()
            if new_title:
                self.update_title(new_title)
            popup.destroy()

        def on_cancel(event=None):
            popup.destroy()

        title_entry.bind('<Return>', on_confirm)
        popup.bind('<Escape>', on_cancel)

        button_frame = ttk.Frame(popup_frame)
        button_frame.pack(fill="x")

        ttk.Button(button_frame, text="확인", command=on_confirm).pack(side="right", padx=5)
        ttk.Button(button_frame, text="취소", command=on_cancel).pack(side="right")

        popup.protocol("WM_DELETE_WINDOW", on_cancel)
        self.master.master.wait_window(popup)

    # 헬퍼 함수: Text 위젯과 Scrollbar를 생성하고 확장 설정
    def _create_text_with_scrollbar(self, parent_frame, height):
        """Text 위젯, Scrollbar를 하나의 Wrapper Frame 안에 배치하고 Text 위젯을 반환합니다.
        Wrapper Frame은 이미 상위 프레임에 grid 또는 pack 되어있어야 합니다."""

        text_widget = tk.Text(parent_frame, height=height, wrap='word', font=('Arial', 10))
        text_widget.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(parent_frame, command=text_widget.yview)
        text_widget['yscrollcommand'] = scrollbar.set
        scrollbar.grid(row=0, column=1, sticky='ns')

        # Wrapper Frame (parent_frame) 내부 설정: Text 위젯이 확장되도록
        parent_frame.grid_columnconfigure(0, weight=1)
        parent_frame.grid_rowconfigure(0, weight=1)

        return text_widget

    def create_widgets(self):
        # 1. 문항 제목 섹션 (R0)
        title_control_frame = ttk.Frame(self)
        title_control_frame.grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky="ew")

        self.title_entry = ttk.Entry(
            title_control_frame,
            textvariable=self.title_var,
            font=('Arial', 14, 'bold'),
            state='readonly',
            justify='left'
        )
        self.title_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

        edit_button = ttk.Button(
            title_control_frame,
            text="수정",
            command=self.open_title_edit_popup
        )
        edit_button.pack(side="left")

        # PanedWindow가 R1을 전체 차지하도록 설정
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # 2. PanedWindow: 질문 섹션과 (문항 유형 + 답변) 섹션을 분리 (R1)
        main_pane = ttk.PanedWindow(self, orient=tk.VERTICAL)
        main_pane.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")

        # 2.1 상단 프레임: 질문 텍스트만 포함 (question_pane_frame)
        question_pane_frame = ttk.Frame(main_pane)
        question_pane_frame.grid_columnconfigure(0, weight=1)
        question_pane_frame.grid_rowconfigure(1, weight=1)

        # 2.2 하단 프레임: 문항 유형 + 답변 포함 (details_pane_frame)
        details_pane_frame = ttk.Frame(main_pane)
        details_pane_frame.grid_columnconfigure(0, weight=1)
        details_pane_frame.grid_rowconfigure(2, weight=1)

        # PanedWindow에 프레임 추가 (weight로 초기 비율 설정)
        main_pane.add(question_pane_frame, weight=1)  # 질문 섹션 (초기 1/4)
        main_pane.add(details_pane_frame, weight=3)  # 유형/답변 섹션 (초기 3/4)

        # A. 질문 섹션 (question_pane_frame)
        # A.1 레이블 (R0)
        ttk.Label(question_pane_frame, text="질문:", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, padx=5, pady=5, sticky="nw"
        )
        # A.2 텍스트 위젯 (R1) - Text와 Scrollbar를 담는 Wrapper Frame
        wrapper_frame_q = ttk.Frame(question_pane_frame)
        wrapper_frame_q.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        self.question_text = self._create_text_with_scrollbar(wrapper_frame_q, height=2)

        # B. 유형/답변 섹션 (details_pane_frame)

        # B.1 문항 유형 (R0)
        type_frame = ttk.Frame(details_pane_frame)
        type_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        type_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(type_frame, text="문항 유형:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky="w", padx=5)
        self.type_entry = ttk.Entry(type_frame)
        self.type_entry.grid(row=0, column=1, sticky="ew", padx=5)

        # B.2 답변 레이블 (R1)
        ttk.Label(details_pane_frame, text="답변:", font=('Arial', 10, 'bold')).grid(
            row=1, column=0, padx=5, pady=(15, 5), sticky="nw"
        )
        # B.3 답변 텍스트 위젯 (R2) - Text와 Scrollbar를 담는 Wrapper Frame
        wrapper_frame_a = ttk.Frame(details_pane_frame)
        wrapper_frame_a.grid(row=2, column=0, padx=5, pady=5, sticky="nsew")
        self.answer_text = self._create_text_with_scrollbar(wrapper_frame_a, height=15)

        # 3. 글자수 측정 (R2) - 단일 tk.Text 위젯으로 색상/크기 적용
        # borderwidth=0, relief="flat"으로 테두리 제거. background 설정으로 배경색 일치.
        self.count_display = tk.Text(self, height=1, wrap='word', font=('Arial', 10), state='disabled',
                                     borderwidth=0, relief="flat", foreground='gray40')
        self.count_display.grid(row=2, column=0, columnspan=2, padx=5, pady=(5, 0), sticky="ew")

        # 태그 설정: 숫자만 크게, 색깔 다르게
        self.count_display.tag_config('count_all', foreground='#00008B', font=('Arial', 12, 'bold'))  # 진한 파란색
        self.count_display.tag_config('count_no_space', foreground='#0A7959', font=('Arial', 12, 'bold'))  # 진한 녹색
        self.count_display.tag_config('normal', font=('Arial', 10, 'normal'))


    def update_char_count(self, event=None):
        """답변 텍스트를 분석하여 글자수를 실시간으로 업데이트하고, tk.Text에 태그를 적용하여 표시합니다."""

        text_content = self.answer_text.get("1.0", tk.END)
        if text_content.endswith('\n'):
            text_content = text_content[:-1]

        char_count_all = len(text_content)
        char_count_no_space = len(text_content.replace(' ', ''))

        # tk.Text를 사용하여 태그로 스타일 적용
        self.count_display.config(state='normal')
        self.count_display.delete("1.0", tk.END)

        # 1. 띄어쓰기 포함 글자수
        msg1_prefix = "총 글자수 (띄어쓰기 포함): "
        self.count_display.insert(tk.END, msg1_prefix, 'normal')
        self.count_display.insert(tk.END, str(char_count_all), 'count_all')

        # 2. 구분자 및 띄어쓰기 제외 글자수
        msg2_prefix = "자 | 제외 (개행 포함): "
        self.count_display.insert(tk.END, msg2_prefix, 'normal')
        self.count_display.insert(tk.END, str(char_count_no_space), 'count_no_space')
        self.count_display.insert(tk.END, "자", 'normal')

        self.count_display.config(state='disabled')

    def get_data(self):
        """이 문항 프레임의 데이터를 딕셔너리로 반환합니다. (제목 포함)"""

        question_content = self.question_text.get("1.0", tk.END).strip()
        answer_content = self.answer_text.get("1.0", tk.END).strip()

        return {
            "제목": self.title_var.get(),
            "질문": question_content,
            "문항유형": self.type_entry.get().strip(),
            "답변": answer_content,
        }


# Application 클래스: 메인 윈도우와 전체 로직을 정의합니다.
class Application(tk.Tk):
    MAX_QUESTIONS = 20

    def __init__(self):
        super().__init__()
        self.title("자소서 문항 정리 및 저장 애플리케이션 (UI 개선)")
        self.geometry("1100x700")

        self.all_companies_data = {}
        self.current_company_name = None
        self.question_counter = 0

        # --- 새로운 상태 변수: 마지막으로 저장/불러온 파일 경로 ---
        self.last_save_path = None

        self.create_menu_bar()
        self.create_widgets()

        #self.add_new_company("새 회사 1")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        """앱 종료 전 현재 편집 중인 내용을 저장합니다."""
        if self.current_company_name:
            self.save_current_company_data()
        self.destroy()

    def create_menu_bar(self):
        """메뉴 바를 생성하고 새로운 파일 관리 기능을 추가합니다."""
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        # 1. 파일 메뉴 (저장/불러오기)
        file_menu = tk.Menu(menubar, tearoff=0)

        # --- 불러오기/추출하기 ---
        file_menu.add_command(label="텍스트 파일 불러오기", command=self.load_text_file)
        file_menu.add_command(label="SQL 파일로부터 추출하기", command=self.load_from_sql_file)
        file_menu.add_separator()

        # --- 저장하기 ---
        # 1. 현재 회사 저장
        file_menu.add_command(label="현재 회사 저장 (단일 텍스트)", command=self.save_current_company_to_file, state=tk.DISABLED)
        # 3. 전체 회사를 저장
        file_menu.add_command(label="전체 회사 저장", command=self.save_all_companies, state=tk.DISABLED)
        # 2. 전체 회사를 다른 이름으로 저장
        file_menu.add_command(label="전체 회사를 다른 이름으로 저장", command=self.save_all_companies_as, state=tk.DISABLED)
        file_menu.add_separator()

        # 4. SQL 파일로 내보내기
        file_menu.add_command(label="SQL 파일로 내보내기", command=self.export_to_sql, state=tk.DISABLED)
        file_menu.add_separator()

        file_menu.add_command(label="종료", command=self.on_closing)
        menubar.add_cascade(label="파일", menu=file_menu)

        # 2. 도구 메뉴 (검색)
        tool_menu = tk.Menu(menubar, tearoff=0)
        tool_menu.add_command(label="전체 문항 검색", command=self.open_search_popup)
        menubar.add_cascade(label="도구", menu=tool_menu)

        # 메뉴바 저장 버튼 상태를 외부에서 접근할 수 있도록 저장 (인덱스 변경됨)
        self.menu_save_current = file_menu.entrycget(3, "label")
        self.menu_save_all = file_menu.entrycget(4, "label")
        self.menu_save_all_as = file_menu.entrycget(5, "label")
        self.menu_export_sql = file_menu.entrycget(7, "label")
        self.file_menu = file_menu

    def create_widgets(self):
        # UI 생성 로직
        paned_window = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned_window.pack(fill="both", expand=True, padx=10, pady=10)

        left_frame = ttk.Frame(paned_window, width=280, padding="5")
        left_frame.pack_propagate(False)

        tree_container = ttk.Frame(left_frame)
        tree_container.pack(fill="both", expand=True)

        tree_scroll = ttk.Scrollbar(tree_container)
        tree_scroll.pack(side="right", fill="y")

        self.company_tree = ttk.Treeview(
            tree_container,
            columns=("Company"),
            show="headings",
            yscrollcommand=tree_scroll.set,
            selectmode='browse'
        )
        tree_scroll.config(command=self.company_tree.yview)

        self.company_tree.heading("Company", text="회사 이름")
        self.company_tree.column("Company", width=250, stretch=tk.YES)

        self.company_tree.pack(side="left", fill="both", expand=True)
        self.company_tree.bind('<<TreeviewSelect>>', self.load_company_data)

        company_control_frame = ttk.Frame(left_frame)
        company_control_frame.pack(fill="x", pady=(10, 0))

        # 회사 추가/제거 버튼 레이블 변경
        ttk.Button(company_control_frame, text="╋회사추가", command=self.add_new_company_popup).pack(side="left", fill="x",
                                                                                                 expand=True,
                                                                                                 padx=(0, 5))

        self.remove_company_button = ttk.Button(company_control_frame, text="━회사제거",
                                                command=self.remove_current_company, state=tk.DISABLED)
        self.remove_company_button.pack(side="left", fill="x", expand=True)

        paned_window.add(left_frame, weight=0)

        right_frame = ttk.Frame(paned_window)

        company_display_frame = ttk.Frame(right_frame)
        company_display_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(company_display_frame, text="회사명:", font=('Arial', 12, 'bold')).pack(side="left", padx=5)
        self.current_company_name_var = tk.StringVar(value="회사를 선택하거나 추가해주세요.")

        self.current_company_name_entry = ttk.Entry(
            company_display_frame,
            textvariable=self.current_company_name_var,
            font=('Arial', 14, 'bold'),
            state='readonly',
            justify='left'
        )
        self.current_company_name_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.edit_company_name_button = ttk.Button(
            company_display_frame,
            text="수정",
            command=self.open_company_name_edit_popup,
            state=tk.DISABLED
        )
        self.edit_company_name_button.pack(side="left")

        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(fill="both", expand=True)

        control_frame = ttk.Frame(right_frame)
        control_frame.pack(fill="x", pady=10)

        question_control_frame = ttk.Frame(control_frame)
        question_control_frame.pack(side="left")

        # 문항 추가/제거 버튼 레이블 변경
        self.add_button = ttk.Button(question_control_frame, text="╋문항추가", command=self.add_question_tab,
                                     state=tk.DISABLED)
        self.add_button.pack(side="left", padx=5)

        self.remove_button = ttk.Button(question_control_frame, text="━문항제거", command=self.remove_question_tab,
                                        state=tk.DISABLED)
        self.remove_button.pack(side="left", padx=5)

        paned_window.add(right_frame, weight=1)

    def _update_treeview(self):
        """데이터를 기반으로 Treeview를 새로고침하고 저장 버튼 상태를 업데이트합니다."""

        num_companies = len(self.all_companies_data)

        for item in self.company_tree.get_children():
            self.company_tree.delete(item)

        for company in sorted(self.all_companies_data.keys()):
            self.company_tree.insert("", tk.END, values=(company,), iid=company)

        # 메뉴바 저장 버튼 상태 업데이트
        save_state = tk.NORMAL if num_companies > 0 else tk.DISABLED
        self.file_menu.entryconfig(self.menu_save_current, state=save_state)
        self.file_menu.entryconfig(self.menu_save_all, state=save_state)
        self.file_menu.entryconfig(self.menu_save_all_as, state=save_state)
        self.file_menu.entryconfig(self.menu_export_sql, state=save_state)

    def _set_controls_state(self, state):
        """문항 및 회사명 관련 제어 버튼의 상태를 설정합니다."""
        state = tk.NORMAL if state else tk.DISABLED
        self.add_button.config(state=state)
        self.remove_button.config(state=state)
        self.edit_company_name_button.config(state=state)
        self.remove_company_button.config(state=state)

        # 현재 회사가 선택되면 단일 저장 버튼을 활성화
        current_save_state = tk.NORMAL if self.current_company_name else tk.DISABLED
        self.file_menu.entryconfig(self.menu_save_current, state=current_save_state)

    def _clear_notebook(self):
        """현재 Notebook의 모든 탭을 제거합니다."""
        for tab in self.notebook.tabs():
            self.notebook.forget(tab)
        self.question_counter = 0

    def add_new_company(self, company_name):
        """새 회사 데이터를 추가하고 목록을 업데이트합니다."""
        if company_name in self.all_companies_data:
            messagebox.showwarning("중복", f"'{company_name}'은(는) 이미 회사 목록에 존재합니다.")
            return

        self.all_companies_data[company_name] = []
        self._update_treeview()

        self.company_tree.selection_set(company_name)
        self.company_tree.focus(company_name)

        self.load_company_data(None)

    def add_new_company_popup(self):
        """회사 이름을 입력받는 팝업 창을 엽니다."""
        popup = tk.Toplevel(self)
        popup.title("새 회사 추가")
        popup.transient(self)
        popup.grab_set()
        popup.geometry("300x120")

        popup_frame = ttk.Frame(popup, padding="15")
        popup_frame.pack(expand=True, fill="both")

        ttk.Label(popup_frame, text="추가할 회사 이름:").pack(pady=(0, 5), anchor="w")

        company_name_var = tk.StringVar()
        name_entry = ttk.Entry(popup_frame, textvariable=company_name_var, width=30)
        name_entry.pack(pady=(0, 10), fill="x", expand=True)
        name_entry.focus_set()
        name_entry.icursor(tk.END)

        def on_confirm(event=None):
            name = company_name_var.get().strip()
            if name:
                self.add_new_company(name)
                popup.destroy()
            else:
                messagebox.showwarning("입력 필요", "회사 이름을 입력해주세요.")

        def on_cancel(event=None):
            popup.destroy()

        name_entry.bind('<Return>', on_confirm)
        popup.bind('<Escape>', on_cancel)

        button_frame = ttk.Frame(popup_frame)
        button_frame.pack(fill="x")

        ttk.Button(button_frame, text="확인", command=on_confirm).pack(side="right", padx=5)
        ttk.Button(button_frame, text="취소", command=on_cancel).pack(side="right")

        self.wait_window(popup)

    def open_company_name_edit_popup(self):
        """회사 이름을 수정하는 팝업 창을 엽니다."""
        if not self.current_company_name:
            return

        popup = tk.Toplevel(self)
        popup.title("회사 이름 수정")
        popup.transient(self)
        popup.grab_set()
        popup.geometry("350x150")

        popup_frame = ttk.Frame(popup, padding="15")
        popup_frame.pack(expand=True, fill="both")

        ttk.Label(popup_frame, text="새 회사 이름:", font=('Arial', 10, 'bold')).pack(pady=(0, 5), anchor="w")

        new_name_var = tk.StringVar(value=self.current_company_name)
        name_entry = ttk.Entry(popup_frame, textvariable=new_name_var, width=40, font=('Arial', 10))
        name_entry.pack(pady=(0, 10), fill="x", expand=True)
        name_entry.focus_set()
        name_entry.icursor(tk.END)

        def on_confirm(event=None):
            new_name = new_name_var.get().strip()
            if not new_name:
                messagebox.showwarning("입력 오류", "회사 이름은 공백일 수 없습니다.")
            elif new_name == self.current_company_name:
                messagebox.showinfo("변경 없음", "회사 이름이 이전과 동일합니다.")
                popup.destroy()
            else:
                self.rename_current_company(new_name)
                popup.destroy()

        def on_cancel(event=None):
            popup.destroy()

        name_entry.bind('<Return>', on_confirm)
        popup.bind('<Escape>', on_cancel)

        button_frame = ttk.Frame(popup_frame)
        button_frame.pack(fill="x")

        ttk.Button(button_frame, text="확인", command=on_confirm).pack(side="right", padx=5)
        ttk.Button(button_frame, text="취소", command=on_cancel).pack(side="right")

        self.wait_window(popup)

    def rename_current_company(self, new_name):
        """내부 데이터와 UI에서 현재 회사의 이름을 변경합니다."""

        if new_name in self.all_companies_data and new_name != self.current_company_name:
            messagebox.showwarning("이름 중복", f"회사명 '{new_name}'이(가) 이미 존재합니다.")
            return

        old_name = self.current_company_name

        try:
            self.save_current_company_data()
            self.all_companies_data[new_name] = self.all_companies_data.pop(old_name)
            self.current_company_name = new_name

            self.current_company_name_var.set(new_name)
            self._update_treeview()

            self.company_tree.selection_set(new_name)
            self.company_tree.focus(new_name)

            messagebox.showinfo("수정 완료", f"회사 이름이 '{old_name}'에서 '{new_name}'(으)로 변경되었습니다.")

        except Exception as e:
            messagebox.showerror("이름 변경 오류", f"회사 이름 변경 중 오류가 발생했습니다: {e}")
            self.current_company_name = old_name
            self.current_company_name_var.set(old_name)
            self._update_treeview()

    def remove_current_company(self):
        """현재 선택된 회사를 목록에서 제거합니다."""
        if not self.current_company_name:
            return

        company_to_remove = self.current_company_name

        confirm = messagebox.askyesno(
            "회사 제거 확인",
            f"회사 '{company_to_remove}'와(과) 관련된 모든 문항 데이터가 영구적으로 제거됩니다.\n정말 제거하시겠습니까?"
        )

        if confirm:
            try:
                del self.all_companies_data[company_to_remove]

                self.current_company_name = None
                self.current_company_name_var.set("회사를 선택하거나 추가해주세요.")
                self._clear_notebook()
                self._set_controls_state(False)
                self._update_treeview()

                messagebox.showinfo("제거 완료", f"회사 '{company_to_remove}'가(이) 성공적으로 제거되었습니다.")

                if self.company_tree.get_children():
                    first_company_id = self.company_tree.get_children()[0]
                    self.company_tree.selection_set(first_company_id)
                    self.company_tree.focus(first_company_id)
                    self.load_company_data(None)

            except Exception as e:
                messagebox.showerror("제거 오류", f"회사 제거 중 오류가 발생했습니다: {e}")

    def save_current_company_data(self):
        """현재 Notebook에 표시된 문항 내용을 내부 데이터에 저장합니다."""
        if not self.current_company_name:
            return

        current_questions = []
        for tab_id in self.notebook.tabs():
            try:
                frame = self.nametowidget(tab_id)
                if isinstance(frame, QuestionFrame):
                    current_questions.append(frame.get_data())
            except KeyError:
                continue

        self.all_companies_data[self.current_company_name] = current_questions

    def load_company_data(self, event):
        """Treeview에서 새 회사가 선택되면 데이터를 로드합니다."""

        selected_item_id = self.company_tree.focus()
        if not selected_item_id:
            self._set_controls_state(False)
            return

        new_company_name = self.company_tree.item(selected_item_id, 'values')[0]

        if new_company_name == self.current_company_name:
            return

        if self.current_company_name:
            self.save_current_company_data()

        self.current_company_name = new_company_name
        self.current_company_name_var.set(new_company_name)

        self._clear_notebook()

        questions_data = self.all_companies_data.get(new_company_name, [])
        if questions_data:
            for data in questions_data:
                self.add_question_tab(initial_data=data, initial_title=data.get('제목'))
        else:
            self.add_question_tab()

        self._set_controls_state(True)

    def add_question_tab(self, initial_data=None, initial_title=None):
        """새로운 문항 탭을 추가하고 데이터를 로드합니다."""
        if not self.current_company_name:
            messagebox.showwarning("선택 오류", "먼저 편집할 회사를 선택하거나 추가해주세요.")
            return

        if len(self.notebook.tabs()) >= self.MAX_QUESTIONS:
            messagebox.showwarning("제한 초과", f"문항은 최대 {self.MAX_QUESTIONS}개까지만 추가할 수 있습니다.")
            return

        self.question_counter = len(self.notebook.tabs()) + 1  # 실제 탭 개수 기반으로 카운트

        frame = QuestionFrame(self.notebook, self.question_counter, initial_title, initial_data)

        tab_name = frame.title_var.get()
        self.notebook.add(frame, text=tab_name)

        self.notebook.select(frame)

    def remove_question_tab(self):
        """현재 선택된 문항 탭을 제거하고, 남은 탭들의 번호를 재조정합니다."""
        if not self.notebook.tabs():
            return

        selected_tab_id = self.notebook.select()
        if not selected_tab_id:
            return

        current_frame = self.nametowidget(selected_tab_id)
        data = current_frame.get_data()

        content_is_empty = not (data['질문'] or data['답변'] or data['문항유형'])

        should_remove = True
        if not content_is_empty:
            should_remove = messagebox.askyesno(
                "문항 제거 확인",
                f"문항 '{current_frame.title_var.get()}'에 작성된 내용이 있습니다.\n정말 제거하시겠습니까?"
            )

        if should_remove:
            self.notebook.forget(selected_tab_id)

            # 남은 탭들의 번호 및 제목 업데이트
            for i, tab_id in enumerate(self.notebook.tabs()):
                frame = self.nametowidget(tab_id)
                new_number = i + 1

                if isinstance(frame, QuestionFrame):
                    frame.update_question_number(new_number)
                    self.notebook.tab(tab_id, text=frame.title_var.get())

            self.save_current_company_data()

    # --- 텍스트 파일 입출력 로직 (변경 없음) ---

    def _parse_file_content(self, content):
        """구조화된 파일 내용을 파싱하여 {회사명: [문항 데이터 리스트]} 형식으로 반환합니다."""
        parsed_data = {}
        current_company = None
        current_question = None

        lines = content.strip().split('\n')

        in_answer_section = False

        for line in lines:
            line = line.strip()

            if not line and current_question:
                # 내용(질문/답변) 섹션에서 빈 줄은 포함
                if not in_answer_section:
                    current_question['질문'] += '\n'
                else:
                    current_question['답변'] += '\n'
                continue

            if line.startswith('[회사명]:'):
                current_company = line.split(':', 1)[1].strip()
                in_answer_section = False
                current_question = None

                if current_company not in parsed_data:
                    parsed_data[current_company] = []

            elif line == '--- 문항 시작 ---':
                current_question = {"제목": "제목 없음", "질문": "", "답변": "", "문항유형": ""}
                in_answer_section = False

            elif line.startswith('<<제목>>:') and current_question:
                current_question['제목'] = line.split(':', 1)[1].strip()
            elif line.startswith('<<유형>>:') and current_question:
                current_question['문항유형'] = line.split(':', 1)[1].strip()

            elif line == '<<질문>>' and current_question:
                in_answer_section = False
            elif line == '<<답변>>' and current_question:
                in_answer_section = True

            elif line == '--- 문항 끝 ---' and current_company and current_question:
                parsed_data[current_company].append(current_question)
                current_question = None
                in_answer_section = False

            elif current_question:
                # 질문/답변 내용 추가 (공백 줄은 위에서 처리했으므로 내용만 추가)
                if not in_answer_section:
                    current_question['질문'] += line + '\n'
                else:
                    current_question['답변'] += line + '\n'

        # 최종 정리 및 끝 공백/개행 제거
        for company, questions in parsed_data.items():
            for q in questions:
                q['질문'] = q['질문'].strip()
                q['답변'] = q['답변'].strip()

        return parsed_data

    def _format_data(self, company_name=None):
        """특정 회사(company_name) 또는 전체 회사 데이터를 구조화된 텍스트 형식으로 포맷합니다."""

        # 항상 현재 작업 내용을 저장
        self.save_current_company_data()

        formatted_text = ""

        # 저장할 회사 목록 결정
        if company_name and company_name in self.all_companies_data:
            companies_to_save = {company_name: self.all_companies_data[company_name]}
        else:
            companies_to_save = self.all_companies_data

        for name, questions in companies_to_save.items():

            formatted_text += f"[회사명]: {name}\n"

            for data in questions:
                formatted_text += "--- 문항 시작 ---\n"
                formatted_text += f"<<제목>>: {data.get('제목', '제목 없음')}\n"
                formatted_text += f"<<유형>>: {data.get('문항유형', '')}\n"

                formatted_text += "<<질문>>\n"
                formatted_text += f"{data.get('질문', '')}\n"

                formatted_text += "<<답변>>\n"
                formatted_text += f"{data.get('답변', '')}\n"

                formatted_text += "--- 문항 끝 ---\n"

            formatted_text += "=== 회사 끝 ===\n\n"

        return formatted_text.strip()

    # 1. 현재 회사 저장 (단일 텍스트)
    def save_current_company_to_file(self):
        """현재 선택된 회사 데이터만 텍스트 파일로 저장합니다."""
        if not self.current_company_name:
            messagebox.showwarning("저장 불가", "먼저 저장할 회사를 선택해주세요.")
            return

        formatted_data = self._format_data(company_name=self.current_company_name)

        initial_filename = f"{self.current_company_name}_문항.txt"

        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            initialfile=initial_filename,
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            title=f"'{self.current_company_name}' 문항 내용을 저장합니다."
        )

        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(formatted_data)
                messagebox.showinfo("저장 완료", f"'{self.current_company_name}' 데이터가 성공적으로 저장되었습니다:\n{file_path}")
            except Exception as e:
                messagebox.showerror("저장 오류", f"파일 저장 중 오류가 발생했습니다: {e}")

    # 2. 전체 회사를 다른 이름으로 저장 (기존 new_save 대체)
    def save_all_companies_as(self):
        """모든 회사 데이터를 새 파일에 저장하고, last_save_path를 업데이트합니다."""
        if not self.all_companies_data:
            messagebox.showwarning("저장 불가", "저장할 회사 데이터가 없습니다.")
            return

        formatted_data = self._format_data()

        initial_filename = "자소서_통합본.txt"
        if self.current_company_name:
            initial_filename = f"{self.current_company_name}_통합본.txt"

        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            initialfile=initial_filename,
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            title="모든 회사 문항 내용을 새 파일로 저장합니다."
        )

        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(formatted_data)

                # 저장 성공 시 경로 업데이트
                self.last_save_path = file_path
                messagebox.showinfo("저장 완료", f"모든 회사 데이터가 새 파일에 성공적으로 저장되었습니다:\n{file_path}")

            except Exception as e:
                messagebox.showerror("저장 오류", f"파일 저장 중 오류가 발생했습니다: {e}")

    # 3. 전체 회사를 저장 (Ctrl+S)
    def save_all_companies(self):
        """last_save_path에 따라 덮어쓰거나 '다른 이름으로 저장'을 실행합니다."""
        if not self.all_companies_data:
            messagebox.showwarning("저장 불가", "저장할 회사 데이터가 없습니다.")
            return

        if self.last_save_path and os.path.exists(self.last_save_path):
            # 저장 경로가 있고 파일이 존재하면 덮어쓰기
            try:
                formatted_data = self._format_data()
                with open(self.last_save_path, 'w', encoding='utf-8') as f:
                    f.write(formatted_data)
                messagebox.showinfo("저장 완료", f"현재 데이터가 다음 파일에 덮어쓰기 저장되었습니다:\n{self.last_save_path}")
            except Exception as e:
                messagebox.showerror("저장 오류", f"파일 덮어쓰기 중 오류가 발생했습니다: {e}")
                self.last_save_path = None  # 저장 실패 시 경로 초기화
        else:
            # 경로가 없거나 유효하지 않으면 '다른 이름으로 저장' 실행
            self.save_all_companies_as()

    # 4. SQL 파일로 내보내기
    def export_to_sql(self):
        """모든 회사 데이터를 SQLite DB 파일로 내보냅니다."""
        if not self.all_companies_data:
            messagebox.showwarning("내보내기 불가", "내보낼 데이터가 없습니다.")
            return

        self.save_current_company_data()

        file_path = filedialog.asksaveasfilename(
            defaultextension=".sqlite",
            initialfile="jaesoseo_db.sqlite",
            filetypes=[("SQLite Database", "*.sqlite"), ("All files", "*.*")],
            title="모든 회사 데이터를 SQLite 파일로 내보냅니다."
        )

        if file_path:
            try:
                conn = sqlite3.connect(file_path)
                cursor = conn.cursor()

                # 테이블 생성 (기존 테이블이 있다면 삭제 후 재생성)
                cursor.execute("DROP TABLE IF EXISTS questions")
                cursor.execute("""
                               CREATE TABLE questions
                               (
                                   id               INTEGER PRIMARY KEY AUTOINCREMENT,
                                   company_name     TEXT NOT NULL,
                                   question_title   TEXT,
                                   question_type    TEXT,
                                   question_content TEXT,
                                   answer_content   TEXT
                               )
                               """)

                # 데이터 삽입
                for company_name, questions in self.all_companies_data.items():
                    for q_data in questions:
                        cursor.execute("""
                                       INSERT INTO questions (company_name, question_title, question_type,
                                                              question_content, answer_content)
                                       VALUES (?, ?, ?, ?, ?)
                                       """, (
                                           company_name,
                                           q_data.get('제목', '제목 없음'),
                                           q_data.get('문항유형', ''),
                                           q_data.get('질문', ''),
                                           q_data.get('답변', '')
                                       ))

                conn.commit()
                conn.close()
                messagebox.showinfo("내보내기 완료", f"데이터가 SQLite 파일에 성공적으로 저장되었습니다:\n{file_path}")

            except Exception as e:
                messagebox.showerror("SQL 내보내기 오류", f"데이터베이스 저장 중 오류가 발생했습니다: {e}")

    # 1. 텍스트 파일 불러오기
    def load_text_file(self):
        """파일을 열어 데이터를 파싱하고 회사 목록에 추가/갱신합니다."""
        file_path = filedialog.askopenfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            title="불러올 자소서 텍스트 파일을 선택하세요."
        )

        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

            new_data = self._parse_file_content(content)

            if not new_data:
                messagebox.showwarning("파싱 오류", "파일에서 유효한 회사 및 문항 데이터를 찾을 수 없습니다.")
                return

            # 기존 데이터에 불러온 데이터 병합 (동일 회사명은 덮어씀)
            self.all_companies_data.update(new_data)
            self._update_treeview()

            # 불러오기 성공 시 last_save_path 설정
            self.last_save_path = file_path

            messagebox.showinfo("불러오기 완료", f"총 {len(new_data)}개의 회사 데이터를 성공적으로 불러왔습니다.")

            first_company = next(iter(new_data))
            self.company_tree.selection_set(first_company)
            self.company_tree.focus(first_company)
            self.load_company_data(None)

        except Exception as e:
            messagebox.showerror("불러오기 오류", f"파일을 읽거나 파싱하는 중 오류가 발생했습니다: {e}")

    # 2. SQL 파일로부터 추출하기
    def load_from_sql_file(self):
        """SQLite DB 파일에서 데이터를 추출하여 회사 목록에 추가/갱신합니다."""
        file_path = filedialog.askopenfilename(
            defaultextension=".sqlite",
            filetypes=[("SQLite Database", "*.sqlite"), ("All files", "*.*")],
            title="추출할 SQLite 데이터베이스 파일을 선택하세요."
        )

        if not file_path:
            return

        if not os.path.exists(file_path):
            messagebox.showerror("파일 오류", "선택한 파일이 존재하지 않습니다.")
            return

        try:
            conn = sqlite3.connect(file_path)
            cursor = conn.cursor()

            cursor.execute(
                "SELECT company_name, question_title, question_type, question_content, answer_content FROM questions")
            rows = cursor.fetchall()
            conn.close()

            if not rows:
                messagebox.showwarning("데이터 없음", "선택한 데이터베이스 파일에 유효한 'questions' 테이블 데이터가 없습니다.")
                return

            new_data = {}
            for row in rows:
                company_name, title, q_type, question, answer = row

                if company_name not in new_data:
                    new_data[company_name] = []

                new_data[company_name].append({
                    "제목": title,
                    "질문": question,
                    "문항유형": q_type,
                    "답변": answer
                })

            # 기존 데이터에 불러온 데이터 병합
            self.all_companies_data.update(new_data)
            self._update_treeview()

            messagebox.showinfo("추출 완료", f"SQLite 파일에서 총 {len(new_data)}개의 회사 데이터를 성공적으로 추출했습니다.")

            first_company = next(iter(new_data))
            self.company_tree.selection_set(first_company)
            self.company_tree.focus(first_company)
            self.load_company_data(None)

        except sqlite3.OperationalError as e:
            messagebox.showerror("DB 오류", f"데이터베이스 구조 오류: 'questions' 테이블을 찾을 수 없거나 형식이 올바르지 않습니다.\n{e}")
        except Exception as e:
            messagebox.showerror("추출 오류", f"SQLite 파일에서 데이터를 추출하는 중 오류가 발생했습니다: {e}")

    # --- 검색 로직 (변경 없음) ---
    def open_search_popup(self):
        """검색 팝업을 열고 검색 결과를 표시합니다."""

        if not self.all_companies_data:
            messagebox.showinfo("검색 불가", "검색할 회사 데이터가 없습니다. 먼저 회사와 문항을 추가해주세요.")
            return

        popup = tk.Toplevel(self)
        popup.title("모든 회사 문항 통합 검색")
        popup.transient(self)
        popup.grab_set()
        popup.geometry("600x400")

        popup_frame = ttk.Frame(popup, padding="10")
        popup_frame.pack(expand=True, fill="both")

        search_control_frame = ttk.Frame(popup_frame)
        search_control_frame.pack(fill="x", pady=(0, 10))

        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_control_frame, textvariable=search_var, width=50, font=('Arial', 10))
        search_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        search_entry.focus_set()

        results_text = tk.Text(popup_frame, wrap='word', font=('Arial', 10), state='disabled')
        results_text.pack(fill="both", expand=True)

        def perform_search(event=None):
            query = search_var.get().strip().lower()
            results_text.config(state='normal')
            results_text.delete("1.0", tk.END)

            if not query:
                results_text.insert(tk.END, "검색어를 입력해주세요.")
                results_text.config(state='disabled')
                return

            found_count = 0
            self.save_current_company_data()

            for company_name, questions in self.all_companies_data.items():
                for i, q_data in enumerate(questions):
                    question_title = q_data.get('제목', f'문항 {i + 1}')

                    fields_to_search = [
                        ("제목", question_title),
                        ("유형", q_data.get('문항유형', '')),
                        ("질문", q_data.get('질문', '')),
                        ("답변", q_data.get('답변', ''))
                    ]

                    match_in_fields = []

                    for field_name, content in fields_to_search:
                        if query in content.lower():
                            match_in_fields.append(field_name)

                    if match_in_fields:
                        found_count += 1
                        results_text.insert(tk.END,
                                            f"회사: {company_name}\n"
                                            f"   - 문항: {question_title}\n"
                                            f"   - 검색 일치: {', '.join(match_in_fields)}에서 발견\n\n",
                                            'result_tag'
                                            )

            if found_count == 0:
                results_text.insert(tk.END, f"'{query}'에 해당하는 항목을 찾을 수 없습니다.")
            else:
                results_text.insert("1.0", f"총 {found_count}개의 항목을 찾았습니다.\n\n", 'summary_tag')

            results_text.config(state='disabled')

        def on_cancel(event=None):
            popup.destroy()

        search_button = ttk.Button(search_control_frame, text="검색", command=perform_search)
        search_button.pack(side="left")

        results_text.tag_configure('summary_tag', font=('Arial', 10, 'bold'), foreground='blue')
        results_text.tag_configure('result_tag', font=('Arial', 10, 'normal'))

        search_entry.bind('<Return>', perform_search)
        popup.bind('<Escape>', on_cancel)

        ttk.Button(popup_frame, text="닫기", command=on_cancel).pack(pady=(10, 0))

        self.wait_window(popup)


# 애플리케이션 실행
if __name__ == "__main__":
    app = Application()
    app.mainloop()
