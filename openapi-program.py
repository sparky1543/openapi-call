import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import requests
import csv
import json
import xml.etree.ElementTree as ET
import urllib.parse

class OpenAPIClient:
    """
    OpenAPI 호출 및 결과 처리를 위한 클라이언트 GUI 클래스
    """
    def __init__(self, root):
        """
        OpenAPI 클라이언트 초기화
        
        Args:
            root: tkinter 루트 윈도우
        """
        self.root = root
        self.root.title("OpenAPI CALL")
        self.root.geometry("1000x800")
        
        # GUI 폰트 설정
        self.font = ("맑은 고딕", 13)
        self.small_font = ("맑은 고딕", 9)
        
        # ttk 스타일 설정
        self.style = ttk.Style()
        self.style.configure("TButton", font=self.font)
        self.style.configure("TLabel", font=self.font)
        self.style.configure("TLabelframe", font=self.font)
        self.style.configure("TLabelframe.Label", font=self.font)
        
        # 창 크기 조절 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # 메인 프레임 생성
        self.main_frame = tk.Frame(root)
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.main_frame.columnconfigure(0, weight=1)
        
        # URL 입력 섹션 구성
        self.url_frame = tk.Frame(self.main_frame)
        self.url_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        self.url_frame.columnconfigure(1, weight=1)

        ttk.Label(self.url_frame, text="요청 URL", font=self.font).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.url_entry = tk.Entry(self.url_frame, width=50, font=self.font, bd=1, highlightthickness=0)
        self.url_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        # 요청변수 설정 체크박스 및 레이블 추가
        self.param_setting_frame = tk.Frame(self.url_frame)
        self.param_setting_frame.grid(row=0, column=2, sticky="e", padx=5, pady=5)

        self.param_setting_var = tk.BooleanVar()
        self.param_setting_label = tk.Label(self.param_setting_frame, text="요청변수 설정", font=self.font)
        self.param_setting_label.pack(side=tk.LEFT, padx=(0, 2))

        self.param_setting_checkbox = tk.Checkbutton(
            self.param_setting_frame, 
            variable=self.param_setting_var,
            command=self.parse_url_parameters
        )
        self.param_setting_checkbox.pack(side=tk.LEFT)

        # URL 안내 문구 추가
        self.url_guide = tk.Label(self.url_frame, text="※ 오픈API 미리보기의 전체 URL을 넣어서 조회 가능, 일반 인증키(Encoding) 값을 넣어주세요", 
                                font=self.small_font, fg="blue", anchor="w")
        self.url_guide.grid(row=1, column=1, sticky="w", padx=5)
        
        # 파라미터 섹션 구성
        self.param_container_frame = ttk.LabelFrame(self.main_frame, text="요청변수", style="TLabelframe")
        self.param_container_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=5)
        self.param_container_frame.columnconfigure(0, weight=1)
        
        # 필수 파라미터 입력 영역
        self.fixed_param_frame = tk.Frame(self.param_container_frame)
        self.fixed_param_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        self.fixed_param_frame.columnconfigure(0, weight=1)  # serviceKey 그룹
        self.fixed_param_frame.columnconfigure(1, weight=0)  # pageNo 그룹
        self.fixed_param_frame.columnconfigure(2, weight=0)  # numOfRows 그룹
        
        # serviceKey 입력 영역
        self.service_key_frame = tk.Frame(self.fixed_param_frame)
        self.service_key_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        self.service_key_frame.columnconfigure(1, weight=1)

        # 라벨을 담을 프레임 생성
        service_key_label_frame = tk.Frame(self.service_key_frame)
        service_key_label_frame.grid(row=0, column=0, sticky="w", padx=5, pady=5)

        ttk.Label(service_key_label_frame, text="서비스키", font=self.font).pack(anchor="w")
        ttk.Label(service_key_label_frame, text="(serviceKey)", font=("맑은 고딕", 10)).pack(anchor="w")

        self.service_key_entry = tk.Entry(self.service_key_frame, font=self.font, bd=1, highlightthickness=0)
        self.service_key_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        # pageNo 입력 영역
        self.page_no_frame = tk.Frame(self.fixed_param_frame)
        self.page_no_frame.grid(row=0, column=1, sticky="e", padx=5, pady=5)

        # 라벨을 담을 프레임 생성
        page_no_label_frame = tk.Frame(self.page_no_frame)
        page_no_label_frame.grid(row=0, column=0, sticky="w", padx=5, pady=5)

        ttk.Label(page_no_label_frame, text="페이지 번호", font=self.font).pack(anchor="w")
        ttk.Label(page_no_label_frame, text="(pageNo)", font=("맑은 고딕", 10)).pack(anchor="w")

        self.page_no_entry = tk.Entry(self.page_no_frame, font=self.font, bd=1, highlightthickness=0, width=12)
        self.page_no_entry.grid(row=0, column=1, sticky="e", padx=5, pady=5)

        # numOfRows 입력 영역
        self.num_rows_frame = tk.Frame(self.fixed_param_frame)
        self.num_rows_frame.grid(row=0, column=2, sticky="e", padx=5, pady=5)

        # 라벨을 담을 프레임 생성
        num_rows_label_frame = tk.Frame(self.num_rows_frame)
        num_rows_label_frame.grid(row=0, column=0, sticky="w", padx=5, pady=5)

        ttk.Label(num_rows_label_frame, text="페이지별 행 수", font=self.font).pack(anchor="w")
        ttk.Label(num_rows_label_frame, text="(numOfRows)", font=("맑은 고딕", 10)).pack(anchor="w")

        self.num_of_rows_entry = tk.Entry(self.num_rows_frame, font=self.font, bd=1, highlightthickness=0, width=12)
        self.num_of_rows_entry.grid(row=0, column=1, sticky="e", padx=5, pady=5)
        
        # serviceKey 안내 문구
        self.service_key_guide = tk.Label(self.fixed_param_frame, text="※ 일반 인증키(Decoding) 값을 넣어주세요", 
                                         font=self.small_font, fg="blue", anchor="w")
        self.service_key_guide.grid(row=1, column=0, sticky="w", padx=10, pady=(0, 5), columnspan=3)
        
        # 스크롤 가능한 파라미터 영역 생성
        self.frame_canvas = tk.Frame(self.param_container_frame)
        self.frame_canvas.grid(row=2, column=0, sticky="ew")
        self.frame_canvas.grid_rowconfigure(0, weight=1)
        self.frame_canvas.grid_columnconfigure(0, weight=1)
        
        # 파라미터 캔버스 설정
        param_height = 160
        self.param_canvas = tk.Canvas(self.frame_canvas, height=param_height, highlightthickness=0, bd=0)
        self.param_canvas.grid(row=0, column=0, sticky="ew")
        
        # 스크롤바 설정
        self.param_scrollbar = ttk.Scrollbar(self.frame_canvas, orient="vertical", command=self.param_canvas.yview)
        self.param_scrollbar.grid(row=0, column=1, sticky="ns")
        self.param_canvas.configure(yscrollcommand=self.param_scrollbar.set)
        
        # 파라미터 프레임을 캔버스 내부에 배치
        self.param_frame = tk.Frame(self.param_canvas)
        self.param_canvas_window = self.param_canvas.create_window((0, 0), window=self.param_frame, anchor="nw")
        
        # 캔버스 크기 동적 조정 설정
        def configure_canvas(event):
            """캔버스 너비 조정"""
            canvas_width = event.width
            self.param_canvas.itemconfig(self.param_canvas_window, width=canvas_width)
            
        self.param_canvas.bind('<Configure>', configure_canvas)
        
        # 스크롤 영역 업데이트 함수
        def on_frame_configure(event):
            """스크롤 영역 설정"""
            bbox = self.param_canvas.bbox("all")
            # 높이가 캔버스보다 작으면 컨텐츠 크기로 설정
            if bbox[3] < self.param_canvas.winfo_height():  
                self.param_canvas.configure(scrollregion=(0, 0, bbox[2], self.param_canvas.winfo_height()))
            else:
                self.param_canvas.configure(scrollregion=bbox)
            
        self.param_frame.bind("<Configure>", on_frame_configure)
        
        # 마우스 휠 이벤트 처리
        def on_mousewheel(event):
            """마우스 휠 이벤트 처리"""
            # Linux 시스템 이벤트
            if hasattr(event, 'num') and event.num == 4:
                self.param_canvas.yview_scroll(-1, "units")
            elif hasattr(event, 'num') and event.num == 5:
                self.param_canvas.yview_scroll(1, "units")
            else:
                # Windows/MacOS 이벤트
                self.param_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
                
        # 마우스 휠 이벤트 바인딩
        self.param_canvas.bind("<MouseWheel>", on_mousewheel)  # Windows/MacOS
        self.param_canvas.bind("<Button-4>", on_mousewheel)  # Linux 위로 스크롤
        self.param_canvas.bind("<Button-5>", on_mousewheel)  # Linux 아래로 스크롤
        
        # 행 추가 프레임
        self.add_frame = tk.Frame(self.param_container_frame)
        self.add_frame.grid(row=3, column=0, sticky="e", padx=10, pady=5)
        
        # 행 추가 라벨 버튼
        self.add_row_label = tk.Label(
            self.add_frame, 
            text="➕ 행 추가", 
            font=("맑은 고딕", 10), 
            fg="blue", 
            cursor="hand2"
        )
        self.add_row_label.pack(side=tk.RIGHT)
        
        # 행 추가 버튼 이벤트 바인딩
        self.add_row_label.bind("<Button-1>", lambda e: self.add_param_row())
        self.add_row_label.bind("<Enter>", lambda e: self.on_hover_enter(self.add_row_label, "blue"))
        self.add_row_label.bind("<Leave>", lambda e: self.on_hover_leave(self.add_row_label, "blue"))
        
        # 파라미터 항목 저장 리스트
        self.param_entries = []
        
        # 결과 프레임과 저장 버튼
        self.results_outer_frame = ttk.LabelFrame(self.main_frame, text="결과", style="TLabelframe")
        self.results_outer_frame.grid(row=4, column=0, sticky="nsew", padx=10, pady=10)
        self.results_outer_frame.columnconfigure(0, weight=1)
        self.results_outer_frame.rowconfigure(1, weight=1)
        
        # 결과 영역 확장 설정
        self.main_frame.rowconfigure(4, weight=1)
        
        # 결과 상단 프레임
        self.results_top_frame = tk.Frame(self.results_outer_frame)
        self.results_top_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        
        # 상태 표시 프레임
        self.status_frame = tk.Frame(self.results_top_frame)
        self.status_frame.pack(side=tk.LEFT, padx=5, fill="x", expand=True)
        
        # 상태 태그 라벨
        self.status_tag = tk.Label(self.status_frame, text="준비", font=(self.font[0], self.font[1], "bold"), fg="#646464")
        self.status_tag.pack(side=tk.LEFT)
        
        # 상태 메시지 라벨
        self.status_detail = tk.Label(self.status_frame, text="URL을 입력하세요", font=self.font)
        self.status_detail.pack(side=tk.LEFT, padx=(5, 0))
        
        # 버튼 배치
        self.save_button = ttk.Button(self.results_top_frame, text="저장하기", command=self.save_to_csv)
        self.save_button.pack(side=tk.RIGHT, padx=5)
        
        self.request_button = ttk.Button(self.results_top_frame, text="조회하기", command=self.send_request)
        self.request_button.pack(side=tk.RIGHT, padx=5)
        
        # 결과 표시 프레임
        self.results_frame = tk.Frame(self.results_outer_frame)
        self.results_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.results_frame.columnconfigure(0, weight=1)
        self.results_frame.rowconfigure(0, weight=1)
        
        # 결과 트리뷰 생성
        self.create_treeview()
        
        # 저장용 데이터 저장소
        self.current_data = {"headers": [], "rows": []}
        
        # 초기 파라미터 행 추가
        for i in range(3):
            self.add_param_row()
            
        # 초기 상태 설정
        self.set_status("준비", "URL을 입력하세요")
        
        # 고정 파라미터 영역에 마우스 휠 이벤트 바인딩
        for widget in [self.service_key_frame, self.page_no_frame, self.num_rows_frame]:
            widget.bind("<MouseWheel>", on_mousewheel)
            widget.bind("<Button-4>", on_mousewheel)
            widget.bind("<Button-5>", on_mousewheel)

    def set_status(self, status_type, message):
        """
        상태 메시지 설정 및 표시
        
        Args:
            status_type: 상태 유형 (준비, 완료, 오류, 경고, 진행 중)
            message: 표시할 메시지
        """
        if status_type == "준비":
            self.status_tag.config(text=f"{status_type}", fg="#646464")
        elif status_type == "완료":
            self.status_tag.config(text=f"{status_type}", fg="green")
        elif status_type == "오류" or status_type == "경고":
            self.status_tag.config(text=f"{status_type}", fg="red")
            # 오류일 경우 상세 메시지는 표시하지 않고 메시지 박스로 표시
            if status_type == "오류":
                self.status_detail.config(text="")
                messagebox.showerror("오류", message)
                return
        elif status_type == "진행 중":
            self.status_tag.config(text=f"{status_type}", fg="blue")
        else:
            self.status_tag.config(text=f"{status_type}", fg="black")
            
        self.status_detail.config(text=message)

    def create_treeview(self):
        """
        결과 데이터 표시를 위한 트리뷰 생성
        """
        # 스크롤바와 함께 트리뷰 프레임 생성
        self.tree_frame = tk.Frame(self.results_frame)
        self.tree_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        self.tree_frame.columnconfigure(0, weight=1)
        self.tree_frame.rowconfigure(0, weight=1)
        
        # 스크롤바 생성
        self.vsb = ttk.Scrollbar(self.tree_frame, orient="vertical")
        self.hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal")
        
        # 트리뷰 스타일 설정
        self.style.configure("Treeview", font=self.font)
        self.style.configure("Treeview.Heading", font=self.font)
        
        # 트리뷰 생성 및 스크롤바 연결
        self.tree = ttk.Treeview(self.tree_frame, yscrollcommand=self.vsb.set, 
                                xscrollcommand=self.hsb.set, style="Treeview")
        
        # 스크롤바 설정
        self.vsb.config(command=self.tree.yview)
        self.hsb.config(command=self.tree.xview)
        
        # 그리드 레이아웃 배치
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.hsb.grid(row=1, column=0, sticky="ew")
        self.tree.grid(row=0, column=0, sticky="nsew")
        
        # 빈 컬럼으로 초기화
        self.tree["columns"] = ()
        self.tree["show"] = "headings"

    def encode_service_key(self, service_key):
        """
        서비스키를 URL 인코딩하는 함수
        
        Args:
            service_key: 인코딩할 서비스키
            
        Returns:
            str: 인코딩된 서비스키
        """
        if not service_key:
            return ""
        
        # URL 인코딩 수행
        encoded_key = urllib.parse.quote_plus(service_key)
        return encoded_key
        
    def decode_service_key(self, service_key):
        """
        서비스키를 URL 디코딩하는 함수
        
        Args:
            service_key: 디코딩할 서비스키
            
        Returns:
            str: 디코딩된 서비스키
        """
        if not service_key:
            return ""
        
        try:
            # URL 디코딩 수행
            decoded_key = urllib.parse.unquote_plus(service_key)
            return decoded_key
        except:
            # 디코딩 실패 시 원본 반환
            return service_key
        
    def parse_url_parameters(self):
        """
        URL에서 파라미터를 파싱하여 UI 요소에 채우는 함수
        체크박스 상태에 따라 URL에서 파라미터 부분을 분리하거나 합치는 작업 수행
        """
        if self.param_setting_var.get():  # 체크박스가 체크되었을 때
            url = self.url_entry.get().strip()
            
            # URL에 파라미터 부분이 없는 경우 (? 문자가 없는 경우)
            if "?" not in url:
                # 체크박스를 다시 해제하고 메시지 표시
                self.param_setting_var.set(False)
                messagebox.showinfo("안내", "설정할 요청변수가 없습니다.")
                return
                
            # 파라미터가 있는 경우 기존 로직 계속 진행
            base_url, params_str = url.split("?", 1)
            
            # URL 입력창 업데이트
            self.url_entry.delete(0, tk.END)
            self.url_entry.insert(0, base_url)
            
            # 파라미터 파싱
            if params_str:
                param_pairs = params_str.split("&")
                
                # 기존 파라미터 행 초기화
                self.clear_param_entries()
                
                # 파라미터 딕셔너리 생성
                params = {}
                for pair in param_pairs:
                    if "=" in pair:
                        name, value = pair.split("=", 1)
                        params[name] = value
                
                # 고정 파라미터 처리
                if "serviceKey" in params:
                    # 서비스키는 디코딩하여 저장
                    encoded_service_key = params.pop("serviceKey")
                    decoded_service_key = self.decode_service_key(encoded_service_key)
                    
                    self.service_key_entry.delete(0, tk.END)
                    self.service_key_entry.insert(0, decoded_service_key)
                
                if "pageNo" in params:
                    self.page_no_entry.delete(0, tk.END)
                    self.page_no_entry.insert(0, params.pop("pageNo"))
                
                if "numOfRows" in params:
                    self.num_of_rows_entry.delete(0, tk.END)
                    self.num_of_rows_entry.insert(0, params.pop("numOfRows"))
                
                # 나머지 파라미터 처리
                for name, value in params.items():
                    name_entry, value_entry = self.add_param_row()
                    name_entry.delete(0, tk.END)
                    name_entry.insert(0, name)
                    value_entry.delete(0, tk.END)
                    value_entry.insert(0, value)
        else:  # 체크박스가 해제되었을 때
            # URL 재조합
            self.combine_url_with_params()

    def clear_param_entries(self):
        """
        현재 파라미터 입력 행 초기화
        """
        # 기존 파라미터 행 모두 제거
        for _, _, frame_row in self.param_entries:
            frame_row.destroy()
        
        # 파라미터 저장 리스트 초기화
        self.param_entries = []
        
        # 스크롤 영역 업데이트
        self.update_scrollregion()

    def combine_url_with_params(self):
        """
        URL과 파라미터를 조합하여 URL 입력창 업데이트
        체크박스 해제 시 요청변수 그룹의 값들을 모두 지움
        """
        base_url = self.url_entry.get().strip()
        
        # 파라미터 수집
        params = []
        
        # 고정 파라미터 처리
        service_key = self.service_key_entry.get().strip()
        if service_key:
            # 서비스키를 인코딩하여 URL에 추가
            encoded_service_key = self.encode_service_key(service_key)
            params.append(f"serviceKey={encoded_service_key}")
        
        page_no = self.page_no_entry.get().strip()
        if page_no:
            params.append(f"pageNo={page_no}")
        
        num_of_rows = self.num_of_rows_entry.get().strip()
        if num_of_rows:
            params.append(f"numOfRows={num_of_rows}")
        
        # 일반 파라미터 처리
        for name_entry, value_entry, _ in self.param_entries:
            name = name_entry.get().strip()
            value = value_entry.get().strip()
            if name:  # 이름이 비어있지 않은 경우만 추가
                params.append(f"{name}={value}")
        
        # URL 재조합
        if params:
            full_url = f"{base_url}?{'&'.join(params)}"
        else:
            full_url = base_url
        
        # URL 입력창 업데이트
        self.url_entry.delete(0, tk.END)
        self.url_entry.insert(0, full_url)
        
        # 체크박스가 해제되었을 때 요청변수 그룹의 값 모두 지우기
        self.service_key_entry.delete(0, tk.END)
        self.page_no_entry.delete(0, tk.END)
        self.num_of_rows_entry.delete(0, tk.END)
        
        # 기존 일반 파라미터 행 초기화
        self.clear_param_entries()
        
        # 기본 파라미터 행 3개 추가
        for i in range(3):
            self.add_param_row()
    
    def add_param_row(self):
        """
        파라미터 입력 행 추가
        
        Returns:
            tuple: 생성된 이름 입력 필드와 값 입력 필드
        """
        row_idx = len(self.param_entries)
        
        # 행 컨테이너 프레임 생성
        frame_row = tk.Frame(self.param_frame)
        frame_row.pack(fill="x", pady=2)
        
        # 파라미터 이름 입력 필드
        ttk.Label(frame_row, text=f"항목명", font=self.font).pack(side="left", padx=5, pady=5)
        name_entry = tk.Entry(frame_row, font=self.font, bd=1, highlightthickness=0, width=20)
        name_entry.pack(side="left", padx=5, pady=5, fill="x", expand=True)
        
        # 파라미터 값 입력 필드
        ttk.Label(frame_row, text=f"값", font=self.font).pack(side="left", padx=10, pady=5)
        value_entry = tk.Entry(frame_row, font=self.font, bd=1, highlightthickness=0, width=20)
        value_entry.pack(side="left", padx=5, pady=5, fill="x", expand=True)
        
        # 행 삭제 버튼
        delete_label = tk.Label(
            frame_row, 
            text="✕", 
            font=("맑은 고딕", 9), 
            fg="red", 
            cursor="hand2"
        )
        delete_label.pack(side="left", padx=5, pady=5)
        
        # 삭제 이벤트 바인딩
        delete_label.bind("<Button-1>", lambda e, fr=frame_row: self.delete_param_row(fr))
        delete_label.bind("<Enter>", lambda e, label=delete_label: self.on_hover_enter(label, "red"))
        delete_label.bind("<Leave>", lambda e, label=delete_label: self.on_hover_leave(label, "red"))
        
        # 마우스 휠 이벤트 바인딩
        for widget in [frame_row, name_entry, value_entry, delete_label]:
            widget.bind("<MouseWheel>", lambda e, w=widget: self.on_mousewheel(e))
            widget.bind("<Button-4>", lambda e, w=widget: self.on_mousewheel(e))
            widget.bind("<Button-5>", lambda e, w=widget: self.on_mousewheel(e))
        
        self.param_entries.append((name_entry, value_entry, frame_row))
        
        # 스크롤 영역 업데이트
        self.update_scrollregion()
        
        return name_entry, value_entry

    def delete_param_row(self, frame_row):
        """
        파라미터 행 삭제
        
        Args:
            frame_row: 삭제할 파라미터 행 프레임
        """
        # 삭제 확인 메시지 표시
        confirm = messagebox.askyesno("삭제 확인", "이 행을 삭제하시겠습니까?")
        
        # '예'를 선택한 경우에만 삭제 진행
        if confirm:
            # 목록에서 항목 제거
            for i, (name_entry, value_entry, fr) in enumerate(self.param_entries):
                if fr == frame_row:
                    self.param_entries.pop(i)
                    break
            
            # 위젯 제거
            frame_row.destroy()
                    
            # 스크롤 영역 업데이트
            self.update_scrollregion()
    
    def on_mousewheel(self, event):
        """
        마우스 휠 스크롤 이벤트 처리
        
        Args:
            event: 마우스 휠 이벤트 객체
        """
        # Linux 시스템 이벤트
        if hasattr(event, 'num') and event.num == 4:
            self.param_canvas.yview_scroll(-1, "units")
        elif hasattr(event, 'num') and event.num == 5:
            self.param_canvas.yview_scroll(1, "units")
        else:
            # Windows/MacOS 이벤트
            delta = event.delta
            self.param_canvas.yview_scroll(int(-1*(delta/120)), "units")

    def on_hover_enter(self, label, color="blue"):
        """
        마우스 호버 진입 이벤트 처리
        
        Args:
            label: 레이블 위젯
            color: 텍스트 색상
        """
        if color == "blue":
            label.config(fg="darkblue")
        else:
            label.config(fg="darkred")

    def on_hover_leave(self, label, color="blue"):
        """
        마우스 호버 이탈 이벤트 처리
        
        Args:
            label: 레이블 위젯
            color: 텍스트 색상
        """
        if color == "blue":
            label.config(fg="blue")
        else:
            label.config(fg="red")
    
    def update_scrollregion(self):
        """
        파라미터 영역 스크롤 영역 업데이트
        """
        self.param_frame.update_idletasks()
        bbox = self.param_canvas.bbox("all")
        # 높이가 캔버스보다 작으면 컨텐츠 크기로 설정
        if bbox[3] < self.param_canvas.winfo_height():  
            self.param_canvas.configure(scrollregion=(0, 0, bbox[2], self.param_canvas.winfo_height()))
        else:
            self.param_canvas.configure(scrollregion=bbox)

    def send_request(self):
        """
        API 요청 전송 및 결과 처리
        """
        # 이전 데이터 지우기
        self.clear_treeview()
        
        url = self.url_entry.get().strip()
        if not url:
            self.set_status("오류", "URL이 필요합니다")
            return
        
        try:
            self.set_status("진행 중", "요청 전송 중...")
            self.root.update()
            
            # URL 처리 로직 수정
            if '?' in url:
                # URL에 파라미터가 이미 포함된 경우, URL 그대로 사용
                response = requests.get(url, verify=True)
            else:
                # URL에 파라미터가 없는 경우, 체크박스 상태와 관계없이 항상 요청변수 합쳐서 요청
                params = {}
                
                # 고정 파라미터 처리
                service_key = self.service_key_entry.get().strip()
                page_no = self.page_no_entry.get().strip()
                num_of_rows = self.num_of_rows_entry.get().strip()
                
                if service_key:
                    # 요청변수로 보낼 때는 디코딩된 형태 그대로 사용
                    params["serviceKey"] = service_key
                if page_no:
                    params["pageNo"] = page_no
                if num_of_rows:
                    params["numOfRows"] = num_of_rows
                
                # 일반 파라미터 처리
                for name_entry, value_entry, _ in self.param_entries:
                    name = name_entry.get().strip()
                    value = value_entry.get().strip()
                    if name:  # 이름이 비어있지 않은 경우만 추가
                        params[name] = value
                
                response = requests.get(url, params=params, verify=True)
            
            response.raise_for_status()
            
            # 응답 타입에 따른 처리
            content_type = response.headers.get('Content-Type', '').lower()
            
            if 'application/json' in content_type:
                self.process_json_response(response.text)
            elif 'application/xml' in content_type or 'text/xml' in content_type:
                self.process_xml_response(response.text)
            else:
                # 타입 명시가 없을 경우 JSON 먼저 시도, 실패 시 XML 시도
                try:
                    self.process_json_response(response.text)
                except:
                    try:
                        self.process_xml_response(response.text)
                    except:
                        self.set_status("오류", "지원되지 않는 응답 형식")
            
        except requests.exceptions.RequestException as e:
            self.set_status("오류", f"{str(e)}")
        except Exception as e:
            self.set_status("오류", f"{str(e)}")

    def process_json_response(self, json_text):
        """
        JSON 응답 처리
        
        Args:
            json_text: JSON 형식의 텍스트 응답
        """
        try:
            json_data = json.loads(json_text)  # JSON 문자열을 객체로 파싱
            
            # cmmMsgHeader 오류 확인
            header = json_data.get("cmmMsgHeader", {})
            if header:
                err_msg = header.get("errMsg", "")
                if "SERVICE ERROR" in err_msg:
                    # 오류 상세 정보 가져오기
                    return_auth_msg = header.get("returnAuthMsg", "")
                    return_reason_code = header.get("returnReasonCode", "")
                    # 오류 메시지 구성
                    error_detail = f"인증 메시지: {return_auth_msg}\n코드: {return_reason_code}"
                    # 상태 설정 및 메시지 박스 표시
                    self.set_status("오류", error_detail)
                    return
            
            # 데이터 항목 찾기
            data_rows = []
            
            # 일반적인 API 응답 구조 분석
            try:
                # 일반적인 API 응답 구조 - response > body > items > item
                items = json_data.get("response", {}).get("body", {}).get("items", {}).get("item", [])
                if items:
                    if isinstance(items, list):
                        data_rows = items
                    elif isinstance(items, dict):  # 단일 항목인 경우 리스트로 변환
                        data_rows = [items]
            except:
                pass
                
            # 다른 일반적인 API 응답 패턴 처리
            if not data_rows:
                # 최상위 배열 확인
                if isinstance(json_data, list) and len(json_data) > 0:
                    data_rows = json_data
                # 요소 중 배열 찾기
                elif isinstance(json_data, dict):
                    for key, value in json_data.items():
                        if isinstance(value, list) and len(value) > 0:
                            if isinstance(value[0], dict):
                                data_rows = value
                                break
                        # 중첩된 구조에서 배열 찾기
                        elif isinstance(value, dict):
                            for subkey, subvalue in value.items():
                                if isinstance(subvalue, list) and len(subvalue) > 0:
                                    if isinstance(subvalue[0], dict):
                                        data_rows = subvalue
                                        break
                            if data_rows:
                                break
            
            # 적절한 데이터를 찾지 못한 경우 전체 JSON 구조를 단일 행으로 처리
            if not data_rows:
                if isinstance(json_data, dict):
                    data_rows = [json_data]
                else:
                    self.set_status("경고", "JSON 응답에서 데이터 구조를 식별할 수 없습니다")
                    return
            
            # 모든 행에서 가능한 모든 헤더 추출
            all_headers = set()
            for row in data_rows:
                if isinstance(row, dict):
                    all_headers.update(row.keys())
            
            headers = sorted(list(all_headers))
            
            # 나중에 저장하기 위해 데이터 저장
            self.current_data["headers"] = headers
            self.current_data["rows"] = data_rows
            
            # 트리뷰 업데이트
            self.update_treeview(headers, data_rows)
            
        except json.JSONDecodeError as e:
            self.set_status("오류", f"JSON 파싱 오류: {str(e)}")

    def process_xml_response(self, xml_text):
        """
        XML 응답 처리
        
        Args:
            xml_text: XML 형식의 텍스트 응답
        """
        try:
            root = ET.fromstring(xml_text)
            
            # cmmMsgHeader 오류 확인
            header = root.find(".//cmmMsgHeader")
            if header is not None:
                err_msg_elem = header.find("errMsg")
                if err_msg_elem is not None and "SERVICE ERROR" in (err_msg_elem.text or ""):
                    # 오류 상세 정보 가져오기
                    return_auth_msg = ""
                    return_reason_code = ""
                    
                    auth_msg_elem = header.find("returnAuthMsg")
                    if auth_msg_elem is not None:
                        return_auth_msg = auth_msg_elem.text or ""
                    
                    reason_code_elem = header.find("returnReasonCode")
                    if reason_code_elem is not None:
                        return_reason_code = reason_code_elem.text or ""
                    
                    # 오류 메시지 구성
                    error_detail = f"오류 메세지: {return_auth_msg}\n오류 코드: {return_reason_code}"
                    
                    # 상태 설정 및 메시지 박스 표시
                    self.set_status("오류", error_detail)
                    return
            
            # 데이터 항목 찾기
            data_rows = []
            headers = []
            
            # 일반적인 API 응답 구조 분석 (items/item 형태)
            items = root.find(".//items")
            if items is not None:
                item_elements = items.findall("item")
                if item_elements:
                    for item in item_elements:
                        row = {child.tag: child.text for child in item}
                        data_rows.append(row)
                    
                    if data_rows:
                        headers = list(data_rows[0].keys())
            
            # 일반적인 반복 구조 처리
            if not data_rows:
                # 여러 동일한 태그가 있는지 확인
                tag_counts = {}
                for child in root:
                    tag = child.tag
                    tag_counts[tag] = tag_counts.get(tag, 0) + 1
                
                repeated_tags = [tag for tag, count in tag_counts.items() if count > 1]
                
                if repeated_tags:
                    # 첫 번째 반복되는 태그를 사용
                    first_repeated = repeated_tags[0]
                    for element in root.findall(f".//{first_repeated}"):
                        row = {child.tag: child.text for child in element}
                        data_rows.append(row)
                    
                    if data_rows:
                        headers = list(data_rows[0].keys())
            
            # 아직도 데이터가 없는 경우 최상위 요소의 자식을 단일 행으로 처리
            if not data_rows:
                # 개별 요소를 하나의 행으로 처리
                row = {}
                for child in root:
                    if len(list(child)) == 0:  # 자식이 없는 요소만 처리
                        row[child.tag] = child.text
                    else:
                        # 중첩된 요소의 경우, 값만 추출하여 '부모.자식' 형태의 키 생성
                        for subchild in child:
                            if len(list(subchild)) == 0:  # 자식이 없는 요소만 처리
                                row[f"{child.tag}.{subchild.tag}"] = subchild.text
                
                if row:
                    data_rows = [row]
                    headers = list(row.keys())
            
            # 결과 처리
            if headers and data_rows:
                # 나중에 저장하기 위해 데이터 저장
                self.current_data["headers"] = headers
                self.current_data["rows"] = data_rows
                
                # 트리뷰 업데이트
                self.update_treeview(headers, data_rows)
            else:
                self.set_status("경고", "XML 응답에서 데이터 항목을 찾을 수 없습니다")
                
        except ET.ParseError as e:
            self.set_status("오류", f"XML 파싱 오류: {str(e)}")

    def clear_treeview(self):
        """
        트리뷰 데이터 초기화
        """
        # 모든 항목과 컬럼 제거
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = ()

    def calculate_column_widths(self, headers, data_rows):
        """
        데이터 값의 최대 길이를 기준으로 컬럼 너비 계산
        
        Args:
            headers: 컬럼 헤더 목록
            data_rows: 데이터 행 목록
            
        Returns:
            dict: 각 컬럼의 너비 정보
        """
        column_widths = {}
        font = self.font
        
        # tkinter에서 텍스트 길이를 픽셀로 측정하기 위한 임시 레이블
        temp_label = tk.Label(self.tree_frame, font=font)
        
        # 각 헤더에 대한 최소 너비 초기화
        for header in headers:
            column_widths[header] = 100
        
        # 각 행의 데이터에 기반하여 너비 조정
        for row in data_rows:
            for header in headers:
                value = row.get(header, "")
                if value is not None:
                    value_str = str(value)
                    temp_label.config(text=value_str)
                    width = temp_label.winfo_reqwidth() + 20
                    # 현재 저장된 최대 너비와 비교해서 더 큰 값으로 업데이트
                    column_widths[header] = max(column_widths[header], width)
        
        # 모든 컬럼 너비를 최대 250으로 제한
        for header in column_widths:
            column_widths[header] = min(column_widths[header], 250)
        
        # 임시 레이블 제거
        temp_label.destroy()
        
        return column_widths

    def update_treeview(self, headers, data_rows):
        """
        트리뷰 데이터 업데이트
        
        Args:
            headers: 컬럼 헤더 목록
            data_rows: 데이터 행 목록
        """
        # 이전 데이터 지우기
        self.clear_treeview()
        
        if not headers or not data_rows:
            self.set_status("경고", "응답이 비어있거나 예상된 형식이 아닙니다.")
            return
        
        # 컬럼 구성
        self.tree["columns"] = headers
        
        # 컬럼 너비 계산
        column_widths = self.calculate_column_widths(headers, data_rows)
        
        # 컬럼 헤더 설정
        for header in headers:
            self.tree.heading(header, text=header)
            # 계산된 데이터 기반 너비 적용
            self.tree.column(header, width=column_widths[header], stretch=False)
        
        # 데이터 추가
        for i, row in enumerate(data_rows):
            values = [row.get(header, "") for header in headers]
            self.tree.insert("", "end", text=f"항목 {i+1}", values=values)
        
        self.set_status("완료", f"{len(data_rows)}개의 레코드 표시 중")

    def save_to_csv(self):
        """
        결과 데이터를 CSV 파일로 저장
        """
        # 저장할 데이터가 있는지 확인
        if not self.current_data["headers"] or not self.current_data["rows"]:
            messagebox.showinfo("안내", "저장할 데이터가 없습니다")
            return
        
        # 파일 위치 선택
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV 파일", "*.csv"), ("모든 파일", "*.*")]
        )
        
        if not file_path:  # 사용자가 취소함
            return
        
        try:
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=self.current_data["headers"])
                writer.writeheader()
                for row in self.current_data["rows"]:
                    # 모든 행에 모든 헤더가 있는지 확인
                    row_data = {header: row.get(header, "") for header in self.current_data["headers"]}
                    writer.writerow(row_data)
            
            self.set_status("완료", "저장 완료")
            messagebox.showinfo("성공", f"데이터가 {file_path}에 저장되었습니다")
        except Exception as e:
            self.set_status("오류", f"파일 저장 오류: {str(e)}")

def main():
    root = tk.Tk()
    app = OpenAPIClient(root)
    root.mainloop()

if __name__ == "__main__":
    main()
