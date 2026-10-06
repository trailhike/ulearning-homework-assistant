import tkinter as tk
from tkinter import ttk, messagebox
import requests
from datetime import datetime
import re
import html
import os
import json
import urllib.parse
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
except ImportError:
    Workbook = None
class HomeworkChecker:
    """作业扫描核心逻辑类"""
    PAGE_SIZE = 20
    def __init__(self, token: str, user_id: str):
        self.token = token
        self.user_id = user_id
        self.base_api = "https://lms.dgut.edu.cn/courseapi"
        self.headers = {
            "AUTHORIZATION": token,
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36/537.36"
        }
        self.state_map = {
            0: "未提交",
            1: "未互评",
            5: "需重写"
        }
        self.bypass_keyword = "读书心得"
        self.session = requests.Session()
        self.session.headers.update(self.headers)

    @staticmethod
    def fix_resource_url(url: str) -> str:
        """补全相对资源链接前缀"""
        if not url.startswith("https"):
            if url.startswith("../uobs"):
                url =f'https://lms.dgut.edu.cn{url}'
                url = re.sub(r'\.\.', '', url)
            elif url.startswith("/res"):
                url =f'https://lms.dgut.edu.cn/uobs/view{url}'
        return url

    def _request(self, url: str) -> dict:
        resp = self.session.get(url, timeout=15)
        resp.raise_for_status()
        return resp.json()
    def get_course_list(self) -> list:
        url = f"{self.base_api}/courses/students?keyword=&publishStatus=1&type=1&pn=1&ps={self.PAGE_SIZE}&lang=zh"
        data = self._request(url)
        return data.get("courseList", [])
    def get_course_homework(self, course_id: str, course_name: str) -> list:
        url = f"{self.base_api}/homeworks/student/v2?ocId={course_id}&pn=1&ps=99&lang=zh"
        try:
            data = self._request(url)
            return data.get("homeworkList", [])
        except Exception:
            print(f"获取课程「{course_name}」作业失败，已跳过")
            return []
    def _download_file(self, url: str, save_path: str):
        """通用下载函数，返回(是否成功,错误信息)"""
        url = self.fix_resource_url(url)
        try:
            resp = self.session.get(url, timeout=20, stream=True)
            resp.raise_for_status()
            with open(save_path, "wb") as f:
                for chunk in resp.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            return True, ""
        except Exception as e:
            return False, str(e)
    def get_homework_detail_raw(self, homework_id, course_id, hw_type: int):
        """
        获取作业原始数据，区分个人作业(type=9)、小组作业(type=8)
        返回元组(ok, data_or_errmsg)
        个人作业返回activityHomework字典
        小组作业返回publishInfo字典
        """
        if hw_type == 9:
            url = f"https://lms.dgut.edu.cn/homeworkapi/stuHomework/homeworkDetail/{homework_id}/{self.user_id}/{course_id}"
            try:
                resp = self.session.get(url, timeout=12)
                resp.raise_for_status()
                data = resp.json()
                if data.get("code") != 1:
                    return False, f"接口返回异常，code={data.get('code')}, msg={data.get('msg', '')}"
                result = data.get("result")
                if not result:
                    return False, "解析失败：返回result字段为空"
                activity_homework = result.get("activityHomework")
                if not activity_homework:
                    return False, "解析失败：result内无activityHomework"
                return True, activity_homework
            except requests.exceptions.HTTPError as e:
                return False, f"HTTP请求失败: {str(e)}"
            except Exception as e:
                return False, f"获取详情异常: {str(e)}"
        elif hw_type == 8:
            url = f"https://lms.dgut.edu.cn/homeworkapi/groupHomework/stu/info?homeworkId={homework_id}&ocId={course_id}&roleId=9"
            try:
                resp = self.session.get(url, timeout=12)
                resp.raise_for_status()
                data = resp.json()
                if data.get("code") != 1:
                    return False, f"小组作业接口异常，code={data.get('code')}, msg={data.get('msg', '')}"
                result = data.get("result")
                if not result:
                    return False, "小组作业解析失败：result为空"
                my_group_hw = result.get("publishInfo")
                if not my_group_hw:
                    return False, "解析失败：result内无publishInfo"
                return True, my_group_hw
            except requests.exceptions.HTTPError as e:
                return False, f"小组作业HTTP请求失败:{str(e)}"
            except Exception as e:
                return False, f"小组作业获取详情异常:{str(e)}"
        else:
            return False, f"未知作业类型type={hw_type}"
    @staticmethod
    def parse_preview_content(raw_homework: dict, course_name, hw_type: int):
        tasks = []
        resource_lines = []
        safe_title = ""
        detail_unescaped = ""
        if hw_type == 9:
            homeworkRequest_raw = raw_homework.get("homeworkRequest", "") or ""
            homeworkTitle = raw_homework.get("homeworkTitle", "unknown_homework") or "unknown_homework"
            safe_title = re.sub(r'[\\/*?:"<>|]', '_', homeworkTitle)
            save_dir = str(course_name)
            detail_unescaped = html.unescape(homeworkRequest_raw)
            img_urls = re.findall(r'<img[^>]+src=["\'](.*?)["\']', detail_unescaped)
            for idx, img_url in enumerate(img_urls, start=1):
                # 修复图片url
                img_url = HomeworkChecker.fix_resource_url(img_url)
                pure_url = img_url.split("?")[0]
                ext = os.path.splitext(pure_url)[-1]
                if not ext:
                    ext = ".jpg"
                filename = f"{safe_title}_{idx}{ext}"
                save_path = os.path.join(save_dir, filename)
                tasks.append({"type": "image", "url": img_url, "save_path": save_path, "display_name": filename})
                resource_lines.append(f"[图片] {filename}")
            fileUpload_str = raw_homework.get("fileUpload", "[]") or "[]"
            try:
                file_upload_list = json.loads(fileUpload_str)
            except json.JSONDecodeError:
                file_upload_list = []
            for item in file_upload_list:
                fu_file_url = item.get("filePath")
                fu_file_name = item.get("fileName")
                if fu_file_url and fu_file_name:
                    fu_file_url = HomeworkChecker.fix_resource_url(fu_file_url)
                    safe_filename = re.sub(r'[\\/*?:"<>|]', '_', fu_file_name)
                    save_path = os.path.join(save_dir, safe_filename)
                    tasks.append({"type": "attachment", "url": fu_file_url, "save_path": save_path, "display_name": safe_filename})
                    resource_lines.append(f"[附件] {safe_filename}")
        elif hw_type == 8:
            content_raw = raw_homework.get("content", "") or ""
            homeworkTitle = raw_homework.get("title", "小组作业_unknown") or "小组作业_unknown"
            safe_title = re.sub(r'[\\/*?:"<>|]', '_', homeworkTitle)
            save_dir = str(course_name)
            detail_unescaped = html.unescape(content_raw)
            file_list = raw_homework.get("fileList", []) or []
            for idx, item in enumerate(file_list, start=1):
                fu_file_url = item.get("filePath")
                fu_file_name = item.get("fileName")
                if fu_file_url and fu_file_name:
                    fu_file_url = HomeworkChecker.fix_resource_url(fu_file_url)
                    safe_filename = re.sub(r'[\\/*?:"<>|]', '_', fu_file_name)
                    save_path = os.path.join(save_dir, safe_filename)
                    tasks.append({"type": "attachment", "url": fu_file_url, "save_path": save_path, "display_name": safe_filename})
                    resource_lines.append(f"[附件] {safe_filename}")
            img_urls = re.findall(r'<img[^>]+src=["\'](.*?)["\']', detail_unescaped)
            for idx, img_url in enumerate(img_urls, start=1):
                img_url = HomeworkChecker.fix_resource_url(img_url)
                pure_url = img_url.split("?")[0]
                ext = os.path.splitext(pure_url)[-1]
                if not ext:
                    ext = ".jpg"
                filename = f"{safe_title}_{idx}{ext}"
                save_path = os.path.join(save_dir, filename)
                tasks.append({"type": "image", "url": img_url, "save_path": save_path, "display_name": filename})
                resource_lines.append(f"[图片] {filename}")
        clean_text = re.sub(r'<[^>]+>', '', detail_unescaped).strip()
        if not clean_text:
            if len(resource_lines) > 0:
                clean_text = "【该作业无文本描述，仅包含附件资源】"
            else:
                clean_text = "解析成功，但作业内容为空"
        preview_text = clean_text
        if resource_lines:
            preview_text += f"\n\n=====待下载资源清单（保存目录:{save_dir}）=====\n"
            preview_text += "\n".join(resource_lines)
        else:
            preview_text += "\n\n=====待下载资源清单=====\n无图片与附件资源"
        return preview_text, save_dir, tasks
    @staticmethod
    def _format_time(timestamp) -> str:
        try:
            if isinstance(timestamp, (int, float)):
                dt = datetime.fromtimestamp(timestamp / 1000)
                return dt.strftime("%Y-%m-%d %H:%M:%S")
            return str(timestamp)
        except Exception:
            return "时间解析失败"
    def get_undone_homework(self) -> list:
        undone_homework = []
        course_list = self.get_course_list()
        for course in course_list:
            course_name = course.get("name", "未知课程")
            course_id = course.get("id", "")
            if self.bypass_keyword in course_name:
                continue
            homework_list = self.get_course_homework(course_id, course_name)
            for hw in homework_list:
                state = hw.get("state")
                if state in self.state_map:
                    hw_name = hw.get("title") or hw.get("homeworkTitle") or "未命名作业"
                    hw_id = hw.get("id")
                    hw_type = hw.get("type", 9)
                    raw_end_time = hw.get("endTime", 0)
                    undone_homework.append({
                        "课程名称": course_name,
                        "作业名称": hw_name,
                        "作业状态": self.state_map[state],
                        "截止时间": self._format_time(raw_end_time),
                        "截止时间戳": raw_end_time,
                        "课程ID": course_id,
                        "作业ID": hw_id,
                        "hw_type": hw_type
                    })
        return undone_homework
LOGIN_URL = "https://lms.dgut.edu.cn/courseapi/users/login/v2"
HEADERS_LOGIN = {
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/137.0.0.0 Safari/537.36"
}
def do_login(loginname, password):
    session = requests.Session()
    form_data = {
        'loginName': loginname,
        'password': password
    }
    try:
        response = session.post(
            LOGIN_URL,
            data=form_data,
            headers=HEADERS_LOGIN,
            timeout=30
        )
        response.raise_for_status()
        if "courseweb" in response.url:
            userinfo = session.cookies.get('USERINFO')
            if not userinfo:
                return False, "登录成功，但未获取USERINFO Cookie"
            decoded_str = urllib.parse.unquote(userinfo)
            data_dict = json.loads(decoded_str)
            auth = data_dict.get("authorization", "")
            uid = data_dict.get("userId", "")
            return True, {"authorization": auth, "userId": uid, "raw": data_dict}
        else:
            return False, "登录失败，未跳转至课程页面，请检查账号密码"
    except requests.exceptions.Timeout:
        return False, "登录请求超时"
    except requests.exceptions.ConnectionError:
        return False, "网络连接错误"
    except Exception as e:
        return False, f"登录异常:{str(e)}"
class ToolTip:
    """悬浮提示工具"""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.widget.bind("<Enter>", self.show)
        self.widget.bind("<Leave>", self.hide)
    def show(self, event=None):
        if self.tip_window:
            return
        x = event.x_root
        y = event.y_root
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.update_idletasks()
        pos_x = x - tw.winfo_width()
        pos_y = y + 15
        tw.geometry(f"+{pos_x}+{pos_y}")
        label = ttk.Label(tw, text=self.text, background="#ffffe0", wraplength=400)
        label.pack()
    def hide(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None
class HomeworkApp:
    CONFIG_FILE = "config.json"
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("优学院作业小帮手")
        self.root.geometry("800x820")
        self.root.resizable(True, True)
        self.style = ttk.Style()
        self.style.configure("My.TButton", font=("微软雅黑", 11))
        self.current_download_tasks = []
        self.current_save_dir = ""
        self.homework_data = []
        self._current_detail_text = ""
        self._current_hw_name = ""
        self.user_id = ""
        self.cfg_data = self.load_config()
        self._build_ui()
        self.fill_config_to_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _set_text_readonly(self):
        """设置详情文本框只读"""
        self.detail_text.config(state=tk.DISABLED)

    def _text_insert(self, content):
        """向只读Text插入内容的封装：临时开启写入，写完立刻禁用"""
        self.detail_text.config(state=tk.NORMAL)
        self.detail_text.insert(tk.END, content)
        self.detail_text.config(state=tk.DISABLED)

    def _text_delete(self, start, end):
        """清空只读Text组件"""
        self.detail_text.config(state=tk.NORMAL)
        self.detail_text.delete(start, end)
        self.detail_text.config(state=tk.DISABLED)

    def load_config(self):
        if os.path.exists(self.CONFIG_FILE):
            try:
                with open(self.CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {"token": "", "account": "", "password": "", "userid": ""}
    def save_config_to_file(self, token, account, password, userid):
        data = {
            "token": token,
            "account": account,
            "password": password,
            "userid": userid
        }
        with open(self.CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    def fill_config_to_ui(self):
        self.token_entry.delete(0, tk.END)
        self.token_entry.insert(0, self.cfg_data.get("token", ""))
        self.account_entry.delete(0, tk.END)
        self.account_entry.insert(0, self.cfg_data.get("account", ""))
        self.password_entry.delete(0, tk.END)
        self.password_entry.insert(0, self.cfg_data.get("password", ""))
        self.user_id = self.cfg_data.get("userid", "")
    def on_closing(self):
        token = self.token_entry.get().strip()
        acc = self.account_entry.get().strip()
        pwd = self.password_entry.get().strip()
        ans = messagebox.askyesno("退出确认", "是否保存当前输入框参数到本地config.json？\n（密码会明文保存在文件）")
        if ans:
            self.save_config_to_file(token, acc, pwd, self.user_id)
        self.root.destroy()
    def _build_ui(self):
        # 主布局使用grid实现自适应，保证状态栏固定在底部
        self.root.grid_rowconfigure(0, weight=0)
        self.root.grid_rowconfigure(1, weight=0)
        self.root.grid_rowconfigure(2, weight=0)
        self.root.grid_rowconfigure(3, weight=0)
        self.root.grid_rowconfigure(4, weight=1)
        self.root.grid_rowconfigure(5, weight=0)
        self.root.grid_columnconfigure(0, weight=1)

        top_frame = ttk.Frame(self.root, padding=(12, 12))
        top_frame.grid(row=0, column=0, sticky="nsew")

        # ==========【第一行：账号密码行】==========
        row_acc_pwd = ttk.Frame(top_frame)
        row_acc_pwd.pack(fill=tk.X, pady=2)
        ttk.Label(row_acc_pwd, text="账号：", font=("微软雅黑", 11), width=10).pack(side=tk.LEFT)
        self.account_entry = ttk.Entry(row_acc_pwd, font=("微软雅黑", 11))
        self.account_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Label(row_acc_pwd, text="密码：", font=("微软雅黑", 11), width=8).pack(side=tk.LEFT)
        self.password_entry = ttk.Entry(row_acc_pwd, font=("微软雅黑", 11), width=28, show="*")
        self.password_entry.pack(side=tk.LEFT)

        # ==========【第二行：Token行，长度和账号框对齐；灯泡+登录按钮整体右对齐】==========
        row_token = ttk.Frame(top_frame)
        row_token.pack(fill=tk.X, pady=2)
        ttk.Label(row_token, text="Token：", font=("微软雅黑", 11), width=10).pack(side=tk.LEFT)
        self.token_entry = ttk.Entry(row_token, font=("微软雅黑", 11))
        self.token_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 右侧容器：灯泡+登录按钮，整体右对齐
        right_btn_container = ttk.Frame(row_token)
        right_btn_container.pack(side=tk.RIGHT)
        tip_label_bulb = ttk.Label(right_btn_container, text="💡", font=("微软雅黑", 12))
        tip_label_bulb.pack(side=tk.LEFT, padx=(0,6))
        self.login_btn = ttk.Button(right_btn_container, text="登录", style="My.TButton", command=self._do_login_btn)
        self.login_btn.pack(side=tk.LEFT)

        ToolTip(tip_label_bulb,
                "使用提示：\n1.填入账号密码点登录自动获取Token\n2.点击扫描作业获取待完成作业列表\n3.双击表格行加载作业详情\n4.保存详情导出文本，下载按钮保存附件图片")

        # ========== 待完成作业标题行：左文字，右侧【扫描作业】按钮 ==========
        hw_title_row = ttk.Frame(self.root, padding=(12,0))
        hw_title_row.grid(row=1, column=0, sticky="ew")
        ttk.Label(hw_title_row, text="待完成作业", font=("微软雅黑",12,"bold")).pack(side=tk.LEFT)
        self.scan_btn = ttk.Button(hw_title_row, text="扫描作业", style="My.TButton", command=self._start_scan)
        self.scan_btn.pack(side=tk.RIGHT)

        # Treeview表格容器，自适应拉伸
        tree_container = ttk.Frame(self.root, padding=(12,6))
        tree_container.grid(row=2, column=0, sticky="nsew")
        tree_container.grid_columnconfigure(0, weight=1)
        tree_container.grid_rowconfigure(0, weight=1)

        columns = ("course", "homework", "state", "deadline")
        self.result_tree = ttk.Treeview(
            tree_container,
            columns=columns,
            show="headings",
            selectmode="browse",
            height=10
        )
        self.result_tree.heading("course", text="课程名称")
        self.result_tree.heading("homework", text="作业名称")
        self.result_tree.heading("state", text="状态")
        self.result_tree.heading("deadline", text="截止时间")
        self.result_tree.column("course", width=220, anchor="w")
        self.result_tree.column("homework", width=220, anchor="w")
        self.result_tree.column("state", width=60, anchor="center")
        self.result_tree.column("deadline", width=170, anchor="center")
        style = ttk.Style()
        style.configure("Treeview", font=("微软雅黑", 10), rowheight=26)
        style.configure("Treeview.Heading", font=("微软雅黑", 10, "bold"))
        scroll_y_tree = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=scroll_y_tree.set)
        scroll_y_tree.grid(row=0, column=1, sticky="ns")
        self.result_tree.grid(row=0, column=0, sticky="nsew")

        # 作业详情标题行：保存作业详情 → 下载附件&图片
        detail_header_frame = ttk.Frame(self.root, padding=(12,0))
        detail_header_frame.grid(row=3, column=0, sticky="ew")
        ttk.Label(detail_header_frame, text="作业详情", font=("微软雅黑", 12, "bold")).pack(side=tk.LEFT)
        self.download_btn = ttk.Button(detail_header_frame, text="下载附件&图片", style="My.TButton", command=self._do_download, state=tk.DISABLED)
        self.save_detail_btn = ttk.Button(detail_header_frame, text="保存作业详情", style="My.TButton", command=self._save_detail_txt, state=tk.DISABLED)
        self.download_btn.pack(side=tk.RIGHT)
        self.save_detail_btn.pack(side=tk.RIGHT, padx=(0,8))

        # 详情文本框容器，占用剩余全部窗口高度
        detail_frame = ttk.Frame(self.root, padding=(12,6))
        detail_frame.grid(row=4, column=0, sticky="nsew")
        detail_frame.grid_columnconfigure(0, weight=1)
        detail_frame.grid_rowconfigure(0, weight=1)

        self.detail_text = tk.Text(detail_frame, wrap=tk.WORD, font=("微软雅黑", 12))
        self._set_text_readonly()
        detail_scroll = ttk.Scrollbar(detail_frame, orient=tk.VERTICAL, command=self.detail_text.yview)
        self.detail_text.configure(yscrollcommand=detail_scroll.set)
        detail_scroll.grid(row=0, column=1, sticky="ns")
        self.detail_text.grid(row=0, column=0, sticky="nsew")

        # 状态栏固定在row5，weight=0永远保留高度，不会被挤压
        self.status_text = tk.StringVar(value="就绪")
        status_bar = ttk.Label(
            self.root,
            textvariable=self.status_text,
            anchor="w",
            padding=(12, 6),
            relief=tk.SUNKEN,
            font=("微软雅黑", 9)
        )
        status_bar.grid(row=5, column=0, sticky="ew")

        self.result_tree.bind("<Double-1>", self._on_double_click_row)

    def _do_login_btn(self):
        account = self.account_entry.get().strip()
        password = self.password_entry.get().strip()
        if not account or not password:
            messagebox.showwarning("提示", "账号和密码不能为空！")
            return
        self.login_btn.config(state=tk.DISABLED)
        self.status_text.set("正在登录...")
        self.root.update_idletasks()
        ok, res = do_login(account, password)
        self.login_btn.config(state=tk.NORMAL)
        if ok:
            token = res["authorization"]
            self.user_id = res["userId"]  # 存入内存变量，无UI
            self.token_entry.delete(0, tk.END)
            self.token_entry.insert(0, token)
            self.status_text.set("登录成功，Token已自动填充")
            messagebox.showinfo("登录成功", "账号密码登录成功，已自动填入Token")
        else:
            self.status_text.set("登录失败")
            messagebox.showerror("登录失败", res)
    def _render_sorted_table(self):
        sorted_list = sorted(self.homework_data, key=lambda x: x["截止时间戳"])
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        for hw in sorted_list:
            self.result_tree.insert("", tk.END, values=(
                hw["课程名称"],
                hw["作业名称"],
                hw["作业状态"],
                hw["截止时间"]
            ))
    def _save_homework_list_xlsx(self):
        if Workbook is None:
            messagebox.showwarning("缺少依赖", "需要安装openpyxl，执行 pip install openpyxl")
            return
        sorted_list = sorted(self.homework_data, key=lambda x: x["截止时间戳"])
        now = datetime.now()
        filename = f"作业列表_{now.strftime('%Y%m%d_%H%M')}.xlsx"
        wb = Workbook()
        ws = wb.active
        headers = ["课程名称","作业名称","作业状态","截止时间"]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for hw in sorted_list:
            row = [hw["课程名称"],hw["作业名称"],hw["作业状态"],hw["截止时间"]]
            ws.append(row)
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max_len + 3
        wb.save(filename)
        self.status_text.set(self.status_text.get() + f"｜已导出 {filename}")
    def _start_scan(self):
        token = self.token_entry.get().strip()
        if not token:
            messagebox.showwarning("输入提示", "请输入登录 Token（可以账号密码登录自动获取）")
            return
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        self._text_delete(1.0, tk.END)
        self.current_download_tasks.clear()
        self.current_save_dir = ""
        self.homework_data.clear()
        self._current_detail_text = ""
        self._current_hw_name = ""
        self.download_btn.config(state=tk.DISABLED)
        self.save_detail_btn.config(state=tk.DISABLED)
        self.scan_btn.config(state=tk.DISABLED)
        self.status_text.set("正在扫描课程与作业，请稍候...")
        self.root.update()
        try:
            checker = HomeworkChecker(token, self.user_id)
            undone_list = checker.get_undone_homework()
            self.homework_data = undone_list
            if not undone_list:
                self.status_text.set("扫描完成，暂无未完成作业 🎉")
                messagebox.showinfo("扫描结果", "恭喜！所有作业都已完成~")
            else:
                self._render_sorted_table()
                self.status_text.set(f"扫描完成，共找到 {len(undone_list)} 项待完成作业")
                self._save_homework_list_xlsx()
        except Exception as e:
            self.status_text.set("扫描失败")
            error_msg = str(e)
            if "401" in error_msg or "403" in error_msg:
                error_msg = "Token 无效或已过期，请重新获取或重新登录"
            messagebox.showerror("扫描出错", f"发生错误：{error_msg}")
        finally:
            self.scan_btn.config(state=tk.NORMAL)
    def _on_double_click_row(self, event):
        selected = self.result_tree.selection()
        if not selected:
            return
        item = self.result_tree.item(selected[0])
        vals = item["values"]
        course_name = vals[0]
        hw_name_show = vals[1]
        target_hw = None
        for hw in self.homework_data:
            if hw["课程名称"] == course_name and hw["作业名称"] == hw_name_show:
                target_hw = hw
                break
        if target_hw is None:
            messagebox.showerror("错误", "找不到对应作业元数据，请重新扫描")
            return
        course_id = target_hw["课程ID"]
        homework_id = target_hw["作业ID"]
        hw_type = target_hw["hw_type"]
        token = self.token_entry.get().strip()
        if not token:
            messagebox.showwarning("提示", "Token为空，无法获取作业详情！")
            return
        self.status_text.set("正在加载作业详情...")
        self.root.update_idletasks()
        checker = HomeworkChecker(token, self.user_id)
        ok, activity_homework = checker.get_homework_detail_raw(homework_id, course_id, hw_type)
        self._text_delete(1.0, tk.END)
        self.current_download_tasks.clear()
        self.current_save_dir = ""
        self._current_detail_text = ""
        self._current_hw_name = ""
        if not ok:
            self._text_insert(activity_homework)
            self.download_btn.config(state=tk.DISABLED)
            self.save_detail_btn.config(state=tk.DISABLED)
            self.status_text.set("详情加载失败")
            return
        preview_text, save_dir, tasks = checker.parse_preview_content(activity_homework, course_name, hw_type)
        self._text_insert(preview_text)
        self.current_download_tasks = tasks
        self.current_save_dir = save_dir
        self._current_detail_text = preview_text
        self._current_hw_name = target_hw["作业名称"]
        self.download_btn.config(state=tk.NORMAL)
        self.save_detail_btn.config(state=tk.NORMAL)
        type_desc = "小组作业" if hw_type == 8 else "个人作业"
        self.status_text.set(f"详情加载完成[{type_desc}]；清单已展示，点击下载按钮执行文件下载")
    def _save_detail_txt(self):
        if not self._current_detail_text or not self._current_hw_name or not self.current_save_dir:
            messagebox.showwarning("提示", "请先双击选中作业加载详情！")
            return
        os.makedirs(self.current_save_dir, exist_ok=True)
        safe_hw_name = re.sub(r'[\\/*?:"<>|]', '_', self._current_hw_name)
        save_file = os.path.join(self.current_save_dir, f"{safe_hw_name}_request.txt")
        try:
            with open(save_file, "w", encoding="utf-8") as f:
                f.write(self._current_detail_text)
            messagebox.showinfo("保存成功", f"详情已保存：{save_file}")
            self.status_text.set(f"作业详情已保存至 {save_file}")
        except Exception as e:
            messagebox.showerror("保存失败", f"写入文件失败：{str(e)}")
    def _do_download(self):
        if not self.current_download_tasks:
            messagebox.showwarning("提示", "请先双击选中一行作业加载详情与资源清单！")
            return
        token = self.token_entry.get().strip()
        if not token:
            messagebox.showwarning("提示", "Token为空！")
            return
        self.download_btn.config(state=tk.DISABLED)
        self.scan_btn.config(state=tk.DISABLED)
        self.root.update_idletasks()
        os.makedirs(self.current_save_dir, exist_ok=True)
        checker = HomeworkChecker(token, self.user_id)
        tasks = self.current_download_tasks
        downloaded_log = []
        total = len(tasks)
        try:
            if total == 0:
                self.status_text.set("没有可下载的图片或附件")
                downloaded_log.append("本次无图片/附件需要下载")
            else:
                for idx, task in enumerate(tasks, start=1):
                    self.status_text.set(f"正在下载 ({idx}/{total}) → {task['display_name']}")
                    self.root.update_idletasks()
                    ok, err = checker._download_file(task["url"], task["save_path"])
                    if ok:
                        downloaded_log.append(f"✅ {task['display_name']}")
                    else:
                        downloaded_log.append(f"❌ {task['display_name']} 失败:{err}")
                self.status_text.set(f"下载全部处理完毕，输出目录：{self.current_save_dir}")
        except Exception as e:
            downloaded_log.append(f"下载发生异常：{str(e)}")
            self.status_text.set("下载过程出错")
        finally:
            log_str = "\n\n=====【本次下载执行结果】=====\n" + "\n".join(downloaded_log)
            self._text_insert(log_str)
            self.download_btn.config(state=tk.NORMAL)
            self.scan_btn.config(state=tk.NORMAL)
def main():
    root = tk.Tk()
    app = HomeworkApp(root)
    root.update_idletasks()   # 先让窗口渲染，拿到真实窗口宽高
    win_w = root.winfo_width()
    win_h = root.winfo_height()
    screen_w = root.winfo_screenwidth()
    x = (screen_w - win_w) // 2
    y = 0   # y=0紧贴屏幕顶部
    root.geometry(f"{win_w}x{win_h}+{x}+{y}")
    root.mainloop()

if __name__ == "__main__":
    main()