import datetime
import mimetypes
import os
import queue
import re
import threading
import tkinter as tk
import tkinter.font as tkfont
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Cm as word_Cm
from PIL import Image, UnidentifiedImageError
from pptx import Presentation
from pptx.util import Cm as ppt_Cm
from tkinter import filedialog, messagebox, scrolledtext, ttk

USER_AGENT = (
    "Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) "
    "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.0 "
    "Mobile/15E148 Safari/604.1"
)
REQUEST_TIMEOUT = 20
MAX_DOWNLOAD_WORKERS = 5
MAX_DOWNLOAD_RETRIES = 3
MAX_FILENAME_LENGTH = 80
DEFAULT_IMAGE_EXTENSION = "jpg"
WINDOW_WIDTH = 760
WINDOW_HEIGHT = 760

APPLE_BLUE = "#0071E3"
APPLE_BLUE_ACTIVE = "#0077ED"
APPLE_BG = "#F5F5F7"
APPLE_CARD = "#FFFFFF"
APPLE_TEXT = "#1D1D1F"
APPLE_MUTED = "#6E6E73"
APPLE_BORDER = "#D2D2D7"
APPLE_INPUT_BG = "#FBFBFD"
APPLE_STATUS_BG = "#FAFAFC"
APPLE_SUCCESS = "#14833B"
APPLE_WARNING = "#A15C00"
APPLE_ERROR = "#C9342F"


def sanitize_filename(name: str) -> str:
    cleaned = re.sub(r'[\\/*?:"<>|]', "_", name or "")
    cleaned = cleaned.replace("..", "_").strip(" ._")
    if not cleaned:
        cleaned = "公众号素材"
    return cleaned[:MAX_FILENAME_LENGTH]


def is_http_url(value: str) -> bool:
    parsed = urlparse(value)
    return parsed.scheme in {"http", "https"} and bool(parsed.netloc)


def post_event(status_queue: queue.Queue, event_type: str, **payload: object) -> None:
    event = {"type": event_type}
    event.update(payload)
    status_queue.put(event)


def log_status(status_queue: queue.Queue, message: str, level: str = "info") -> None:
    post_event(status_queue, "log", message=message, level=level)


def create_session_folder(base_path: str) -> str:
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    folder_path = os.path.join(base_path, f"微信素材_{timestamp}")
    os.makedirs(folder_path, exist_ok=False)
    return folder_path


def build_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})
    return session


def extract_image_urls(page_url: str, html: bytes) -> list[str]:
    soup = BeautifulSoup(html, "lxml")
    attrs = ("data-src", "data-actualsrc", "data-original", "data-lazy-src", "src")
    found_urls: list[str] = []
    seen: set[str] = set()

    for img in soup.find_all("img"):
        for attr in attrs:
            raw_url = img.get(attr)
            if not raw_url:
                continue
            image_url = urljoin(page_url, raw_url.strip())
            if not is_http_url(image_url) or image_url in seen:
                continue
            seen.add(image_url)
            found_urls.append(image_url)
            break

    return found_urls


def choose_extension(content_type: str, image_bytes: bytes) -> str:
    extension = ""
    if content_type:
        guessed = mimetypes.guess_extension(content_type.split(";")[0].strip())
        if guessed:
            extension = guessed.lstrip(".")

    if not extension:
        try:
            with Image.open(BytesIO(image_bytes)) as image:
                extension = (image.format or "").lower()
        except (UnidentifiedImageError, OSError):
            extension = DEFAULT_IMAGE_EXTENSION

    if extension == "jpeg":
        return "jpg"
    return extension or DEFAULT_IMAGE_EXTENSION


def download_single_image(
    session: requests.Session,
    image_url: str,
    save_folder: str,
    index: int,
) -> str:
    last_error: Exception | None = None
    for attempt in range(1, MAX_DOWNLOAD_RETRIES + 1):
        try:
            response = session.get(image_url, timeout=REQUEST_TIMEOUT)
            response.raise_for_status()
            image_bytes = response.content
            extension = choose_extension(response.headers.get("Content-Type", ""), image_bytes)
            file_name = f"{index:03d}.{extension}"
            file_path = os.path.join(save_folder, file_name)
            with open(file_path, "wb") as file_obj:
                file_obj.write(image_bytes)
            return file_path
        except (requests.RequestException, OSError) as exc:
            last_error = exc
            if attempt == MAX_DOWNLOAD_RETRIES:
                break
    if last_error:
        raise last_error
    raise RuntimeError("图片下载失败")


def download_images_from_url(article_url: str, save_folder: str, status_queue: queue.Queue) -> list[str]:
    session = build_session()
    try:
        log_status(status_queue, "正在获取文章页面")
        response = session.get(article_url, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
    except requests.RequestException as exc:
        log_status(status_queue, f"请求失败：{exc}", "error")
        return []

    image_urls = extract_image_urls(article_url, response.content)
    if not image_urls:
        log_status(status_queue, "未发现可下载的图片", "warning")
        return []

    log_status(status_queue, f"共发现 {len(image_urls)} 张图片")
    downloaded_files: list[str] = []
    completed_count = 0
    total_count = len(image_urls)
    post_event(status_queue, "progress", value=0, total=total_count, current=0)

    with ThreadPoolExecutor(max_workers=MAX_DOWNLOAD_WORKERS) as executor:
        future_map = {
            executor.submit(download_single_image, session, image_url, save_folder, index): image_url
            for index, image_url in enumerate(image_urls, start=1)
        }
        for future in as_completed(future_map):
            image_url = future_map[future]
            try:
                file_path = future.result()
                downloaded_files.append(file_path)
                log_status(status_queue, f"已下载 {os.path.basename(file_path)}")
            except OSError as exc:
                log_status(status_queue, f"保存图片失败：{exc}", "warning")
            except requests.RequestException as exc:
                log_status(
                    status_queue,
                    f"已跳过图片：{image_url}，已重试 {MAX_DOWNLOAD_RETRIES} 次，原因：{exc}",
                    "warning",
                )
            finally:
                completed_count += 1
                percent = int(completed_count / total_count * 100)
                post_event(
                    status_queue,
                    "progress",
                    value=percent,
                    total=total_count,
                    current=completed_count,
                )

    downloaded_files.sort()
    log_status(status_queue, f"图片下载完成，共 {len(downloaded_files)} 个文件")
    return downloaded_files


def generate_word_document(prefix: str, image_paths: list[str], folder: str, status_queue: queue.Queue) -> None:
    if not image_paths:
        return

    log_status(status_queue, "正在生成 Word 文档")
    document = Document()
    for image_path in image_paths:
        try:
            document.add_picture(image_path, width=word_Cm(16))
        except Exception as exc:
            log_status(status_queue, f"Word 跳过 {os.path.basename(image_path)}：{exc}", "warning")

    file_path = os.path.join(folder, f"{sanitize_filename(prefix)}.docx")
    document.save(file_path)
    log_status(status_queue, f"已保存 {os.path.basename(file_path)}", "success")


def fit_image_to_slide(image_width: int, image_height: int) -> tuple[float, float, float, float]:
    slide_width = 25.4
    slide_height = 19.05
    margin = 1.0
    max_width = slide_width - margin * 2
    max_height = slide_height - margin * 2
    scale = min(max_width / image_width, max_height / image_height)
    final_width = image_width * scale
    final_height = image_height * scale
    left = (slide_width - final_width) / 2
    top = (slide_height - final_height) / 2
    return left, top, final_width, final_height


def generate_ppt_presentation(prefix: str, image_paths: list[str], folder: str, status_queue: queue.Queue) -> None:
    if not image_paths:
        return

    log_status(status_queue, "正在生成 PowerPoint")
    presentation = Presentation()
    while presentation.slides:
        rel_id = presentation.slides._sldIdLst[0].rId
        presentation.part.drop_rel(rel_id)
        del presentation.slides._sldIdLst[0]

    for image_path in image_paths:
        try:
            with Image.open(image_path) as image:
                width, height = image.size
            left, top, final_width, final_height = fit_image_to_slide(width, height)
            slide = presentation.slides.add_slide(presentation.slide_layouts[6])
            slide.shapes.add_picture(
                image_path,
                ppt_Cm(left),
                ppt_Cm(top),
                width=ppt_Cm(final_width),
                height=ppt_Cm(final_height),
            )
        except Exception as exc:
            log_status(status_queue, f"PPT 跳过 {os.path.basename(image_path)}：{exc}", "warning")

    file_path = os.path.join(folder, f"{sanitize_filename(prefix)}.pptx")
    presentation.save(file_path)
    log_status(status_queue, f"已保存 {os.path.basename(file_path)}", "success")


class AppleStyleApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("微信公众号素材整理工具")
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.root.minsize(680, 680)
        self.root.configure(bg=APPLE_BG)

        self.selected_path = ""
        self.status_queue: queue.Queue = queue.Queue()
        self.is_running = False
        self.progress_value = tk.IntVar(value=0)

        self.font_title = tkfont.Font(family="Microsoft YaHei UI", size=22, weight="bold")
        self.font_header = tkfont.Font(family="Microsoft YaHei UI", size=11, weight="bold")
        self.font_body = tkfont.Font(family="Microsoft YaHei UI", size=10)
        self.font_input = tkfont.Font(family="Microsoft YaHei UI", size=11)
        self.font_button = tkfont.Font(family="Microsoft YaHei UI", size=10, weight="bold")
        self.font_mono = tkfont.Font(family="Consolas", size=9)

        self._setup_styles()
        self._build_ui()
        self._poll_status_queue()

    def _setup_styles(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("Card.TFrame", background=APPLE_CARD, relief="flat")
        style.configure("Title.TLabel", background=APPLE_BG, foreground=APPLE_TEXT, font=self.font_title)
        style.configure("Muted.TLabel", background=APPLE_CARD, foreground=APPLE_MUTED, font=self.font_body)
        style.configure("Header.TLabel", background=APPLE_CARD, foreground=APPLE_TEXT, font=self.font_header)
        style.configure(
            "Secondary.TButton",
            background=APPLE_INPUT_BG,
            foreground=APPLE_TEXT,
            bordercolor=APPLE_BORDER,
            lightcolor=APPLE_INPUT_BG,
            darkcolor=APPLE_INPUT_BG,
            focuscolor="none",
            font=self.font_body,
            padding=(12, 8),
        )

    def _build_ui(self) -> None:
        shell = tk.Frame(self.root, bg=APPLE_BG)
        shell.pack(fill="both", expand=True, padx=32, pady=28)

        hero = tk.Frame(shell, bg=APPLE_BG)
        hero.pack(fill="x", pady=(0, 18))

        ttk.Label(hero, text="微信公众号素材整理工具", style="Title.TLabel").pack(anchor="center")
        tk.Label(
            hero,
            text="提取文章图片，并自动整理为 Word 与 PowerPoint 文件。",
            bg=APPLE_BG,
            fg=APPLE_MUTED,
            font=self.font_body,
        ).pack(anchor="center", pady=(8, 0))

        card = ttk.Frame(shell, style="Card.TFrame", padding=28)
        card.pack(fill="both", expand=True)

        ttk.Label(card, text="文章链接", style="Header.TLabel").pack(anchor="w")
        self.url_entry = self._create_entry(card)
        self.url_entry.pack(fill="x", pady=(8, 20), ipady=10)

        ttk.Label(card, text="文件名前缀", style="Header.TLabel").pack(anchor="w")
        self.prefix_entry = self._create_entry(card)
        self.prefix_entry.insert(0, "公众号素材")
        self.prefix_entry.pack(fill="x", pady=(8, 20), ipady=10)

        location_row = tk.Frame(card, bg=APPLE_CARD)
        location_row.pack(fill="x")
        ttk.Label(location_row, text="保存位置", style="Header.TLabel").pack(side="left")
        ttk.Button(location_row, text="选择文件夹", style="Secondary.TButton", command=self._choose_path).pack(side="right")

        self.path_label = ttk.Label(card, text="尚未选择文件夹", style="Muted.TLabel")
        self.path_label.pack(anchor="w", pady=(8, 20))

        option_row = tk.Frame(card, bg=APPLE_CARD)
        option_row.pack(fill="x", pady=(0, 16))

        self.gen_word_var = tk.BooleanVar(value=True)
        self.gen_ppt_var = tk.BooleanVar(value=True)
        self._create_checkbox(option_row, "Word 文档", self.gen_word_var).pack(side="left", padx=(0, 24))
        self._create_checkbox(option_row, "PowerPoint 演示文稿", self.gen_ppt_var).pack(side="left")

        progress_row = tk.Frame(card, bg=APPLE_CARD)
        progress_row.pack(fill="x", pady=(0, 12))
        ttk.Label(progress_row, text="运行状态", style="Header.TLabel").pack(side="left")
        self.state_label = ttk.Label(progress_row, text="空闲", style="Muted.TLabel")
        self.state_label.pack(side="right")

        self.progress = ttk.Progressbar(card, mode="determinate", maximum=100, variable=self.progress_value)
        self.progress.pack(fill="x", pady=(0, 16))

        action_bar = tk.Frame(card, bg=APPLE_CARD)
        action_bar.pack(fill="x", pady=(0, 16))

        self.run_button = tk.Button(
            action_bar,
            text="开始生成",
            command=self._start_task,
            font=self.font_button,
            bg=APPLE_BLUE,
            fg="#FFFFFF",
            activebackground=APPLE_BLUE_ACTIVE,
            activeforeground="#FFFFFF",
            relief="flat",
            bd=0,
            cursor="hand2",
            pady=12,
        )
        self.run_button.pack(fill="x")

        self.status_area = scrolledtext.ScrolledText(
            card,
            height=10,
            bg=APPLE_STATUS_BG,
            fg=APPLE_TEXT,
            insertbackground=APPLE_TEXT,
            relief="flat",
            borderwidth=1,
            highlightthickness=1,
            highlightbackground=APPLE_BORDER,
            highlightcolor=APPLE_BORDER,
            font=self.font_mono,
            padx=12,
            pady=12,
            state="disabled",
        )
        self.status_area.pack(fill="both", expand=True)


    def _create_entry(self, parent: tk.Widget) -> tk.Entry:
        return tk.Entry(
            parent,
            font=self.font_input,
            bg=APPLE_INPUT_BG,
            fg=APPLE_TEXT,
            insertbackground=APPLE_TEXT,
            relief="flat",
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=APPLE_BORDER,
            highlightcolor=APPLE_BLUE,
        )

    def _create_checkbox(self, parent: tk.Widget, text: str, variable: tk.BooleanVar) -> tk.Checkbutton:
        return tk.Checkbutton(
            parent,
            text=text,
            variable=variable,
            bg=APPLE_CARD,
            fg=APPLE_TEXT,
            activebackground=APPLE_CARD,
            activeforeground=APPLE_TEXT,
            selectcolor=APPLE_CARD,
            font=self.font_body,
            highlightthickness=0,
            bd=0,
        )

    def _choose_path(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            self.selected_path = folder
            self.path_label.config(text=folder)

    def _append_status(self, message: str, level: str) -> None:
        colors = {
            "info": APPLE_TEXT,
            "success": APPLE_SUCCESS,
            "warning": APPLE_WARNING,
            "error": APPLE_ERROR,
        }
        tag_name = f"level_{level}"
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.status_area.config(state="normal")
        self.status_area.tag_configure(tag_name, foreground=colors.get(level, APPLE_TEXT))
        self.status_area.insert("end", f"[{timestamp}] {message}\n", tag_name)
        self.status_area.see("end")
        self.status_area.config(state="disabled")

    def _set_running_state(self, running: bool) -> None:
        self.is_running = running
        if running:
            self.run_button.config(state="disabled", bg="#B9D3F4", cursor="arrow")
            self.progress_value.set(0)
            self.state_label.config(text="准备中 0%")
        else:
            self.run_button.config(state="normal", bg=APPLE_BLUE, cursor="hand2")
            self.state_label.config(text="空闲")

    def _update_progress(self, value: int, current: int, total: int) -> None:
        self.progress_value.set(max(0, min(100, value)))
        if total > 0 and current < total:
            self.state_label.config(text=f"下载中 {value}% ({current}/{total})")
        elif total > 0:
            self.state_label.config(text=f"处理中 {value}% ({current}/{total})")
        else:
            self.state_label.config(text=f"处理中 {value}%")

    def _poll_status_queue(self) -> None:
        try:
            while True:
                item = self.status_queue.get_nowait()
                if item.get("type") == "progress":
                    self._update_progress(
                        int(item.get("value", 0)),
                        int(item.get("current", 0)),
                        int(item.get("total", 0)),
                    )
                else:
                    message = str(item["message"])
                    if message.startswith("正在生成 Word"):
                        self.state_label.config(text="正在生成 Word")
                    elif message.startswith("正在生成 PowerPoint"):
                        self.state_label.config(text="正在生成 PowerPoint")
                    elif message == "导出完成":
                        self.state_label.config(text="已完成")
                    self._append_status(message, str(item["level"]))
        except queue.Empty:
            pass
        self.root.after(120, self._poll_status_queue)

    def _validate_inputs(self) -> tuple[str, str] | None:
        article_url = self.url_entry.get().strip()
        prefix = self.prefix_entry.get().strip()

        if not is_http_url(article_url):
            messagebox.showerror("链接错误", "请输入有效的 http 或 https 文章链接。")
            return None
        if not self.selected_path or not os.path.isdir(self.selected_path):
            messagebox.showerror("保存位置错误", "请选择一个已存在的保存文件夹。")
            return None
        if not prefix:
            messagebox.showerror("文件名前缀", "请输入导出文件名前缀。")
            return None
        if not self.gen_word_var.get() and not self.gen_ppt_var.get():
            messagebox.showerror("输出格式", "请至少选择一种输出格式。")
            return None
        return article_url, prefix

    def _start_task(self) -> None:
        if self.is_running:
            return

        validated = self._validate_inputs()
        if not validated:
            return

        article_url, prefix = validated
        self.status_area.config(state="normal")
        self.status_area.delete("1.0", "end")
        self.status_area.config(state="disabled")
        self._set_running_state(True)
        threading.Thread(target=self._worker, args=(article_url, prefix), daemon=True).start()

    def _worker(self, article_url: str, prefix: str) -> None:
        try:
            log_status(self.status_queue, "开始导出")
            session_folder = create_session_folder(self.selected_path)
            log_status(self.status_queue, f"输出目录：{session_folder}")

            image_paths = download_images_from_url(article_url, session_folder, self.status_queue)
            if not image_paths:
                log_status(self.status_queue, "未下载到图片，未生成导出文件", "warning")
                return

            if self.gen_word_var.get():
                generate_word_document(prefix, image_paths, session_folder, self.status_queue)
            if self.gen_ppt_var.get():
                generate_ppt_presentation(prefix, image_paths, session_folder, self.status_queue)

            log_status(self.status_queue, "导出完成", "success")
        except FileExistsError:
            log_status(self.status_queue, "输出目录已存在，请重新执行导出", "error")
        except Exception as exc:
            log_status(self.status_queue, f"发生未预期错误：{exc}", "error")
        finally:
            self.root.after(0, lambda: self._set_running_state(False))


if __name__ == "__main__":
    root = tk.Tk()
    app = AppleStyleApp(root)
    root.mainloop()
