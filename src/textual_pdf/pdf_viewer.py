import io
import mimetypes
import tempfile
from pathlib import Path
import fitz
import textual_image.widget as timg
from PIL import Image as PILImage, ImageDraw, ImageFont
from pymupdf import EmptyFileError, FileDataError
from textual import events, work
from textual.app import ComposeResult
from textual.containers import Container
from textual.reactive import reactive
from textual.widgets import Label
from markdown import markdown
from bs4 import BeautifulSoup
import textwrap

from .exceptions import NotAPDFError, PDFHasAPasswordError, PDFRuntimeError


class PDFViewer(Container):
    """A universal document viewer widget supporting PDF, DOCX, TXT and Markdown."""

    DEFAULT_CSS = """
    PDFViewer {
        height: 1fr;
        width: 1fr;
        Image {
            width: auto;
            height: auto;
            align: center bottom;
        }
    }
    """

    current_page: reactive[int] = reactive(0)
    """The current page in the document. Starts from `0` until `total_pages - 1`"""
    protocol: reactive[str] = reactive("Auto")
    """Protocol to use ["Auto", "TGP", "Sixel", "Halfcell", "Unicode"]"""
    path: reactive[str | Path] = reactive("")  # ty: ignore[invalid-assignment]
    """Path to a document file"""
    total_pages: reactive[int] = reactive(1)
    """The total number of pages in the currently open document"""

    def __init__(
        self,
        path: str | Path,
        protocol: str = "",
        use_keys: bool = True,
        name: str | None = None,
        id: str | None = None,
        classes: str | None = None,
        font_path: str | None = None,
        font_size: int = 10,
    ) -> None:
        """Initialize the PDFViewer widget.

        Args:
            path(str): Path to a document file.
            protocol(str): The protocol to use (leave empty or 'Auto' to use auto protocol)
            use_keys(bool): Whether to use the default key assignments
            name(str): The name of this widget.
            id(str): The ID of the widget in the DOM.
            classes(str): The CSS classes for this widget.
            font_path(str): Optional path to a TTF font for text rendering.
            font_size(int): Font size for text document rendering.

        Raises:
            PDFHasAPasswordError: When the PDF file is password protected
            NotAPDFError: When the file is not a valid PDF or supported format
        """  # noqa: DOC502
        super().__init__(name=name, id=id, classes=classes, disabled=False, markup=True)
        assert protocol in ["Auto", "TGP", "Sixel", "Halfcell", "Unicode", ""]
        self.protocol = protocol
        self.path = path
        self.use_keys = use_keys
        self.font_path = font_path
        self.font_size = font_size
        self.doc = None
        self.doc_type = None
        self.pages_cache = []
        # _check_file is deferred to on_mount to avoid blocking the UI thread

    def _guess_type(self, path) -> str:
        """Guess the MIME type of the given path or BytesIO object.

        Args:
            path: A file path or BytesIO stream.

        Returns:
            A simplified type string such as 'pdf', 'plain', 'msword', or 'markdown'.
        """
        if isinstance(path, io.BytesIO):
            header = path.getvalue()[:4]
            if header.startswith(b"%PDF"):
                return "pdf"
            elif str(header).startswith("b'PK"):
                return "vnd.openxmlformats-officedocument.wordprocessingml.document"
            else:
                return "plain"

        path = Path(path)
        suffix = path.suffix.lower()
        if suffix in (".md", ".markdown"):
            return "markdown"
        if suffix in (".txt",):
            return "plain"
        if suffix in (".docx",):
            return "vnd.openxmlformats-officedocument.wordprocessingml.document"
        if suffix in (".doc",):
            return "msword"
        if suffix in (".pdf",):
            return "pdf"

        guess, _ = mimetypes.guess_type(path)
        if not guess:
            return "plain"
        mime_sub = guess.split("/")[-1]
        # Normalise text/markdown
        if mime_sub in ("markdown", "x-markdown"):
            return "markdown"
        return mime_sub

    def _check_file(self, path) -> None:
        """Check and load the document at the given path.

        Args:
            path: Path to the document or a BytesIO stream.

        Raises:
            NotAPDFError: When the file is not a valid or supported document.
            PDFHasAPasswordError: When the PDF file is password protected.
        """
        file_type = self._guess_type(path)
        self.doc_type = file_type

        if file_type == "pdf":
            try:
                self.doc = fitz.open(stream=path.getvalue(), filetype="pdf") if isinstance(
                    path, io.BytesIO) else fitz.open(path)
                if self.doc.is_encrypted and self.doc.needs_pass:
                    raise PDFHasAPasswordError(
                        f"{path} is a document that is encrypted, and cannot be read.")
                self.total_pages = self.doc.page_count
            except (FileDataError, EmptyFileError):
                raise NotAPDFError(f"{path} does not point to a valid PDF file.")

        elif file_type in ("plain", "txt"):
            text = path.getvalue().decode(errors="ignore") if isinstance(
                path, io.BytesIO) else Path(path).read_text(encoding="utf-8", errors="ignore")
            lines = text.splitlines()
            lines_per_page = 40
            self.doc = [lines[i:i + lines_per_page] for i in range(0, max(1, len(lines)), lines_per_page)]
            self.doc_type = "text_pages"
            self.total_pages = len(self.doc)

        elif file_type in ("msword", "vnd.openxmlformats-officedocument.wordprocessingml.document"):
            result = self._docx_to_fitz(path)
            if isinstance(result, fitz.Document):
                self.doc = result
                if self.doc.is_encrypted and self.doc.needs_pass:
                    raise PDFHasAPasswordError(f"{path} is a document that is encrypted, and cannot be read.")
                self.doc_type = "pdf"
                self.total_pages = self.doc.page_count
            else:
                # Fallback: list of page line-lists from mammoth text extraction
                self.doc = result
                self.doc_type = "text_pages"
                self.total_pages = len(self.doc)

        elif file_type in ("markdown", "md"):
            text = path.getvalue().decode(errors="ignore") if isinstance(
                path, io.BytesIO) else Path(path).read_text(encoding="utf-8", errors="ignore")
            html = markdown(text, extensions=["tables", "fenced_code"])
            self.doc = self._split_html_pages(html)
            self.doc_type = "text_pages"
            self.total_pages = len(self.doc)

        else:
            raise NotAPDFError(f"Unsupported format: {file_type}")

    def _docx_to_fitz(self, path) -> "fitz.Document | list":
        """Convert a DOCX file to a fitz PDF document, or fall back to text pages.

        Tries docx2pdf, then win32com (Word), then LibreOffice headless.
        If all PDF converters fail, falls back to mammoth HTML extraction rendered as text pages.

        Args:
            path: Path to the DOCX file or a BytesIO stream.

        Returns:
            fitz.Document if PDF conversion succeeded, or a list of page line-lists as fallback.

        Raises:
            NotAPDFError: only if even the mammoth fallback fails.
        """
        import shutil
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            if isinstance(path, io.BytesIO):
                docx_path = tmp_path / "input.docx"
                docx_path.write_bytes(path.getvalue())
            else:
                docx_path = Path(path).resolve()
            pdf_path = tmp_path / "output.pdf"

            # Try docx2pdf (Word COM on Windows)
            try:
                from docx2pdf import convert
                convert(str(docx_path), str(pdf_path))
                if pdf_path.exists():
                    return fitz.open(stream=pdf_path.read_bytes(), filetype="pdf")
            except Exception:
                pass

            # Try win32com directly (Word COM automation, Windows)
            try:
                import win32com.client
                import pythoncom
                pythoncom.CoInitialize()
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                doc = word.Documents.Open(str(docx_path.resolve()))
                doc.SaveAs(str(pdf_path.resolve()), FileFormat=17)  # 17 = wdFormatPDF
                doc.Close()
                word.Quit()
                pythoncom.CoUninitialize()
                if pdf_path.exists():
                    return fitz.open(stream=pdf_path.read_bytes(), filetype="pdf")
            except Exception:
                pass

            # Try LibreOffice headless
            soffice = shutil.which("soffice") or shutil.which("libreoffice")
            if soffice:
                import subprocess
                try:
                    subprocess.run(
                        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(tmp_path), str(docx_path)],
                        timeout=30, capture_output=True
                    )
                    candidates = list(tmp_path.glob("*.pdf"))
                    if candidates:
                        return fitz.open(stream=candidates[0].read_bytes(), filetype="pdf")
                except Exception:
                    pass

            # Fallback: mammoth HTML extraction → text pages
            try:
                import mammoth
                with open(docx_path, "rb") as f:
                    result = mammoth.convert_to_html(f)
                pages = self._split_html_pages(result.value)
                if pages:
                    return pages
            except Exception:
                pass

            raise NotAPDFError(
                f"Could not convert {docx_path.name} to PDF and text extraction also failed."
            )

    def _split_html_pages(self, html: str) -> list:
        """Convert HTML into a list of page line-lists, preserving structure.

        Handles headings, paragraphs, lists, tables, code blocks, and horizontal rules.

        Args:
            html: HTML string.

        Returns:
            A list of page content, each being a list of strings.
        """
        soup = BeautifulSoup(html, "html.parser")
        all_lines = []

        for el in soup.find_all(
            ["p", "h1", "h2", "h3", "h4", "h5", "h6", "table", "ul", "ol", "pre", "hr", "blockquote"]
        ):
            if el.name == "table":
                all_lines.extend(self._table_to_lines(el))
                all_lines.append("")

            elif el.name in ("ul", "ol"):
                for i, li in enumerate(el.find_all("li", recursive=False)):
                    prefix = f"  {i + 1}." if el.name == "ol" else "  •"
                    text = li.get_text(separator=" ", strip=True)
                    wrapped = textwrap.wrap(text, width=85)
                    if wrapped:
                        all_lines.append(f"{prefix} {wrapped[0]}")
                        for cont in wrapped[1:]:
                            all_lines.append(f"      {cont}")
                    else:
                        all_lines.append(f"{prefix}")
                all_lines.append("")

            elif el.name == "h1":
                text = el.get_text(strip=True)
                if text:
                    all_lines.append("")
                    all_lines.append(text.upper())
                    all_lines.append("═" * min(len(text), 80))
                    all_lines.append("")

            elif el.name in ("h2", "h3"):
                text = el.get_text(strip=True)
                if text:
                    all_lines.append("")
                    all_lines.append(text)
                    all_lines.append("─" * min(len(text), 80))
                    all_lines.append("")

            elif el.name in ("h4", "h5", "h6"):
                text = el.get_text(strip=True)
                if text:
                    all_lines.append(f"▸ {text}")
                    all_lines.append("")

            elif el.name == "pre":
                code_lines = el.get_text().splitlines()
                all_lines.append("┌─ code " + "─" * 33 + "┐")
                for cl in code_lines:
                    all_lines.append(f"│ {cl}")
                all_lines.append("└" + "─" * 40 + "┘")
                all_lines.append("")

            elif el.name == "blockquote":
                for line in el.get_text().splitlines():
                    if line.strip():
                        for w in textwrap.wrap(line.strip(), width=82):
                            all_lines.append(f"  ▏ {w}")
                all_lines.append("")

            elif el.name == "hr":
                all_lines.append("─" * 80)
                all_lines.append("")

            else:
                text = el.get_text(separator=" ", strip=True)
                if text:
                    for line in textwrap.wrap(text, width=90):
                        all_lines.append(line)
                    all_lines.append("")

        lines_per_page = 40
        pages = []
        for i in range(0, max(1, len(all_lines)), lines_per_page):
            pages.append(all_lines[i:i + lines_per_page])
        return pages

    def _table_to_lines(self, table_el) -> list:
        """Render an HTML table element as ASCII grid lines.

        Args:
            table_el: BeautifulSoup table element.

        Returns:
            A list of strings representing the table rows.
        """
        rows = []
        for tr in table_el.find_all("tr"):
            cells = [td.get_text(separator=" ", strip=True) for td in tr.find_all(["td", "th"])]
            rows.append(cells)

        if not rows:
            return []

        col_count = max(len(r) for r in rows)
        rows = [r + [""] * (col_count - len(r)) for r in rows]

        col_widths = [0] * col_count
        for row in rows:
            for i, cell in enumerate(row):
                col_widths[i] = max(col_widths[i], min(len(cell), 30))

        def make_row(cells):
            parts = []
            for i, cell in enumerate(cells):
                truncated = cell[:col_widths[i]]
                parts.append(truncated.ljust(col_widths[i]))
            return "│ " + " │ ".join(parts) + " │"

        sep = "├─" + "─┼─".join("─" * w for w in col_widths) + "─┤"
        top = "┌─" + "─┬─".join("─" * w for w in col_widths) + "─┐"
        bot = "└─" + "─┴─".join("─" * w for w in col_widths) + "─┘"

        lines = [top]
        for i, row in enumerate(rows):
            lines.append(make_row(row))
            if i < len(rows) - 1:
                lines.append(sep)
        lines.append(bot)
        return lines

    def _split_markdown_pages(self, text: str):
        """Split markdown text into pages based on headings and content length.

        Args:
            text: Raw markdown string.

        Returns:
            A list of plain-text page strings.
        """
        html = markdown(text)
        soup = BeautifulSoup(html, "html.parser")

        pages = []
        buffer = ""

        for element in soup.descendants:
            if element.name in ("h1", "h2", "h3") or (isinstance(element, str) and len(buffer) > 1500):
                if buffer.strip():
                    pages.append(buffer.strip())
                    buffer = ""
                if element.name in ("h1", "h2", "h3"):
                    buffer += f"# {element.text}\n\n"
            elif isinstance(element, str):
                buffer += element.strip() + " "

        if buffer.strip():
            pages.append(buffer.strip())

        return pages or ["(empty)"]

    def on_mount(self) -> None:
        """Start async document loading when the widget is mounted."""
        self.can_focus = True
        self._load_document(self.path)

    def on_mount(self) -> None:
        """Start async document loading when the widget is mounted."""
        self.can_focus = True
        self._load_and_render(self.path)

    @work(thread=True, exclusive=True, name="load-document")
    def _load_and_render(self, path) -> None:
        """Worker: load the document and pre-render all pages in a background thread.

        Renders page 0 first and pushes it to the UI immediately, then continues
        rendering the remaining pages into pages_cache in the background.
        """
        try:
            self._check_file(path)
        except Exception as exc:
            self.app.call_from_thread(self._show_error, str(exc))
            return

        # Pre-render all pages while still in the worker thread (thread-safe)
        self.pages_cache = []
        for i in range(self.total_pages):
            try:
                img = self._render_page_pil(i)
            except Exception:
                img = PILImage.new("RGB", (800, 600), "white")
            self.pages_cache.append(img)
            # Push page 0 to the UI as soon as it's ready
            if i == 0:
                self.app.call_from_thread(self._update_image, img)

    def _show_error(self, message: str) -> None:
        """Replace the image placeholder with an error label."""
        try:
            self.query_one("#pdf-image").remove()
        except Exception:
            pass
        self.mount(Label(f"[red]{message}[/red]", id="pdf-error"))

    def compose(self) -> ComposeResult:
        """Compose the widget with a loading placeholder."""  # noqa: DOC402
        yield Label("Loading document...", id="pdf-loading")
        yield timg.__dict__[
            self.protocol + "Image" if self.protocol not in ("Auto", "") else "Image"
        ](PILImage.new("RGB", (800, 600), "white"), id="pdf-image")

    def on_mount_ready(self) -> None:
        """Hide image until document is loaded."""
        self.query_one("#pdf-image").display = False

    def _render_page_pil(self, page_index: int) -> PILImage.Image:
        """Renders the given page index and returns a PIL image.

        Args:
            page_index: Zero-based index of the page to render.

        Returns:
            PIL.Image: a valid PIL image.

        Raises:
            PDFRuntimeError: when a document is not open.
            PDFHasAPasswordError: when the document has a password.
        """
        if self.doc_type == "pdf":
            try:
                page = self.doc.load_page(page_index)
            except ValueError:
                raise PDFHasAPasswordError(
                    f"{self.path} is a document that is encrypted, and cannot be read."
                )
            pix = page.get_pixmap()
            mode = "RGBA" if pix.alpha else "RGB"
            return PILImage.frombytes(mode, (pix.width, pix.height), pix.samples)

        elif self.doc_type == "text_pages":
            lines = self.doc[page_index] if page_index < len(self.doc) else []
            return self._draw_text_page(lines)

        else:
            raise PDFRuntimeError(f"Unknown document type: {self.doc_type}")

    def _draw_text_page(self, lines) -> PILImage.Image:
        """Render a list of text lines into a PIL image.

        Args:
            lines: List of strings to render.

        Returns:
            PIL.Image: rendered page image.
        """
        width = 900

        try:
            if getattr(self, "font_path", None):
                font = ImageFont.truetype(self.font_path, self.font_size)
            else:
                font = ImageFont.truetype("arial.ttf", self.font_size)
        except Exception:
            try:
                font = ImageFont.truetype("DejaVuSans.ttf", self.font_size)
            except Exception:
                font = ImageFont.load_default()

        def wrap_text_by_pixel(draw_obj, text, font_obj, max_width):
            if not text:
                return [""]
            words = text.split()
            lines_out = []
            line = words[0]
            for w in words[1:]:
                test = line + " " + w
                try:
                    bbox = draw_obj.textbbox((0, 0), test, font=font_obj)
                    test_width = bbox[2] - bbox[0]
                except Exception:
                    test_width = len(test) * (self.font_size // 2)
                if test_width <= max_width:
                    line = test
                else:
                    lines_out.append(line)
                    line = w
            lines_out.append(line)
            return lines_out

        temp_img = PILImage.new("RGB", (width, 2000), "white")
        temp_draw = ImageDraw.Draw(temp_img)
        max_text_width = width - 20

        est_lines = 0
        for line in lines:
            wrapped = wrap_text_by_pixel(temp_draw, line, font, max_text_width)
            est_lines += max(1, len(wrapped))

        try:
            ascent, descent = font.getmetrics()
            line_height = ascent + descent + 4
        except Exception:
            line_height = self.font_size + 6

        height = max(300, est_lines * line_height + 40)

        image = PILImage.new("RGB", (width, height), "white")
        draw = ImageDraw.Draw(image)

        y = 10
        for line in lines:
            wrapped = wrap_text_by_pixel(draw, line, font, max_text_width)
            for wrapped_line in wrapped:
                draw.text((10, y), wrapped_line, fill="black", font=font)
                y += line_height
            if wrapped == [""]:
                y += 6

        final_height = max(300, y + 10)
        if final_height != height:
            image = image.crop((0, 0, width, final_height))
        return image

    def render_page(self) -> None:
        """Show the current page from cache (or trigger a reload if cache is empty)."""
        if not self.doc:
            raise PDFRuntimeError("`render_page` was called before a document was opened.")
        if self.pages_cache and self.current_page < len(self.pages_cache):
            self._update_image(self.pages_cache[self.current_page])
        else:
            # Cache not ready yet — re-trigger the full load
            self._load_and_render(self.path)

    def _update_image(self, image: PILImage.Image) -> None:
        """Update the image widget on the main thread and hide the loading label."""
        try:
            loading = self.query_one("#pdf-loading")
            loading.display = False
        except Exception:
            pass
        try:
            image_widget = self.query_one("#pdf-image")
            image_widget.display = True
            image_widget.image = image
        except Exception:
            pass

    def watch_current_page(self, new_page: int) -> None:
        """Change the current page to a different page based on the value provided.

        Args:
            new_page(int): The page to switch to.
        """
        if self.is_mounted and self.doc:
            if self.pages_cache and new_page < len(self.pages_cache):
                self._update_image(self.pages_cache[new_page])

    def watch_protocol(self, protocol: str) -> None:
        """Change the rendering protocol.

        Args:
            protocol(str): The protocol to use.

        Raises:
            AssertionError: When the protocol is not a valid option.
        """
        assert protocol in ["Auto", "TGP", "Sixel", "Halfcell", "Unicode", ""]
        if self.is_mounted:
            self.refresh(recompose=True)
            self.render_page()

    def watch_path(self, path: str | Path) -> None:
        """Reload the document when the path changes.

        Args:
            path(str|Path): The path to the document.
        """
        if not self.is_mounted:
            return
        self.doc = None
        self.pages_cache = []
        self.current_page = 0
        try:
            self.query_one("#pdf-loading").display = True
        except Exception:
            pass
        self._load_and_render(path)

    def load(self, path: str | Path | io.BytesIO, reset_page: bool = True) -> None:
        """Load a new document into the viewer asynchronously.

        Args:
            path: Path to the document or a BytesIO stream.
            reset_page(bool): Whether to reset to page 0 after loading.
        """
        self.path = path
        if reset_page:
            self.current_page = 0
        self.doc = None
        self.pages_cache = []
        try:
            self.query_one("#pdf-loading").display = True
        except Exception:
            pass
        self._load_and_render(path)

    def on_key(self, event: events.Key) -> None:
        """Handle key presses.

        Args:
            event(events.Key): The key event.
        """
        if not self.use_keys:
            return
        match event.key:
            case "down" | "page_down" | "right":
                event.stop()
                self.next_page()
            case "up" | "page_up" | "left":
                event.stop()
                self.previous_page()
            case "home":
                event.stop()
                self.go_to_start()
            case "end":
                event.stop()
                self.go_to_end()

    def next_page(self) -> None:
        """Go to the next page."""
        if self.doc and self.current_page < self.total_pages - 1:
            self.current_page += 1

    def previous_page(self) -> None:
        """Go to the previous page."""
        if self.doc and self.current_page > 0:
            self.current_page -= 1

    def go_to_start(self) -> None:
        """Go to the first page."""
        if self.doc:
            self.current_page = 0

    def go_to_end(self) -> None:
        """Go to the last page."""
        if self.doc:
            self.current_page = self.total_pages - 1
