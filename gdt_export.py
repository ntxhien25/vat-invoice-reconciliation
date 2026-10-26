"""GDT invoice exports. Human CAPTCHA, automatic continuation, validated XLSX files.

Run --help or --plan for usage. Authentication is never read from another browser.
"""
from __future__ import annotations

import argparse
import getpass
import hashlib
import json
import re
import sys
import uuid
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

from gdt_support import (ExportError, InvoiceTable, Period, SAIGON, clean, export_filename,
                         load_session, previous_month, read_invoice_table,
                         save_session, session_path, validate_tax_id,
                         windows_keyring, write_sales_workbook)

HOME = "https://hoadondientu.gdt.gov.vn/"
LOOKUP = HOME + "tra-cuu/tra-cuu-hoa-don"
ELECTRONIC = "Hóa đơn điện tử"
CASH = "Hóa đơn có mã khởi tạo từ máy tính tiền"
CASH_RESULT = "Cục Thuế đã nhận hóa đơn có mã khởi tạo từ máy tính tiền"
KEYRING_SERVICE = "GDTInvoiceExporter"


@dataclass(frozen=True)
class Job:
    name: str
    direction: str
    table: str
    result: str


JOBS = (
    Job("purchases_coded", "purchases", ELECTRONIC, "Đã cấp mã hóa đơn"),
    Job("purchases_no_code", "purchases", ELECTRONIC, "Cục Thuế đã nhận không mã"),
    Job("purchases_cash_register", "purchases", CASH, CASH_RESULT),
    Job("sales_electronic", "sales", ELECTRONIC, "Tất cả"),
    Job("sales_cash_register", "sales", CASH, "Tất cả"),
)


def selected_jobs(mode: str) -> list[Job]:
    return [j for j in JOBS if mode == "all" or j.direction == mode]


def write_report(path: Path, report: dict) -> None:
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(path)


def saved_password_or_manual(tax_id: str) -> str | None:
    """Optional autofill must not prevent a normal manual login."""
    try:
        password = windows_keyring().get_password(KEYRING_SERVICE, tax_id)
    except Exception:
        print("Windows password autofill is unavailable. Enter your password in the browser instead.")
        return None
    if not password:
        print("No password is saved for this MST. Enter your password in the browser instead.")
    return password or None


def failure_message(exc: BaseException, stage: str, password: str | None = None) -> str:
    """Show the useful first error line, without credentials or URL queries."""
    detail = str(exc).splitlines()[0] if str(exc) else type(exc).__name__
    if password:
        detail = detail.replace(password, "[password redacted]")
    detail = re.sub(r"(https?://[^\s?]+)\?[^\s]+", r"\1?[query redacted]", detail)
    return f"{stage}: {detail[:600]}"


def ask_validated(prompt, validator, default=None):
    while True:
        value = input(prompt).strip()
        if not value and default is not None:
            value = default
        try:
            return validator(value)
        except (ValueError, ExportError) as exc:
            print(f"Invalid input: {exc} Please try again.")


def mode_value(value):
    value = value.strip().casefold()
    if value not in ("all", "purchases", "sales"):
        raise ExportError("Choose ALL, PURCHASES or SALES.")
    return value


def yes_no_value(value):
    value = value.strip().casefold()
    if value not in ("y", "n"):
        raise ExportError("Enter Y or N.")
    return value == "y"


def retry_transient(operation, attempts=3, pause=None, on_retry=None):
    """Retry only temporary gateway failures/timeouts, never validation or access errors."""
    for attempt in range(1, attempts + 1):
        try:
            return operation()
        except Exception as exc:
            temporary = (type(exc).__name__ == "TimeoutError"
                         or isinstance(exc, ExportError) and bool(re.search(r"HTTP (502|503|504)\b", str(exc))))
            if not temporary or attempt == attempts:
                raise
            if on_retry:
                on_retry(exc, attempt)
            if pause:
                pause(attempt * 2000)


class Portal:
    """Selectors verified against the GDT interface on 14 October 2026.

    Forms are scoped by counterparty input ID, not DOM order. Ant Design v3
    dropdowns are scoped by aria-controls. Dates use the editable calendar popup.
    """

    def __init__(self, page, timeout: int = 60_000):
        from playwright.sync_api import expect
        self.page = page
        self.expect = expect
        self.timeout = timeout

    def dismiss_notice(self):
        self.page.keyboard.press("Escape")
        for button in self.page.get_by_role("button", name="Close", exact=True).all():
            if button.is_visible():
                button.click()

    def is_authenticated(self, timeout: int) -> bool:
        from playwright.sync_api import TimeoutError as PlaywrightTimeout
        try:
            self.page.get_by_text("Quản lý hóa đơn phát sinh", exact=True).wait_for(
                state="visible", timeout=timeout)
            return True
        except PlaywrightTimeout:
            return False

    def login(self, tax_id: str, password: str | None, login_timeout: int):
        self.dismiss_notice()
        # The page also contains a hidden dialog title with identical text.
        self.page.locator("header").get_by_text("Đăng nhập", exact=True).click()
        modal = self.page.locator(".ant-modal:visible").filter(
            has=self.page.locator("#username"))
        modal.locator("#username").fill(tax_id)
        if password:
            modal.locator("#password").fill(password)
        print("In the browser: enter your password if needed, enter CAPTCHA, then click Đăng nhập.")
        print("The exporter continues automatically after login. No F8 or Inspector is needed.")
        if not self.is_authenticated(login_timeout):
            raise ExportError("Login was not completed within the allowed time.")
        self.expect(modal).to_be_hidden(timeout=login_timeout)

    def open_lookup(self):
        """Use authenticated SPA menus; a full URL reload can lose app state."""
        purchase_tab = self.page.get_by_role("tab", name="Tra cứu hóa đơn điện tử mua vào", exact=True)
        if purchase_tab.is_visible():
            return
        self.page.get_by_text("Tra cứu", exact=True).filter(visible=True).click()
        self.page.get_by_text("Tra cứu hóa đơn", exact=True).filter(visible=True).click()
        outcome = self.page.wait_for_function("""() => {
            const text = document.body.innerText;
            if (/403/.test(text) && /không có quyền truy cập/i.test(text)) return 'denied';
            const tabs = [...document.querySelectorAll('[role="tab"]')];
            if (tabs.some(tab => tab.textContent.trim() === 'Tra cứu hóa đơn điện tử mua vào'
                && tab.getClientRects().length > 0)) return 'ready';
            return false;
        }""", timeout=self.timeout)
        if outcome.json_value() == "denied":
            raise ExportError("GDT denied invoice lookup (403) after menu navigation. Check this account's access in the normal portal; no export started.")
        self.expect(purchase_tab).to_be_visible(timeout=self.timeout)

    def form_and_root(self, direction: str):
        tab_name = "Tra cứu hóa đơn điện tử mua vào" if direction == "purchases" else "Tra cứu hóa đơn điện tử bán ra"
        self.page.get_by_role("tab", name=tab_name, exact=True).click()
        marker = "#nbmst" if direction == "purchases" else "#nmmst"
        form = self.page.locator("form").filter(has=self.page.locator(marker))
        self.expect(form).to_have_count(1)
        root = form.locator("xpath=ancestor::div[@role='tabpanel'][1]")
        self.expect(root).to_have_attribute("aria-hidden", "false")
        return form, root

    def select(self, form, field: str, value: str):
        combo = form.locator(field + ' [role="combobox"]')
        current = form.locator(field + " .ant-select-selection-selected-value")
        if current.count() and clean(current.inner_text()) == value:
            return
        combo.click()
        menu_id = combo.get_attribute("aria-controls")
        if not menu_id:
            raise ExportError(f"The dropdown {field} no longer exposes its option list.")
        menu = self.page.locator(f'[id="{menu_id}"]')
        self.expect(menu).to_be_visible()
        option = menu.locator("li").filter(has_text=re.compile(r"^\s*" + re.escape(value) + r"\s*$"))
        if option.count() != 1:
            raise ExportError(f"The portal does not offer the required option: {value}.")
        option.click()
        self.expect(current).to_have_text(value)

    def set_date(self, form, field: str, value: str):
        display = form.locator(field + " input")
        display.click()
        editor = self.page.locator(".ant-calendar-input:visible")
        self.expect(editor).to_have_count(1)
        editor.fill(value)
        editor.press("Enter")
        self.expect(display).to_have_value(value)
        self.page.keyboard.press("Escape")

    def prepare(self, job: Job, period: Period):
        form, root = self.form_and_root(job.direction)
        root.get_by_role("tab", name=job.table, exact=True).click()
        self.expect(root.get_by_role("tab", name=job.table, exact=True)).to_have_attribute("aria-selected", "true")
        for field in ("#nbmst" if job.direction == "purchases" else "#nmmst",
                      "#nmcccd", "#shdon", "#khhdon"):
            form.locator(field).fill("")
        model = form.locator("#khmshdon")
        if model.locator(".ant-select-selection-selected-value").count():
            clear_button = model.locator(".ant-select-clear")
            if not clear_button.count():
                raise ExportError("Clear the invoice-form filter before exporting all invoices.")
            clear_button.click()
        form.locator('input[type="checkbox"]').uncheck()
        self.select(form, "#tthai", "Tất cả")
        self.select(form, "#ttxly", job.result)
        self.set_date(form, "#tngay", period.start.strftime("%d/%m/%Y"))
        self.set_date(form, "#dngay", period.end.strftime("%d/%m/%Y"))
        self.verify(form, root, job, period)
        return form, root

    def verify(self, form, root, job: Job, period: Period):
        self.expect(form.locator("#tthai .ant-select-selection-selected-value")).to_have_text("Tất cả")
        self.expect(form.locator("#ttxly .ant-select-selection-selected-value")).to_have_text(job.result)
        self.expect(form.locator("#tngay input")).to_have_value(period.start.strftime("%d/%m/%Y"))
        self.expect(form.locator("#dngay input")).to_have_value(period.end.strftime("%d/%m/%Y"))
        self.expect(root.get_by_role("tab", name=job.table, exact=True)).to_have_attribute("aria-selected", "true")

    @staticmethod
    def invoice_response(response, direction: str) -> bool:
        """Observe the normal UI search response; no direct/private API requests.

        Match fresh search direction and invoice-list payload, not any unrelated
        XHR. Fail with a timeout if the portal changes its response contract.
        """
        parsed = urlparse(response.url)
        host = parsed.hostname or ""
        path = parsed.path.lower()
        matches_direction = "purchase" in path if direction == "purchases" else bool(re.search(r"sold|sales?", path))
        if (not (host == "gdt.gov.vn" or host.endswith(".gdt.gov.vn"))
                or "invoice" not in path or not matches_direction
                or response.request.resource_type not in ("xhr", "fetch")):
            return False
        if response.status >= 400:
            return True
        try:
            body = response.json()
            return isinstance(body, dict) and isinstance(body.get("datas"), list) and isinstance(body.get("total"), int)
        except Exception:
            return False

    def search(self, form, root, job: Job, period: Period) -> int:
        with self.page.expect_response(lambda r: self.invoice_response(r, job.direction), timeout=self.timeout) as pending:
            form.get_by_role("button", name="Tìm kiếm", exact=True).click()
        response = pending.value
        if response.status in (401, 403):
            raise ExportError("Session expired or access was denied. Sign in again and rerun.")
        if response.status >= 400:
            raise ExportError(f"GDT search returned HTTP {response.status}; no export confirmed.")
        expected = response.json()["total"]
        if expected < 0:
            raise ExportError("GDT returned an invalid invoice count.")
        count_pattern = re.compile(r"^Có\s+[\d.,]+\s+kết quả$")
        counter = root.get_by_text(count_pattern)
        self.expect(counter).to_have_text(re.compile(r"^Có\s+" + f"{expected:,}".replace(",", r"[.,]?") + r"\s+kết quả$"), timeout=self.timeout)
        self.expect(root.locator(".ant-spin-spinning:visible")).to_have_count(0, timeout=self.timeout)
        self.verify(form, root, job, period)
        return expected

    def export(self, root, destination: Path):
        button = root.locator("button:has(svg g#icon_ketxuat)")
        self.expect(button).to_have_count(1)
        with self.page.expect_download(timeout=self.timeout) as pending:
            button.click()
        download = pending.value
        if Path(download.suggested_filename).suffix.lower() != ".xlsx":
            raise ExportError("GDT returned an unexpected export format; expected XLSX.")
        if download.failure():
            raise ExportError("The browser download failed.")
        download.save_as(destination)


def empty_table() -> InvoiceTable:
    return InvoiceTable(["STT", "Ký hiệu mẫu số", "Ký hiệu hóa đơn", "Số hóa đơn", "Ngày lập",
                         "MST người bán/MST người xuất hàng", "MST người mua/MST người nhận hàng"], [])


def create_context(browser, remembered: dict | None):
    options = {"accept_downloads": True, "locale": "vi-VN", "viewport": {"width": 1600, "height": 1000}}
    if remembered:
        options["storage_state"] = remembered["browser"]
    context = browser.new_context(**options)
    if remembered and remembered.get("session_storage"):
        # Domain restricted and restore only into a new/empty tab session.
        encoded = json.dumps(remembered["session_storage"], ensure_ascii=True)
        context.add_init_script("""(() => {
            if (location.origin === 'https://hoadondientu.gdt.gov.vn' && sessionStorage.length === 0) {
                const values = """ + encoded + """;
                for (const [key, value] of Object.entries(values)) sessionStorage.setItem(key, value);
            }
        })();""")
    return context


def run(args) -> int:
    from playwright.sync_api import sync_playwright
    period = Period.month(args.month)
    jobs = selected_jobs(args.mode)
    password = None
    if args.use_credentials:
        password = saved_password_or_manual(args.tax_id)
    cache = session_path(args.tax_id) if args.remember_session else None
    run_id = datetime.now(SAIGON).strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:6]
    folder = args.output.resolve() / args.tax_id / period.label / run_id
    raw = folder / "raw"
    raw.mkdir(parents=True)
    report = {"run_id": run_id, "tax_id": args.tax_id, "month": period.label,
              "started_at": datetime.now(SAIGON).isoformat(), "status": "running",
              "sales_scope": "Both result tabs, all invoice statuses and processing results",
              "jobs": [], "outputs": []}
    report_path = folder / "run_report.json"
    write_report(report_path, report)
    print(f"Output folder: {folder}")
    remembered = None
    if cache and cache.exists():
        try:
            remembered = load_session(cache)
            if remembered.get("tax_id") != args.tax_id:
                remembered = None
        except Exception:
            print("Saved session cannot be read. A fresh login will be used.")
    validated = {}
    context = browser = None
    stage = "Starting the browser"
    try:
        with sync_playwright() as playwright:
            launch = {"headless": False}
            if args.browser != "chromium":
                launch["channel"] = "msedge" if args.browser == "edge" else "chrome"
            browser = playwright.chromium.launch(**launch)
            stage = "Creating the browser session"
            context = create_context(browser, remembered)
            page = context.new_page()
            page.set_default_timeout(args.timeout * 1000)
            stage = "Opening the GDT homepage"
            page.goto(HOME, wait_until="domcontentloaded")
            portal = Portal(page, args.timeout * 1000)
            stage = "Dismissing the GDT notice"
            portal.dismiss_notice()
            stage = "Checking the saved login session"
            if not remembered or not portal.is_authenticated(8000):
                if remembered:
                    print("Saved session expired. Opening a fresh login.")
                    context.close()
                    context = create_context(browser, None)
                    page = context.new_page()
                    page.set_default_timeout(args.timeout * 1000)
                    page.goto(HOME, wait_until="domcontentloaded")
                    portal = Portal(page, args.timeout * 1000)
                stage = "Opening and completing GDT login"
                portal.login(args.tax_id, password, args.login_timeout * 1000)
            password = None
            stage = "Opening invoice lookup"
            portal.open_lookup()
            for job in jobs:
                print(f"Exporting {job.name} ...")
                entry = {"name": job.name, "direction": job.direction,
                         "table": job.table, "processing_result": job.result,
                         "invoice_status": "Tất cả", "status": "running"}
                report["jobs"].append(entry)
                write_report(report_path, report)
                def perform_job():
                    nonlocal stage
                    entry["attempts"] = entry.get("attempts", 0) + 1
                    stage = "Setting filters for " + job.name
                    form, root = portal.prepare(job, period)
                    stage = "Searching " + job.name
                    expected = portal.search(form, root, job, period)
                    entry["portal_count"] = expected
                    if expected == 0:
                        table = empty_table()
                        entry.update(status="empty", validated_count=0,
                                     note="No invoices found; no portal download requested.")
                        if job.direction == "purchases":
                            target = folder / export_filename(job.name, period.label)
                            write_sales_workbook(target, [("Invoices", table)],
                                                 {"Source": "Empty search result, not a downloaded GDT file", "MST": args.tax_id, "Month": period.label})
                            report["outputs"].append(target.name)
                    else:
                        target = raw / export_filename(job.name, period.label)
                        stage = "Downloading " + job.name
                        portal.export(root, target)
                        stage = "Validating " + job.name
                        table = read_invoice_table(target, expected, period, args.tax_id,
                                                   job.direction, job.result)
                        entry.update(status="validated", validated_count=len(table.rows),
                                     duplicate_keys=table.duplicate_keys,
                                     sha256=hashlib.sha256(target.read_bytes()).hexdigest(),
                                     raw_file=str(target.relative_to(folder)))
                        if job.direction == "purchases":
                            # Publish only after validation; invalid downloads stay under raw/.
                            published = folder / target.name
                            published.write_bytes(target.read_bytes())
                            report["outputs"].append(published.name)
                    validated[job.name] = table
                    print(f"  {entry['status']}: {expected} invoices")
                def report_retry(exc, attempt):
                    entry.setdefault("retry_errors", []).append(failure_message(exc, stage))
                    print(f"  Temporary failure; retrying {job.name} (attempt {attempt + 1}/{args.attempts}).")
                    write_report(report_path, report)
                try:
                    retry_transient(perform_job, args.attempts, page.wait_for_timeout, report_retry)
                except Exception as exc:
                    entry.update(status="failed", error=failure_message(exc, stage))
                    print(f"  FAILED: {entry['error']}")
                    if args.debug:
                        try:
                            page.screenshot(path=str(raw / f"{job.name}_error.png"), full_page=True)
                        except Exception:
                            pass
                write_report(report_path, report)
            if args.mode in ("all", "sales") and all(j in validated for j in ("sales_electronic", "sales_cash_register")):
                stage = "Assembling the sales workbook"
                target = folder / export_filename("sales_all", period.label)
                write_sales_workbook(target, [("Electronic", validated["sales_electronic"]),
                                              ("Cash register", validated["sales_cash_register"])],
                                     {"MST": args.tax_id, "Month": period.label, "Run ID": run_id,
                                      "Scope": report["sales_scope"],
                                      "Electronic invoices": len(validated["sales_electronic"].rows),
                                      "Cash register invoices": len(validated["sales_cash_register"].rows)})
                report["outputs"].append(target.name)
            if cache and portal.is_authenticated(1000):
                try:
                    save_session(cache, {"tax_id": args.tax_id, "browser": context.storage_state(indexed_db=True),
                                         "session_storage": page.evaluate("() => Object.fromEntries(Object.entries(sessionStorage))")})
                    print("Encrypted session saved for this Windows user.")
                except Exception:
                    print("Session could not be saved; exports are unaffected.")
            context.close()
            browser.close()
            context = browser = None
    except BaseException as exc:
        report["error"] = failure_message(exc, stage, password)
        report["failed_step"] = stage
        report["status"] = "interrupted" if isinstance(exc, KeyboardInterrupt) else "failed"
        report["finished_at"] = datetime.now(SAIGON).isoformat()
        write_report(report_path, report)
        if isinstance(exc, KeyboardInterrupt):
            raise
        raise ExportError(report["error"]) from exc
    report["status"] = "complete" if len(validated) == len(jobs) else "partial_failure"
    report["finished_at"] = datetime.now(SAIGON).isoformat()
    write_report(report_path, report)
    print(f"Run status: {report['status']}. Report: {report_path}")
    return 0 if report["status"] == "complete" else 2


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description="Export three purchase categories and all sales invoice types from GDT.")
    result.add_argument("--mode", type=str.lower, choices=("all", "purchases", "sales"), default="all")
    result.add_argument("--month", help="Calendar month YYYY-MM; previous month by default.")
    result.add_argument("--tax-id", help="Requested taxpayer MST; checked against downloaded invoices.")
    result.add_argument("--output", type=Path, default=Path(__file__).resolve().parent / "exports")
    result.add_argument("--browser", choices=("chromium", "edge", "chrome"), default="edge")
    result.add_argument("--remember-session", action="store_true", help="Save/reuse DPAPI-encrypted session in LOCALAPPDATA.")
    result.add_argument("--use-credentials", action="store_true", help="Autofill saved Windows Credential Manager password.")
    result.add_argument("--setup-credentials", action="store_true", help="Securely prompt and save a password, then exit.")
    result.add_argument("--plan", action="store_true", help="Print jobs without a browser, credentials or downloads.")
    result.add_argument("--interactive", action="store_true", help="Ask for month, MST, mode and session preferences.")
    result.add_argument("--timeout", type=int, default=90, help="Search/download timeout in seconds.")
    result.add_argument("--login-timeout", type=int, default=600, help="Time for manual login in seconds.")
    result.add_argument("--attempts", type=int, default=3, help="Maximum attempts per job for temporary HTTP gateway failures/timeouts.")
    result.add_argument("--debug", action="store_true", help="Save invoice-screen error screenshots locally (may contain financial data).")
    return result


def main(argv=None) -> int:
    args = parser().parse_args(argv)
    try:
        if args.interactive:
            args.month = ask_validated("Invoice month [YYYY-MM] (example 2026-09): ", lambda value: Period.month(value).label)
            args.tax_id = ask_validated("Taxpayer MST: ", validate_tax_id)
            args.mode = ask_validated("Export ALL / PURCHASES / SALES [ALL]: ", mode_value, "ALL")
            args.remember_session = ask_validated("Remember encrypted login session on this PC? [Y/N] (Enter = N): ", yes_no_value, "N")
            args.use_credentials = ask_validated("Try password autofill from Windows Credential Manager? [Y/N] (Enter = N): ", yes_no_value, "N")
        args.month = args.month or previous_month()
        period = Period.month(args.month)
        if args.timeout <= 0 or args.login_timeout <= 0:
            raise ExportError("Timeouts must be greater than zero.")
        if not 1 <= args.attempts <= 5:
            raise ExportError("Attempts must be between 1 and 5.")
        if args.plan:
            print(f"Month: {period.label} ({period.start:%d/%m/%Y} - {period.end:%d/%m/%Y})")
            for job in selected_jobs(args.mode):
                print(f"{job.name}: {job.table}; result={job.result}; status=Tất cả")
            return 0
        if not args.tax_id:
            raise ExportError("Supply --tax-id, or use --interactive.")
        args.tax_id = validate_tax_id(args.tax_id)
        if args.setup_credentials:
            secret = getpass.getpass("GDT password (saved only in Windows Credential Manager): ")
            if not secret:
                raise ExportError("Password cannot be empty.")
            windows_keyring().set_password(KEYRING_SERVICE, args.tax_id, secret)
            print("Password saved to Windows Credential Manager.")
            return 0
        return run(args)
    except KeyboardInterrupt:
        print("Run interrupted. Check the run report for any completed files.")
        return 130
    except Exception as exc:
        message = str(exc) if isinstance(exc, ExportError) else f"{type(exc).__name__}: operation failed. Check setup and the local run report."
        print(f"ERROR: {message}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    raise SystemExit(main())
