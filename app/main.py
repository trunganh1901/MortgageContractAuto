import xlwings as xw
from config.templates import load_cfg_templates
from app.workflow import run

def main():
    wb = xw.Book.caller()
    cfg = load_cfg_templates()

    code = wb.sheets["UI_DASHBOARD"].range("B2").value
    folder = run(wb, code, cfg)

    wb.sheets["UI_DASHBOARD"].range("B7").value = folder

if __name__ == "__main__":
    xw.Book().set_mock_caller()
    main()
