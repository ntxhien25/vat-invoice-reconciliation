"""Public demo gate: reject lookup links and workstation paths without logging values."""
from pathlib import Path
import re
from openpyxl import load_workbook
PATTERN=re.compile(r'https?://|www\.|meinvoice\.vn|[A-Z]:[\\/]|/Users/|/home/',re.I)
def unsafe_text(value):
    return isinstance(value,str) and bool(PATTERN.search(value))
def check_workbook(path):
    book=load_workbook(path,read_only=False,data_only=False)
    try:
        for sheet in book:
            for row in sheet:
                for cell in row:
                    if unsafe_text(cell.value) or (cell.hyperlink and unsafe_text(cell.hyperlink.target)):
                        raise ValueError(f'Public-demo privacy check failed: {Path(path).name}, {sheet.title}, {cell.coordinate}; value withheld.')
        if book._external_links:
            raise ValueError('Public-demo workbook contains external links; details withheld.')
    finally:book.close()
def check_release(root):
    paths=list(Path(root).glob('*.xlsx'))+list((Path(root)/'reports').rglob('*.xlsx'))
    for path in paths:check_workbook(path)
    return len(paths)
if __name__=='__main__':
    count=check_release(Path(__file__).resolve().parent)
    print(f'Privacy checks passed for {count} fictional workbooks.')
