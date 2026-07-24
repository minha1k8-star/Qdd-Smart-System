import zipfile, re, io, shutil, sys
import openpyxl

def load_patched(path, **kwargs):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for n in names:
                data = z.read(n)
                if n == 'xl/styles.xml':
                    text = data.decode('utf-8', errors='replace')
                    text = re.sub(r'(<family val=")(\d+)(")', lambda m: m.group(1) + (m.group(2) if int(m.group(2)) <= 14 else '0') + m.group(3), text)
                    data = text.encode('utf-8')
                zout.writestr(n, data)
        buf.seek(0)
        return openpyxl.load_workbook(buf, **kwargs)

if __name__ == '__main__':
    wb = load_patched(sys.argv[1], data_only=True)
    print(wb.sheetnames)
