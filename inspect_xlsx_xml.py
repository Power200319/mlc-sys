import sys
import zipfile
import xml.etree.ElementTree as ET

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
path = r"D:\New folder\Telegram Desktop\E-Gradebook_6A_English Q4.xlsx"
ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
with zipfile.ZipFile(path) as z:
    shared=[]
    if "xl/sharedStrings.xml" in z.namelist():
        ss=ET.fromstring(z.read("xl/sharedStrings.xml"))
        shared=["".join(t.text or "" for t in item.findall('.//m:t',ns)) for item in ss.findall('m:si',ns)]
    wb=ET.fromstring(z.read("xl/workbook.xml")); rels=ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    relmap={x.attrib['Id']:x.attrib['Target'] for x in rels}
    for s in wb.findall('m:sheets/m:sheet',ns):
        title=s.attrib['name']; rid=s.attrib['{'+ns['r']+'}id']; target='xl/'+relmap[rid].lstrip('/')
        root=ET.fromstring(z.read(target))
        print('\n---',title)
        for row in root.findall('m:sheetData/m:row',ns):
            n=int(row.attrib['r'])
            if n>18: break
            out=[]
            for c in row.findall('m:c',ns):
                ref=c.attrib.get('r'); typ=c.attrib.get('t'); v=c.find('m:v',ns); f=c.find('m:f',ns)
                value=v.text if v is not None else ''
                if typ=='s' and value: value=shared[int(value)]
                if f is not None: value='='+str(f.text or '')
                if value: out.append(f'{ref}={value[:80]}')
            if out: print(' | '.join(out))
