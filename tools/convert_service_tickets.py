#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convert_service_tickets.py
Chuyển file RAW "Service Tickets - Yêu cầu đặt hàng" (xuất từ Base) thành file
chỉ có 1 sheet "Data_sạch" (28+2 cột, đúng cấu trúc bản đã duyệt) để nạp cho Claude phân tích.

Cách dùng:
    pip install pandas openpyxl
    python convert_service_tickets.py RAW.xlsx                 # ra RAW_CLEAN.xlsx
    python convert_service_tickets.py RAW.xlsx -o OUT.xlsx
    python convert_service_tickets.py RAW.xlsx --map mapping.json   # ghi đè/bổ sung mapping

mapping.json (tùy chọn), ví dụ:
    {"users": {"abcxyz": "Nguyễn Văn A"},
     "agents": {"wintech jsc": "Wintech"},
     "vendors": {"fortinet": "Fortinet"},
     "roster": ["Nguyễn Văn A"]}
Username/đại lý/hãng chưa có mapping sẽ được giữ nguyên và in cảnh báo cuối lệnh.
"""
import argparse, json, re, sys, unicodedata
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

SHEET_NAME = 'Data_sạch'
REQUIRED = ['ID','Tên','Tên Hãng','Tên khách hàng','Đại lý','Sale','Người tạo','Người thực hiện','Tạo lúc',
            'Cập nhật lần cuối lúc','Thời gian giao hàng','License start date','License end date',
            'Luồng nhập hàng','Tình trạng Hợp đồng với khách hàng','Trạng thái','Tên khối']
OPTIONAL = ['No.','Công ty ký HĐ','Biên bản nghiệm thu thanh lý','SLA/Thời hạn','Quá hạn','Mô tả phiếu','Thông tin hàng hóa (BOM)']

# ---- Cấu hình mapping (sửa tại đây hoặc dùng --map) ----
ROSTER=['Vũ Việt Anh','Đỗ Tuấn Anh','Lê Tuấn Anh','Hoàng Việt Anh','Phạm Thành Chung','Lê Trọng Đại','Nguyễn Công Đoàn','Phạm Mạnh Hà','Nguyễn Thu Hằng','Đoàn Thị Hiền','Trần Hoàng Hiệp','Đinh Văn Hiệu','Nguyễn Thị Huệ','Lê Huy Hùng','Phạm Quốc Hùng','Trần Quang Hưng','Hà Thị Mai Hương','Ngọ Lan Hương','Nguyễn Thị Hường','Nguyễn Văn Hiệp','Mai Văn Khá','Tạ Diên Khải','Nguyễn Văn Khánh','Khổng Đức Kiên','Nguyễn Trung Kiên','Phí Ngọc Linh','Phan Việt Linh','Nguyễn Văn Minh','Nguyễn Đình Minh','Phạm Văn Nam','Nguyễn Thị Nga','Phan Quế Nghiêm','Bùi Minh Ngọc','Nguyễn Thị Hồng Nhung','Lê Thị Phương','Vũ Thị Lan Phương','Lý Anh Quốc','Vy Công Quý','Vũ Thị Thanh Quý','Phạm Văn Quyết','Vũ Xuân Sơn','Nguyễn Mỹ Tâm','Phạm Thanh Tâm','Phạm Duy Thắng','Phùng Văn Thành','Hoàng Văn Thảo','Dương Văn Thế','Đinh Văn Thuỷ','Nguyễn Duy Tiến','Trần Thị Minh Trâm','Đặng Đình Trường','Phạm Cao Anh Tú','Trần Văn Tuấn','Nguyễn Thị Tuyển','Nguyễn Đức Phúc Tường','Lương Thị Tuyết Trinh','Lê Phạm Hà Tân','Vũ Ngọc Phương','Đỗ Trung Hiếu','Trần Thị Mai Anh','Nguyễn Thị Thu Thuỷ','Khương Anh Khôi','Hoàng Thị Hương','Nguyễn Thị Thảo Lam','Nguyễn Duy Thắng','Đồng Văn Khoa','Nguyễn Thị Mỹ']
UMAP={'huongngo':'Ngọ Lan Hương','huenguyen':'Nguyễn Thị Huệ','tramtran':'Trần Thị Minh Trâm','tuantran':'Trần Văn Tuấn','tupham':'Phạm Cao Anh Tú','nhungnguyen':'Nguyễn Thị Hồng Nhung','thuydinh':'Đinh Văn Thuỷ','minhnguyen':'Nguyễn Văn Minh','minhnguyen2':'Nguyễn Đình Minh','kienkhong':'Khổng Đức Kiên','nganguyen':'Nguyễn Thị Nga','anhtranmai':'Trần Thị Mai Anh','phuongle':'Lê Thị Phương','huongha':'Hà Thị Mai Hương','linhphan':'Phan Việt Linh','hungle':'Lê Huy Hùng'}
VMAP={'kaspersky':'Kaspersky','hilstone':'Hillstone','infoexpress':'InfoExpress','inforexpress':'InfoExpress','barracuda':'Barracuda','safetica':'Safetica','opentext':'OpenText','optswat':'OPSWAT','gtb':'GTB','hcl':'HCL','netscout':'NetScout','netgear':'Netgear','cloudflare':'Cloudflare','progress':'Progress','zecurion':'Zecurion','qualys':'Qualys','acronis':'Acronis','sophos':'Sophos','delinea':'Delinea','module fs':'Module FS','owl cyber defense':'Owl Cyber Defense'}


def convert(src, mapping=None):
    mapping = mapping or {}
    global ROSTER, UMAP, VMAP
    ROSTER = list(ROSTER) + [r for r in mapping.get('roster', []) if r not in ROSTER]
    UMAP = {**UMAP, **mapping.get('users', {})}
    VMAP = {**VMAP, **{k.lower(): v for k, v in mapping.get('vendors', {}).items()}}
    df = pd.read_excel(src, dtype=str)
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        sys.exit('File raw thiếu cột bắt buộc: ' + ', '.join(missing))
    for c in OPTIONAL:
        if c not in df.columns: df[c] = None
    n0 = len(df)
    warn = []
    # ---------- customer
    KEYW=r'(tên|company|e\.?u\.?|end[- ]?user|sử dụng|bên|đơn vị|name|khách hàng|chủ đầu tư|thông tin)'
    DROP=re.compile(r'^\W*(địa chỉ|address|điện thoại|phone|tel\b|sđt|người liên hệ|liên hệ|contact|e-?mail|mst|mã số thuế|website)',re.I)
    PLACEHOLDER=re.compile(r'(theo file|thể hiện trên|trên license|theo tài liệu|theo hợp đồng|đính kèm|sale bổ sung|thông tin khách hàng trên|chưa có|^-+$|^n/?a$|^tbd)',re.I)
    def clean_label(l):
        l=l.replace(' ',' ').strip()
        l=re.sub(r'^[\s\-•·▪\*\d\.\)\t]+','',l)
        for _ in range(2):
            m=re.match(r'^([^:]{0,60}):\s*(.+)$',l)
            if m and re.search(KEYW,m.group(1),re.I): l=m.group(2)
        l=re.sub(r'^(ên|ủ đầu tư là|là)\s*:?\s*(?=[A-ZÀ-Ỹ])','',l)
        l=re.split(r'\s*(?:\b(?:địa chỉ|address|điện thoại|phone|người liên hệ|e-?mail)\s*:)',l,flags=re.I)[0]
        return re.sub(r'\s+',' ',l).strip(' -•:;,.')
    def cust_text(t):
        if not isinstance(t,str): return None
        for l in t.split('\n'):
            if not l.strip(): continue
            c=clean_label(l)
            if not c or DROP.match(c) or DROP.match(l.strip()): continue
            if re.fullmatch(KEYW+r'.{0,40}',c,re.I) and c.endswith(':'): continue
            return c
        return None
    def cust_name(n):
        m=re.search(r'EU\s*:\s*(.+)$',n,re.I)
        if m: return m.group(1).strip()
        p=n.split('_')
        return p[-1].strip() if len(p)>1 else None
    kt=df['Tên khách hàng'].map(cust_text); kn=df['Tên'].map(cust_name)
    def pick(a,b):
        a=a if isinstance(a,str) else None; b=b if isinstance(b,str) else None
        if a and not PLACEHOLDER.search(a) and len(a)>=3: return a,'Form khách hàng'
        if b: return b,'Tên phiếu'
        return None,'Không xác định'
    res=[pick(a,b) for a,b in zip(kt,kn)]
    df['KH']=[r[0] for r in res]; df['KH_nguon']=[r[1] for r in res]
    df['KH_tu_ten']=kn

    # ---------- vendor
    pass
    def vendor(r):
        v=r['Tên Hãng'].strip()
        if v=='Sản phẩm khác':
            n=r['Tên'].lower()
            if 'synology' in n: return 'Synology'
            if 'yokogawa' in n: return 'Yokogawa'
            return 'Khác'
        return VMAP.get(v.lower(),v)
    df['Hang']=df.apply(vendor,axis=1)

    # ---------- agent
    def akey(s): 
        s=unicodedata.normalize('NFC',s).lower().strip()
        s=re.sub(r'\s+',' ',s); return s
    AGENT_ALIAS={'wintech':'Wintech','công ty cổ phần wintech':'Wintech','công ty cổ phần phát triển đô thị wintech':'Wintech'}
    AGENT_ALIAS.update({'ntshn':'NTSHN','nts hn':'NTSHN'})
    def agent(s):
        if not isinstance(s,str) or not s.strip(): return None
        k=akey(s)
        if k in AGENT_ALIAS: return AGENT_ALIAS[k]
        s2=re.sub(r'\s+',' ',s.replace(' ',' ')).strip()
        return s2
    df['Dai_ly']=df['Đại lý'].map(agent)
    # unify same agent differing only by case
    canon={}
    for v in df['Dai_ly'].dropna():
        canon.setdefault(akey(v),v)
    # prefer a non-ALLCAPS spelling when variants exist
    for v in df['Dai_ly'].dropna().unique():
        k=akey(v)
        if canon[k].isupper() and not v.isupper(): canon[k]=v
    df['Dai_ly']=df['Dai_ly'].map(lambda v: canon[akey(v)] if isinstance(v,str) else None)

    # ---------- people
    def user(s): return s.lstrip('@').strip() if isinstance(s,str) else None
    df['Nguoi_tao']=df['Người tạo'].map(user)
    df['Nguoi_thuc_hien']=df['Người thực hiện'].map(user)
    def titlefix(s):
        return ' '.join(w[:1].upper()+w[1:] if w.islower() else w for w in s.split())
    pass
    pass
    
    df['Nguoi_tao_ten']=df['Nguoi_tao'].map(UMAP); df['Nguoi_thuc_hien_ten']=df['Nguoi_thuc_hien'].map(UMAP)
    df['Nguoi_tao_ten']=df['Nguoi_tao_ten'].fillna(df['Nguoi_tao']); df['Nguoi_thuc_hien_ten']=df['Nguoi_thuc_hien_ten'].fillna(df['Nguoi_thuc_hien'])
    _unm=set(df.Nguoi_tao[~df.Nguoi_tao.isin(UMAP)].dropna())|set(df.Nguoi_thuc_hien[~df.Nguoi_thuc_hien.isin(UMAP)].dropna())
    def nk(x): return re.sub(r'\s+',' ',''.join(c for c in unicodedata.normalize('NFD',x.lower().replace('đ','d')) if unicodedata.category(c)!='Mn')).strip()
    RK={nk(r):r for r in ROSTER}
    def salefix(x):
        x=re.sub(r'\s+',' ',x).strip()
        if x.lower()=='nguyễn hồng nhung': return 'Nguyễn Thị Hồng Nhung'
        return RK.get(nk(x),x)
    df['Sale_c']=df['Sale'].map(salefix)
    df['Sale_trong_ds']=df['Sale_c'].isin(ROSTER)
    _old=df['Sale'].map(lambda s: re.sub(r'\s+',' ',s).strip().title() if s.strip().lower()=='nguyễn hồng nhung' else re.sub(r'\s+',' ',s).strip())

    # ---------- dates
    def dt(c,fmt): return pd.to_datetime(df[c],format=fmt,errors='coerce')
    df['d_tao']=dt('Tạo lúc','%H:%M %d/%m/%Y'); df['d_cn']=dt('Cập nhật lần cuối lúc','%H:%M %d/%m/%Y')
    df['d_giao']=dt('Thời gian giao hàng','%d/%m/%Y'); df['d_ls']=dt('License start date','%d/%m/%Y'); df['d_le']=dt('License end date','%d/%m/%Y')
    df['d_sla']=dt('SLA/Thời hạn','%H:%M %d/%m/%Y')
    done=df['Trạng thái']=='Hoàn thành'
    df['Thang_tao']=df['d_tao'].dt.strftime('%Y-%m')
    df['TG_xu_ly']=((df['d_cn']-df['d_tao']).dt.total_seconds()/86400).round(1).where(done)

    # ---------- free-text PII mask
    def mask(s):
        if not isinstance(s,str): return s
        s=re.sub(r'[\w.\-+]+@[\w\-]+(\.[\w\-]+)+','[email]',s)
        s=re.sub(r'(?<!\d)(?:\+?84|0)[\d .\-]{8,12}\d','[sdt]',s)
        return re.sub(r'\s+',' ',s.replace(' ',' ')).strip()
    df['Mo_ta']=df['Mô tả phiếu'].map(mask)
    df['BOM']=df['Thông tin hàng hóa (BOM)'].map(mask)
    df['Ten_phieu']=df['Tên'].map(lambda s: re.sub(r'\s+',' ',s).strip())

    # ---------- flags
    def flag(r):
        f=[]
        if r['Trạng thái']=='Hoàn thành' and r['Tình trạng Hợp đồng với khách hàng']=='Chưa ký': f.append('Hoàn thành nhưng HĐ chưa ký')
        if r['Trạng thái']=='Hoàn thành' and r['Tình trạng Hợp đồng với khách hàng']=='KH đã ký nhưng chưa trả Hợp đồng': f.append('Hoàn thành nhưng KH chưa trả HĐ')
        if r['Quá hạn']=='Có': f.append('Quá hạn SLA')
        if r['KH_nguon']!='Form khách hàng': f.append('Tên KH lấy từ tên phiếu/không rõ - cần review')
        if pd.isna(r['Đại lý']): f.append('Thiếu đại lý')
        return '; '.join(f) if f else None
    df['Co_review']=df.apply(flag,axis=1)

    # ---------- output table
    cols=[('No.','No. (gốc)'),('ID','ID phiếu'),('Ten_phieu','Tên phiếu'),('Hang','Hãng'),('KH','Khách hàng (chuẩn hóa)'),('KH_nguon','Nguồn tên KH'),
    ('Dai_ly','Đại lý'),('Sale_c','Sale'),('Nguoi_tao_ten','Người tạo'),('Nguoi_thuc_hien_ten','Người thực hiện'),
    ('Luồng nhập hàng','Luồng nhập hàng'),('Tình trạng Hợp đồng với khách hàng','Tình trạng HĐ với KH'),('Công ty ký HĐ','Công ty ký HĐ (*)'),
    ('Biên bản nghiệm thu thanh lý','Biên bản nghiệm thu/thanh lý (*)'),
    ('d_tao','Ngày tạo'),('d_cn','Cập nhật lần cuối'),('Thang_tao','Tháng tạo'),('d_giao','Ngày giao hàng'),('d_ls','License start (*)'),('d_le','License end'),
    ('Tên khối','Bước hiện tại'),('Trạng thái','Trạng thái'),('TG_xu_ly','Thời gian xử lý (ngày, phiếu hoàn thành)'),
    ('d_sla','SLA/Hạn (*)'),('Quá hạn','Quá hạn (*)'),('Co_review','Cờ review'),('Mo_ta','Mô tả phiếu'),('BOM','BOM'),('Nguoi_tao_ten','Người tạo (username)'),('Nguoi_thuc_hien_ten','Người thực hiện (username)')]
    out=df[[c for c,_ in cols]].copy(); out.columns=[n for _,n in cols]
    out['No. (gốc)']=out['No. (gốc)'].astype(int); out['ID phiếu']=out['ID phiếu'].astype(int)



    cols=[('No.','No. (gốc)'),('ID','ID phiếu'),('Ten_phieu','Tên phiếu'),('Hang','Hãng'),('KH','Khách hàng (chuẩn hóa)'),('KH_nguon','Nguồn tên KH'),
    ('Dai_ly','Đại lý'),('Sale_c','Sale'),('Nguoi_tao_ten','Người tạo'),('Nguoi_thuc_hien_ten','Người thực hiện'),
    ('Luồng nhập hàng','Luồng nhập hàng'),('Tình trạng Hợp đồng với khách hàng','Tình trạng HĐ với KH'),('Công ty ký HĐ','Công ty ký HĐ (*)'),
    ('Biên bản nghiệm thu thanh lý','Biên bản nghiệm thu/thanh lý (*)'),
    ('d_tao','Ngày tạo'),('d_cn','Cập nhật lần cuối'),('Thang_tao','Tháng tạo'),('d_giao','Ngày giao hàng'),('d_ls','License start (*)'),('d_le','License end'),
    ('Tên khối','Bước hiện tại'),('Trạng thái','Trạng thái'),('TG_xu_ly','Thời gian xử lý (ngày, phiếu hoàn thành)'),
    ('d_sla','SLA/Hạn (*)'),('Quá hạn','Quá hạn (*)'),('Co_review','Cờ review'),('Mo_ta','Mô tả phiếu'),('BOM','BOM'),
    ('Nguoi_tao_ten','Người tạo (username)'),('Nguoi_thuc_hien_ten','Người thực hiện (username)')]
    out=df[[c for c,_ in cols]].copy(); out.columns=[n for _,n in cols]
    out['No. (gốc)']=pd.to_numeric(out['No. (gốc)'],errors='coerce').fillna(pd.Series(range(1,len(out)+1),index=out.index)).astype(int)
    out['ID phiếu']=pd.to_numeric(out['ID phiếu'],errors='coerce').astype('Int64')
    # cảnh báo mapping thiếu
    um=_unm
    if um: warn.append('Username chưa có trong mapping (giữ nguyên username): '+', '.join(sorted(map(str,um))))
    badv=sorted(set(df.Hang.dropna())-set(VMAP.values())-{'Synology','Yokogawa','Khác'})
    if badv: warn.append('Hãng chưa có trong bảng chuẩn hóa (giữ nguyên): '+', '.join(badv))
    ns=sorted(set(df.loc[~df.Sale_trong_ds,'Sale_c'].dropna()))
    if ns: warn.append('Sale không có trong danh sách nhân sự: '+', '.join(ns))
    bad_d=[c for c in ['Tạo lúc','Thời gian giao hàng','License end date'] if (df[c].notna()&df[{'Tạo lúc':'d_tao','Thời gian giao hàng':'d_giao','License end date':'d_le'}[c]].isna()).any()]
    if bad_d: warn.append('Có ngày sai định dạng ở cột: '+', '.join(bad_d))
    return out, warn

def write_xlsx(out, path):
    H=Font(bold=True,color='FFFFFF'); HF=PatternFill('solid',fgColor='1F4E78')
    wb=Workbook(); ws=wb.active; ws.title=SHEET_NAME
    ws.append(list(out.columns))
    for r in out.itertuples(index=False):
        ws.append([None if pd.isna(v) else (v.to_pydatetime() if isinstance(v,pd.Timestamp) else (int(v) if hasattr(v,'item') and isinstance(v.item(),int) else v)) for v in r])
    for c in ws[1]: c.font=H; c.fill=HF; c.alignment=Alignment(wrap_text=True,vertical='center')
    ws.freeze_panes='A2'; ws.auto_filter.ref=ws.dimensions; ws.row_dimensions[1].height=45
    W=[8,9,48,13,42,15,28,22,22,22,12,26,12,14,17,17,10,13,13,13,26,12,14,17,10,40,50,40,16,16]
    for i,w in enumerate(W,1): ws.column_dimensions[get_column_letter(i)].width=w
    for i,n in enumerate(out.columns,1):
        L=get_column_letter(i)
        if n in('Ngày tạo','Cập nhật lần cuối','SLA/Hạn (*)'):
            for c in ws[L][1:]: c.number_format='yyyy-mm-dd hh:mm'
        elif n in('Ngày giao hàng','License start (*)','License end'):
            for c in ws[L][1:]: c.number_format='yyyy-mm-dd'
    wb.save(path)

def main():
    ap=argparse.ArgumentParser(description='Convert raw Base service-tickets Excel -> 1 sheet Data_sạch')
    ap.add_argument('src'); ap.add_argument('-o','--out'); ap.add_argument('--map',help='JSON mapping bổ sung')
    a=ap.parse_args()
    mp=json.load(open(a.map,encoding='utf-8')) if a.map else None
    out,warn=convert(a.src,mp)
    dst=a.out or re.sub(r'\.xlsx?$','',a.src,flags=re.I)+'_CLEAN.xlsx'
    write_xlsx(out,dst)
    print(f'OK: {len(out)} dòng, {len(out.columns)} cột -> {dst}  (sheet "{SHEET_NAME}")')
    for w in warn: print('CẢNH BÁO:',w)

if __name__=='__main__':
    main()
