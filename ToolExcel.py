from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askopenfilename, askopenfile
import datetime
from xlrd import open_workbook
from xlutils.copy import copy

FontBig = ('Arial',13,'bold')
FontSmall = ('Arial',11,'bold')
BG = '#a7beaf'

window = Tk()

window.title('Tỷ lệ nguồn vốn ngắn hạn cho vay trung và dài hạn')

window.geometry('750x750')

window.config(bg=BG)

filetypes = [('Excel Files', '*.xls')]

TG_KKH = Label(text='TG KKH: ', font=FontBig, bg=BG)
TG_KKH.place(x=5, y=300)

TGCKH_Less1Year = Label(text='TGCKH có thời hạn còn lại đến 1 năm: ', font=FontBig, bg=BG)
TGCKH_Less1Year.place(x=5, y=340)

TGCKH_More1Year = Label(text='TGCKH có thời hạn còn lại trên 1 năm: ', font=FontBig, bg=BG)
TGCKH_More1Year.place(x=5, y=380)

TongDuNo = Label(text='Tổng dư nợ:', font=FontBig, bg=BG)
TongDuNo.place(x=5, y=10)

DuNoGocDenHanConHanDuoi1Nam = Label(text='Dư nợ gốc đến hạn có thời hạn còn lại đến 1 năm:', font=FontBig, bg=BG)
DuNoGocDenHanConHanDuoi1Nam.place(x=5, y=50)

VonDieuLe = Label(text='Vốn điều lệ:', font=FontBig, bg=BG)
VonDieuLe.place(x=5, y=90)

CacQuyDuTru = Label(text='Các quỹ dự trữ 61 (611, 612, 613):', font=FontBig, bg=BG)
CacQuyDuTru.place(x=5, y=130)

VayTCTDKhacCoThoiHanTren1Nam = Label(text='Vay TCTD khác có thời hạn còn lại trên 1 năm:', font=FontBig, bg=BG)
VayTCTDKhacCoThoiHanTren1Nam.place(x=5, y=170)

VayTCTDKhacCoThoiHanDuoi1Nam = Label(text='Vay TCTD khác có thời hạn còn lại đến 1 năm:', font=FontBig, bg=BG)
VayTCTDKhacCoThoiHanDuoi1Nam.place(x=5, y=210)

Ngay = Label(text='Ngày:', font=FontBig, bg=BG)
Ngay.place(x=5, y=420)

Thang = Label(text='Tháng:', font=FontBig, bg=BG)
Thang.place(x=125, y=420)

Nam = Label(text='Năm:', font=FontBig, bg=BG)
Nam.place(x=260, y=420)

TongDuNoChoVayTrungVaDaiHan = Label(text='Tổng dư nợ cho vay trung và dài hạn:', font=FontBig, bg=BG)
TongDuNoChoVayTrungVaDaiHan.place(x=5, y=520)

TongNguonVonTrungVaDaiHan = Label(text='Tổng nguồn vốn trung và dài hạn:', font=FontBig, bg=BG)
TongNguonVonTrungVaDaiHan.place(x=5, y=560)

NguonVonNganHan = Label(text='Nguồn vốn ngắn hạn:', font=FontBig, bg=BG)
NguonVonNganHan.place(x=5, y=600)

A_label = Label(text='Tỷ lệ nguồn vốn ngắn hạn cho vay trung và dài hạn:', font=FontBig, bg=BG)
A_label.place(x=5, y=640)

TongDuNo_entry = Entry(width=20, font=FontBig)
TongDuNo_entry.place(x=450, y=10)

DuNoGocDuoi1Nam_entry = Entry(width=20, font=FontBig)
DuNoGocDuoi1Nam_entry.place(x=450, y=50)

VonDieuLe_entry = Entry(width=20, font=FontBig)
VonDieuLe_entry.place(x=450, y=90)

CacQuyDuTru_entry = Entry(width=20, font=FontBig)
CacQuyDuTru_entry.place(x=450, y=130)

VayTCTDKhacCoThoiHanTren1Nam_entry = Entry(width=20, font=FontBig)
VayTCTDKhacCoThoiHanTren1Nam_entry.place(x=450, y=170)

VayTCTDKhacCoThoiHanDuoi1Nam_entry = Entry(width=20, font=FontBig)
VayTCTDKhacCoThoiHanDuoi1Nam_entry.place(x=450, y=210)

TGKKH_entry = Entry(width=20, font=FontBig)
TGKKH_entry.place(x=450, y=300)

TGCKH_Less1Year_entry = Entry(width=20, font=FontBig)
TGCKH_Less1Year_entry.place(x=450, y=340)

TGCKH_More1Year_entry = Entry(width=20, font=FontBig)
TGCKH_More1Year_entry.place(x=450, y=380)

FileName_label = Label(text='', font=FontBig, bg=BG)
FileName_label.place(x=200, y=260)

Ngay_entry = Entry(width=7, font=FontBig, justify='right')
Ngay_entry.place(x=55, y=420)

Thang_entry = Entry(width=7, font=FontBig, justify='right')
Thang_entry.place(x=185, y=420)

Nam_entry = Entry(width=7, font=FontBig, justify='right')
Nam_entry.place(x=305, y=420)

def choosefileC():
    try:
        file = askopenfile(mode='r', initialdir=r'C:\Users', title='Select File', filetypes=filetypes)
        s = str(file)
        a = s.find('name')
        b = s.find('mode')
        filename = s[a + 6:b - 5]
        book = open_workbook(filename + 'xls')
        sheet = book.sheet_by_index(0)
        sumOfLessNegative15 = 0
        sumOfNegative15To365 = 0
        sumOfMore365 = 0
        for i in range(0, sheet.nrows):
            if "/" in str(sheet.cell_value(rowx=i, colx=8)):
                dayI, monthI, yearI = str(sheet.cell_value(rowx=i, colx=8)).split('/')
                dayL = int(Ngay_entry.get())
                monthL = int(Thang_entry.get())
                yearL = int(Nam_entry.get())
                dateI = datetime.datetime(int(yearI), int(monthI), int(dayI), 0, 0, 0, 0)
                dateL = datetime.datetime(int(yearL), int(monthL), int(dayL), 0, 0, 0, 0)
                date = (dateI - dateL).days
                if date <= -15:
                    sumOfLessNegative15 += int(sheet.cell_value(rowx=i, colx=10))
                elif -15< date <= 365:
                    sumOfNegative15To365 += int(sheet.cell_value(rowx=i, colx=10))
                else:
                    sumOfMore365 += int(sheet.cell_value(rowx=i, colx=10))
        TGKKH_entry.delete(0, END)
        TGKKH_entry.insert(INSERT, f'{sumOfLessNegative15:,}')
        TGCKH_Less1Year_entry.delete(0, END)
        TGCKH_Less1Year_entry.insert(INSERT, f'{sumOfNegative15To365:,}')
        TGCKH_More1Year_entry.delete(0, END)
        TGCKH_More1Year_entry.insert(INSERT, f'{sumOfMore365:,}')
    except Exception as e:
        messagebox.showerror("Có lỗi", e)



def xuly(s):
    kytu = [',','.']
    s = s.strip()
    for j in kytu:
        if j in s:
            s = s.replace(j,'')
    return s

def cal():
    global tongDuNo, duNoGocDuoi1Nam, vonDieuLe, cacQuy, tgckhTren1Nam, vayTCTDTren1Nam, tgckhDuoi1Nam, tgkkh, vayTCTDDuoi1Nam, tongDuNoTrungVaDai, tongVonTrungVaDai, nguonVonNganHan, A, i

    try:
        tongDuNo = int(float(xuly(TongDuNo_entry.get())))
        duNoGocDuoi1Nam = int(float(xuly(DuNoGocDuoi1Nam_entry.get())))
        vonDieuLe = int(float(xuly(VonDieuLe_entry.get())))
        cacQuy = int(float(xuly(CacQuyDuTru_entry.get())))
        tgckhTren1Nam = int(float(xuly(TGCKH_More1Year_entry.get())))
        vayTCTDTren1Nam = int(float(xuly(VayTCTDKhacCoThoiHanTren1Nam_entry.get())))
        tgckhDuoi1Nam = int(float(xuly(TGCKH_Less1Year_entry.get())))
        tgkkh = int(float(xuly(TGKKH_entry.get())))
        vayTCTDDuoi1Nam = int(float(xuly(VayTCTDKhacCoThoiHanDuoi1Nam_entry.get())))
        tongDuNoTrungVaDai = tongDuNo - duNoGocDuoi1Nam
        tongVonTrungVaDai = vonDieuLe + cacQuy + tgckhTren1Nam + vayTCTDTren1Nam
        nguonVonNganHan = tgckhDuoi1Nam + tgkkh + vayTCTDDuoi1Nam
        A = float("{:.2f}".format((tongDuNoTrungVaDai - tongVonTrungVaDai) / nguonVonNganHan * 100))

        TongDuNoChoVayTrungVaDaiHan.config(text=f'Tổng dư nợ cho vay trung và dài hạn: {tongDuNoTrungVaDai:,}', fg='green', font = FontSmall)
        TongNguonVonTrungVaDaiHan.config(text=f'Tổng nguồn vốn trung và dài hạn: {tongVonTrungVaDai:,}', fg='green', font = FontSmall)
        NguonVonNganHan.config(text=f'Nguồn vốn ngắn hạn: {nguonVonNganHan:,}', fg='green', font = FontSmall)
        A_label.config(text=f"Tỷ lệ nguồn vốn ngắn hạn cho vay trung và dài hạn: {A:,}%", fg='green', font = FontBig)
        i=True

    except Exception as e:
        messagebox.showerror("Có lỗi", e)
        i=False

def save_file():
    if i:
        file = askopenfilename(initialdir=r'C:\Users',
                           title='Select File', filetypes=filetypes)
        s = str(file)
        a = s.find('name')
        b = s.find('mode')
        filename = s[a + 1:b - 2]
        w = copy(open_workbook(filename + 'xls', formatting_info=True))

        value = [tongDuNoTrungVaDai,tongDuNoTrungVaDai,tongDuNo,duNoGocDuoi1Nam,tongVonTrungVaDai,
                 vonDieuLe, cacQuy, tgckhTren1Nam, vayTCTDTren1Nam, nguonVonNganHan, tgckhDuoi1Nam, tgkkh, vayTCTDDuoi1Nam]

        row = 9
        for item in value:
            w.get_sheet(0).write(row, 5, item)
            row+=1
        w.get_sheet(0).write(3, 4, A)
        w.save(filename+'xls')
    else:
        messagebox.showerror("Có lỗi", "Lỗi dữ liệu")

def format_value():
    value0 = f'{int(xuly(TongDuNo_entry.get())):,}'
    value1 = f'{int(xuly(DuNoGocDuoi1Nam_entry.get())):,}'
    value2 = f'{int(xuly(VonDieuLe_entry.get())):,}'
    value3 = f'{int(xuly(CacQuyDuTru_entry.get())):,}'
    value4 = f'{int(xuly(VayTCTDKhacCoThoiHanTren1Nam_entry.get())):,}'
    value5 = f'{int(xuly(VayTCTDKhacCoThoiHanDuoi1Nam_entry.get())):,}'
    TongDuNo_entry.delete(0,END)
    TongDuNo_entry.insert(INSERT, value0)
    DuNoGocDuoi1Nam_entry.delete(0,END)
    DuNoGocDuoi1Nam_entry.insert(INSERT, value1)
    VonDieuLe_entry.delete(0, END)
    VonDieuLe_entry.insert(INSERT, value2)
    CacQuyDuTru_entry.delete(0, END)
    CacQuyDuTru_entry.insert(INSERT, value3)
    VayTCTDKhacCoThoiHanTren1Nam_entry.delete(0,END)
    VayTCTDKhacCoThoiHanTren1Nam_entry.insert(INSERT, value4)
    VayTCTDKhacCoThoiHanDuoi1Nam_entry.delete(0, END)
    VayTCTDKhacCoThoiHanDuoi1Nam_entry.insert(INSERT, value5)

choosefileAbtn = Button(text='Chọn file TGTK', width=15, height=1, font=('Arial', 15, 'bold'), bg=BG, command=choosefileC)
choosefileAbtn.place(x=5, y=250)

cal_btn = Button(text='Tính toán', width=15, height=1, font=('Arial', 15, 'bold'), bg=BG, command=cal)
cal_btn.place(x=260, y=460)

save_to_file = Button(text='Lưu vào file', width=15, height=1, font=('Arial', 15, 'bold'), bg=BG, command=save_file)
save_to_file.place(x=260, y=680)

format_value_btn = Button(text='Format', width=7, height=1, font=('Arial', 15, 'bold'), bg=BG, command=format_value)
format_value_btn.place(x=650, y=120)





























window.mainloop()