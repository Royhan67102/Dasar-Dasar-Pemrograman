from tkinter import *
from pygame import mixer
from openpyxl import Workbook
from tkinter import Tk, Label, Button, filedialog, font


window = Tk()
window.config(background="light blue")
window.geometry('500x500+400+100')
workbook = Workbook()
sheet = workbook.active

volume_current=float(0.5)

#def music
def play_song():
    filename=filedialog.askopenfilename(initialdir="C:/", title="Pilih lagu")
    judul_lagu=filename.split("/")
    judul_lagu=judul_lagu[-1]
    try:
        mixer.init()
        mixer.music.load(filename)
        mixer.music.set_volume(volume_current)
        mixer.music.play() 
    except Exception as e:
        print(e)

def pause():
    mixer.music.pause()

def unpause():
    mixer.music.unpause()


#def hitung input nilai 
def hitung_nilai():
    nama = entry_nama.get()
    nim = entry_nim.get()
    nilai_absen = float(entry_absen.get())
    nilai_tugas = float(entry_tugas.get())
    nilai_uts = float(entry_uts.get())
    nilai_uas = float(entry_uas.get())

    nilai_akhir = (nilai_absen * 0.1) + (nilai_tugas * 0.2) + (nilai_uts * 0.3) + (nilai_uas * 0.4)

    if nilai_akhir >= 80:
        grade = "A"
    elif nilai_akhir >= 70:
        grade = "B"
    elif nilai_akhir >= 60:
        grade = "C"
    elif nilai_akhir >= 50:
        grade = "D"
    else:
        grade = "E"

    output.config(state=NORMAL)
    output.delete(1.0, END)
    output.insert(END, f"Nama: {nama}\n")
    output.insert(END, f"NIM: {nim}\n")
    output.insert(END, f"Nilai Akhir: {nilai_akhir}\n")
    output.insert(END, f"Grade: {grade}\n")
    output.config(state=DISABLED)

#def tombol insert dan save
# Definisi fungsi untuk menghitung nilai
def hitung_nilai():
    nama = entry_nama.get()
    nim = entry_nim.get()
    nilai_absen = float(entry_absen.get())
    nilai_tugas = float(entry_tugas.get())
    nilai_uts = float(entry_uts.get())
    nilai_uas = float(entry_uas.get())

    nilai_akhir = (nilai_absen * 0.1) + (nilai_tugas * 0.2) + (nilai_uts * 0.3) + (nilai_uas * 0.4)

    if nilai_akhir >= 80:
        grade = "A"
    elif nilai_akhir >= 70:
        grade = "B"
    elif nilai_akhir >= 60:
        grade = "C"
    elif nilai_akhir >= 50:
        grade = "D"
    else:
        grade = "E"

    output.config(state=NORMAL)
    output.delete(1.0, END)
    output.insert(END, f"Nama: {nama}\n")
    output.insert(END, f"NIM: {nim}\n")
    output.insert(END, f"Nilai Akhir: {nilai_akhir}\n")
    output.insert(END, f"Grade: {grade}\n")
    output.config(state=DISABLED)

# Definisi fungsi untuk menyimpan data ke Excel
def saveData():
    sheet.append([entry_nama.get(), entry_nim.get(), entry_absen.get(), entry_tugas.get(), entry_uts.get(), entry_uas.get()])
    workbook.save(filename=str(entry_nama.get()) + ".xlsx")

# Definisi fungsi untuk menghapus data dari input fields
def delateData():
    entry_nama.delete(0, 'end')
    entry_nim.delete(0, 'end')
    entry_absen.delete(0, 'end')
    entry_tugas.delete(0, 'end')
    entry_uts.delete(0, 'end')
    entry_uas.delete(0, 'end')

# Definisi fungsi untuk menyisipkan data ke dalam Excel
def insertData():
    global num
    num = num + 1
    sheetnum = num + 3

    sheet['A' + str(sheetnum)] = entry_nama.get()
    DataNama = sheet['A' + str(sheetnum)]

    sheet['B' + str(sheetnum)] = entry_nim.get()
    DataNama = sheet['A' + str(sheetnum)]



#letak sheet
sheet['A1'] = "INPUT NILAI MAHASISWA\t"
A1 = sheet['A1']

sheet['A3'] = "Nama\t"
A3 = sheet['A3']
sheet['B3'] = "NIM\t"
B3 = sheet['B3']
sheet['c3'] = "Nilai Absen\t"
B3 = sheet['B3']
sheet['D3'] = "Nilai Tugas\t"
B3 = sheet['B3']
sheet['E3'] = "Nilai UTS\t"
B3 = sheet['B3']
sheet['F3'] = "Nilai UAS\t"
B3 = sheet['B3']



button = Button(window, text="Music", command=play_song, activebackground="green")
button.pack()
button.place(rely=0, relx=0)
button = Button(window, text="II", command=pause, activebackground="green")
button.pack()
button.place(rely=0, relx=0.085)
button = Button(window, text="I>", command=unpause, activebackground="green")
button.pack()
button.place(rely=0, relx=0.12)


label_nama = Label(window, text="Nama:", background="light blue")
label_nama.pack()

entry_nama = Entry(window)
entry_nama.pack()
entry_nama.get()

label_nim = Label(window, text="NIM:", background="light blue")
label_nim.pack()

entry_nim = Entry(window)
entry_nim.pack()
entry_nim.get()

label_absen = Label(window, text="Nilai Absen:", background="light blue")
label_absen.pack()

entry_absen = Entry(window)
entry_absen.pack()
entry_absen.get()

label_tugas = Label(window, text="Nilai Tugas:", background="light blue")
label_tugas.pack()

entry_tugas = Entry(window)
entry_tugas.pack()
entry_tugas.get()

label_uts = Label(window, text="Nilai UTS:", background="light blue")
label_uts.pack()

entry_uts = Entry(window)
entry_uts.pack()
entry_uts.get()

label_uas = Label(window, text="Nilai UAS:", background="light blue")
label_uas.pack()

entry_uas = Entry(window)
entry_uas.pack()
entry_uas.get()

button = Button(window, text="Hitung", command=hitung_nilai, activebackground="green")
button.pack()
button = Button(window, text="Save", activebackground="green", command=saveData)
button.pack()
button = Button(window, text="Delate", activebackground="red", command=delateData)
button.pack()
button = Button(window, text="Insert", activebackground="green", command=insertData)
button.pack()

output = Text(window, height=10, width=30)
output.config(state=DISABLED)
output.pack()


screenwidth = window.winfo_screenwidth()
screenheighht = window.winfo_screenheight()
window.title('Input Nilai Mahasiswa')

window.mainloop()