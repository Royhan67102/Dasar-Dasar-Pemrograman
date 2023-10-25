#luas persegi
ket = """masukkan pilihan:
1. untuk menghitung luas persegi
2. untuk menghitung luas lingkaran
3. untuk menghitung luas segitiga
"""
pilihan = input (ket)

match pilihan:
    case "1":
        print ("memilih 1 untuk menghitung luas persegi")
        sisi = int (input("masukan sisi?"))
        luasp = sisi * sisi
        print ("luas persegi yang sisinya",sisi,"adalah",luasp)

    case "2":
        print ("memilih 2 untuk menghitung luas lingkaran")
        jari2 = float (input("masukkan jari2?"))
        luasL = 3.14*jari2*jari2
        print("luas lingkaran yang jari2nya",jari2,"adalah",int(luasL))

    case "3":
        print ("memilih 3 untuk menghitung luas segitiga")
        alas = float (input("masukkan panjang alas segitiga:"))
        tinggi = float (input("masukkan tinggi segitiga"))
        luasS = 0.5*alas*tinggi
        print("luas segitiga yang alasnya",alas,"dan tingginya",tinggi,"adalah",luasS)
    case _:
        print("salah memilih pilihan")
        

