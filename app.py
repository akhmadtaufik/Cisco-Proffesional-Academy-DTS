from tkinter import *
import sqlite3
import requests
import csv

root = Tk()
root.title('Scraping Produk From Shopee')
root.geometry("1300x400")

conn = sqlite3.connect('data.db')
c = conn.cursor()

# Membuat Tabel dari database data.db
'''
c.execute("""CREATE TABLE data_scraping(
    nama_produk text,
    harga integer,
    stok integer,
    rating decimal,
    kota text
)""")
'''

def submit():
    try:
        conn = sqlite3.connect('data.db')
        c = conn.cursor()

        c.execute("""CREATE TABLE IF NOT EXISTS data_scraping(
            nama_produk text,
            harga integer,
            stok integer,
            rating decimal,
            kota text
            )""")

        # Mendefinisikan url yang terhubung dengan API Shoppe
        url = 'https://shopee.co.id/api/v4/search/search_items'
        newest = 0
        count = 0
        for page in range(1,6):
            newest += 50
            # Membuat parameter yang berhubungan dengan detail pencarian produk di Shopee
            parameter = {
                'by': 'sales',
                'keyword': KeyWord.get(),
                'limit': 50,
                'locations': LokasiBar.get(),
                'newest': newest,
                'order': 'desc',
                'page_type': 'search',
                'skip_autocorrect': 1,
                'version': 2
            }

            # Parsing JSON data dengan Key Values 'items'
            api_request = requests.get(url, params=parameter).json()
            api = api_request['items']

            for produk in api:
                conn = sqlite3.connect('data.db')
                c = conn.cursor()
                count += 1
                nama_produk = produk['item_basic']['name']
                harga = (produk['item_basic']['price_max']//100000)
                stok_produk = (produk['item_basic']['stock'])
                rating = round(produk['item_basic']['item_rating']['rating_star'], 2)
                lokasi_toko = produk['item_basic']['shop_location']

                # Memasukan data hasil scraping ke dalam database data.db
                c.execute("INSERT INTO data_scraping VALUES (:nama_produk, :harga, :stok_produk, :rating, :lokasi_toko)",
                                {
                                   'nama_produk' : nama_produk,
                                   'harga' : harga,
                                   'stok_produk' : stok_produk,
                                   'rating' : rating,
                                   'lokasi_toko' : lokasi_toko
                                })
                conn.commit()
                conn.close()
    except Exception:
        api = 'Erroor...'

def query():
    try:
        conn = sqlite3.connect('data.db')
        c = conn.cursor()

        # Menampilkan 10 data pertama pada aplikasi
        c.execute("SELECT *, oid FROM data_scraping")
        data = c.fetchmany(10)
        #print(data)

        print_data = ''
        for record in data:
            print_data += str(record[5]) + "\t" + " | Nama Produk :" + str(record[0]) + " | Harga : " + str(record[1]) + " | Stok :" + str(record[2]) + " | Rating : " + str(record[3]) + " | Lokasi : " + str(record[4]) + "\n"
            

        QLabel = Label(root, text=print_data)
        QLabel.grid(row=5, column=0, columnspan=3, sticky=W, padx=10, pady=(5, 0))

    except Exception:
        QLabel = Label(root, text='Database Tidak Ditemukan')
        QLabel.grid(row=5, column=0, columnspan=3, sticky=W, padx=10, pady=(5, 0))

def AllData():
    try:
        #global data
        conn = sqlite3.connect('data.db')
        c = conn.cursor()

        # Menyimpan semua data ke dalam format csv
        c.execute("SELECT *, oid FROM data_scraping")
        data =  c.fetchall()

        filename = 'data_{}_{}.csv'.format(KeyWord.get(), LokasiBar.get())
        write = csv.writer(open(filename, 'w', newline='', encoding='utf-8'))
        header = ['Nama Produk', 'Harga', 'Stok', 'Rating', 'Lokasi', 'Key ID']
        write.writerow(header)

        with open(filename, 'a', newline='', encoding='utf-8') as f:
            write = csv.writer(f, dialect='excel')
            for record in data:
                write.writerow(record)

    except Exception:
        warningLabel = Label(root, text='Tidak Ada Data Silahkan Masukan Keyword dan Lokasi pada Kolom Pencarian')
        warningLabel.grid(row=5, column=0, columnspan=3, sticky=W, padx=10, pady=(5, 0))

def clearData():
    try:
        # Hapus Semua Masukan pada Label
        KeyWord.delete(0, END)
        LokasiBar.delete(0, END)

        # Hapus tabel data_sraping pada database

        conn = sqlite3.connect('data.db')
        c = conn.cursor()

        c.execute("DROP TABLE data_scraping")

        conn.commit()
        conn.close()

        # Hapus Pencarian pad layar

        for label in root.grid_slaves():
            if int(label.grid_info()['row']) >= 5:
                label.destroy()

    except Exception:
        warningLabel = Label(root, text='Tidak Ada Data Silahkan Masukan Keyword dan Lokasi pada Kolom Pencarian')
        warningLabel.grid(row=5, column=0, columnspan=3, sticky=W, padx=10, pady=(5, 0))


KeyWordLabel = Label(root, text='Keyword')
KeyWordLabel.grid(row=0, column=0, sticky=W, padx=10)
KeyWord = Entry(root)
KeyWord.grid(row=1, column=0, padx=10, pady=5)

LokasiLabel = Label(root, text='Lokasi')
LokasiLabel.grid(row=0, column=1, sticky=W, padx=10)
LokasiBar = Entry(root)
LokasiBar.grid(row= 1, column=1, padx=10, pady=5)

searchButton = Button(root, text='Masukan ke Database', command=submit)
searchButton.grid(row=1, column=2)

showData = Button(root, text="Tampilkan Data", command=query)
showData.grid(row=2, column=0, columnspan=3, pady=10, padx=10, ipadx=137)

csvButton = Button(root, text="Simpan ke Excel", command=lambda: AllData())
csvButton.grid(row=3, column=0, columnspan=3, pady=10, padx=10, ipadx=137)

clearButton = Button(root, text='Bersihkan Data', command=clearData)
clearButton.grid(row=4, column=0, columnspan=3, pady=10, padx=10, ipadx=137)

conn.commit()
conn.close()

root.mainloop()