
import pandas as pd
import PySimpleGUI as sg
import os

# membuat Class untuk menampung method untuk kebutuham mengelola data


class Perpustakaan:
    # membuat constructer untuk menjenerate file excel sebagai database
    def __init__(self, file_path='perpustakaan.xlsx'):
        self.file_path = file_path

        # validasi jika file tidak ada maka buat file jika ada maka gunakan
        if os.path.exists(self.file_path):
            self.df = pd.read_excel(self.file_path)
        else:
            self.df = pd.DataFrame(
                columns=["id", 'Judul', 'Penulis', "penerbit", 'Tahun Terbit', "Jumlah Buku"])

    # method untuk memasukan data buku kedalam file excel
    def create_buku(self, Id, judul, penulis, penerbit, tahun_terbit, jumlah_buku):
        self.df = self.df._append(
            {"id": Id, 'Judul': judul, 'Penulis': penulis, "penerbit": penerbit, 'Tahun Terbit': tahun_terbit, 'Jumlah Buku': jumlah_buku}, ignore_index=True)
        self.save_to_excel()

    # method/behavior untuk mengganti data yang dipilih dari data yang sudah ada
    def update_buku(self, index, judul, penerbit, penulis, tahun_terbit, jumlah_buku):
        self.df.loc[index] = [self.df.loc[index, 'id'], judul, penulis,
                              penerbit, tahun_terbit, jumlah_buku]
        self.save_to_excel()

    # method/behavior untuk menghapus data yang dipilih dari data yang sudah ada
    def delete_buku(self, index):
        self.df = self.df.drop(index)
        self.save_to_excel()

    # method untuk membuat id
    def last_id(self):
        numeric_ids = pd.to_numeric(self.df['id'], errors='coerce')
        last_id = numeric_ids.max() if not numeric_ids.empty else 0

        if pd.isna(last_id):
            last_id = 0

        new_id = last_id + 1
        return new_id

    # method menyimpan perubahan data ke excel
    def save_to_excel(self):
        self.df.to_excel(self.file_path, index=False)

    # method untuk membaca data excel
    def read_buku(self):
        return self.df


# function exesekusi
def main():
    perpustakaan = Perpustakaan()  # pemanggilas class atau membuat object baru

    # membuat lyout
    layout = [
        [sg.Text('Judul', size=(10, 1)), sg.InputText(key='judul')],
        [sg.Text('Penulis', size=(10, 1)), sg.InputText(key='penulis')],
        [sg.Text('Penerbit', size=(10, 1)), sg.InputText(key='penerbit')],
        [sg.Text('Tahun Terbit', size=(10, 1)),
         sg.InputText(key='tahun_terbit')],

        [sg.Text('Jumlah Buku', size=(10, 1)),
         sg.InputText(key='Jumlah Buku')],
        [sg.Column(layout=[
            [sg.Button('Tambah', key='tambah'), sg.Button(
                'Perbarui', key='perbarui'), sg.Button('Hapus', key='hapus')],
        ], justification='center')],
        [sg.Column(layout=[[sg.Table(values=[], headings=["id", 'Judul', 'Penulis', 'Penerbit', 'Tahun Terbit', 'Jumlah Buku'],
                   auto_size_columns=True, justification='right', vertical_scroll_only=False, key='tabel')]], justification='center')],
    ]

    window = sg.Window('Menejement Data Buku Perpustakaan', layout)

    # action button / event button atau apa yang terjadi setelah meng klik button
    while True:
        new_id = perpustakaan.last_id()
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        elif event == 'tambah':
            sg.popup(new_id)
            perpustakaan.create_buku(
                new_id, values['judul'], values['penulis'], values['penerbit'], values['tahun_terbit'], values['Jumlah Buku'])
        elif event == 'perbarui':
            selected_row = window['tabel'].SelectedRows
            if selected_row:
                index = selected_row[0]
                sg.popup(index)

                perpustakaan.update_buku(
                    index, values['judul'], values['penulis'], values["penerbit"], values['tahun_terbit'], values['Jumlah Buku'])

        elif event == 'hapus':
            selected_row = window['tabel'].SelectedRows
            if selected_row:
                index = selected_row[0]
                perpustakaan.delete_buku(index)

        buku_df = perpustakaan.read_buku()
        window['tabel'].update(values=buku_df.values.tolist())

    window.close()


if __name__ == '__main__':
    main()
