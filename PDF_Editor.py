import os
import customtkinter as ctk
from tkinter import *
from functools import partial
from tkinter import filedialog
from tkinter import ttk, messagebox
from PyPDF2 import PdfWriter, PdfReader
from PIL import Image, ImageTk

class PDF_Editor:
    def __init__(self, root):
        self.window = root
        self.window.geometry("740x480")
        self.window.title('PDF Editor')

        # Color Options
        self.color_1 = "white"
        self.color_2 = "#C9B977"
        self.color_3 = "black"
        self.color_4 = '#960018'
        # self.color_4 = 'orange red'

        # Font Options
        self.font_1 = "Helvetica"
        self.font_2 = "Times New Roman"
        self.font_3 = "trebuchet ms"

        self.saving_location = ''
        self.PDF_path = []
        self.PDF_List = None

        # Menubar
        self.menubar = Menu(self.window)

        # Adding Edit Menu and its sub menus
        edit = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label='Menu', menu=edit)
        edit.add_command(label='Dividir PDF', command=partial(self.SelectPDF, 1))
        edit.add_command(label='Mesclar PDF', command=self.Merge_PDFs_Data)
        edit.add_separator()
        edit.add_command(label='Girar PDF', command=partial(self.SelectPDF, 2))

        # Adding About Menu
        about = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label='Sobre', menu=about)
        about.add_command(label='Sobre', command=self.AboutWindow)

        # Exit the Application
        exit = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label='Sair', menu=exit)
        exit.add_command(label='Sair', command=self.Exit)

        # Configuring the menubar
        self.window.config(menu=self.menubar)

        # Creating a Frame
        self.frame_1 = Frame(self.window, bg=self.color_2, width=740, height=480)
        self.frame_1.place(x=0, y=0)
        self.Home_Page()

    # Miscellaneous Functions
    def AboutWindow(self):
        messagebox.showinfo("Editor de PDF", "PDF Editor 13.05.24\nDeveloped by Tecla Fernandes")

    def ClearScreen(self):
        for widget in self.frame_1.winfo_children():
            widget.destroy()

    def Update_Path_Label(self):
        self.path_label.config(text=self.saving_location)

    def Update_Rotate_Page(self):
        self.saving_location = ''
        self.ClearScreen()
        self.Home_Page()

    def Exit(self):
        self.window.destroy()

    def Home_Page(self):
        self.ClearScreen()

    def SelectPDF(self, to_call):
        self.PDF_path = filedialog.askopenfilename(initialdir="/",
                                                   title="Selecione um arquivo PDF", filetypes=(("PDF files", "*.pdf*"),))
        if self.PDF_path:
            if to_call == 1:
                self.Split_PDF_Data()
            else:
                self.Rotate_PDFs_Data()

    def SelectPDF_Merge(self):
        selected_files = filedialog.askopenfilenames(initialdir="/",
                                                     title="Selecione um arquivo PDF", filetypes=(("PDF files", "*.pdf*"),))
        if selected_files:
            for path in selected_files:
                self.PDF_List.insert(END, path)

    def Select_Directory(self):
        self.saving_location = filedialog.askdirectory(title="Selecione um local")
        self.Update_Path_Label()

    # Tela para dividir os PDFs
    def Split_PDF_Data(self):
        pdfReader = PdfReader(self.PDF_path)
        total_pages = len(pdfReader.pages)

        self.ClearScreen()
        home_btn = Button(self.frame_1, text="Inicio",
                          font=(self.font_1, 10, 'bold'), command=self.Home_Page)
        home_btn.place(x=10, y=10)

        header = Label(self.frame_1, text="Dividir PDF",
                       font=(self.font_3, 25, "bold"), bg=self.color_2, fg=self.color_1)
        header.place(x=265, y=10)

        self.pages_label = Label(self.frame_1,
                                 text=f"Número Total de Páginas: {total_pages}",
                                 font=(self.font_2, 15, 'bold'), bg=self.color_2, fg=self.color_3)
        self.pages_label.place(x=40, y=70)

        From = Label(self.frame_1, text="De",
                     font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_1)
        From.place(x=40, y=120)

        self.From_Entry = Entry(self.frame_1, font=(self.font_2, 12, 'bold'),
                                width=8)
        self.From_Entry.place(x=40, y=150)

        To = Label(self.frame_1, text="Para", font=(self.font_2, 16, 'bold'),
                   bg=self.color_2, fg=self.color_1)
        To.place(x=160, y=120)

        self.To_Entry = Entry(self.frame_1, font=(self.font_2, 12, 'bold'),
                              width=8)
        self.To_Entry.place(x=160, y=150)

        Cur_Directory = Label(self.frame_1, text="Local de Armazenamento",
                              font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_1)
        Cur_Directory.place(x=40, y=200)

        self.path_label = Label(self.frame_1, text='/',
                                font=(self.font_2, 15, 'bold'), bg=self.color_2, fg=self.color_3)
        self.path_label.place(x=40, y=230)

        select_loc_btn = Button(self.frame_1, text="Selecionar local",
                                font=(self.font_1, 8, 'bold'), command=self.Select_Directory)
        select_loc_btn.place(x=40, y=260)

        split_button = Button(self.frame_1, text="Dividir",
                              font=(self.font_3, 15, 'bold'), bg=self.color_4, fg=self.color_1,
                              width=12, command=self.Split_PDF)
        split_button.place(x=280, y=390)

    # Tela para mesclar os PDFs
    def Merge_PDFs_Data(self):
        self.ClearScreen()
        home_btn = Button(self.frame_1, text="Inicio",
                          font=(self.font_1, 10, 'bold'), command=self.Home_Page)
        home_btn.place(x=10, y=10)

        header = Label(self.frame_1, text="Mesclar PDFs",
                       font=(self.font_3, 25, "bold"), bg=self.color_2, fg=self.color_1)
        header.place(x=265, y=15)

        select_pdf_label = Label(self.frame_1, text="Selecione os PDFs",
                                 font=(self.font_2, 15, 'bold'), bg=self.color_2, fg=self.color_3)
        select_pdf_label.place(x=40, y=70)

        open_button = Button(self.frame_1, text="Abrir pasta",
                             font=(self.font_1, 9, 'bold'), command=self.SelectPDF_Merge)
        open_button.place(x=40, y=100)

        Cur_Directory = Label(self.frame_1, text="Local de Armazenamento",
                              font=(self.font_2, 18, 'bold'), bg=self.color_2, fg=self.color_1)
        Cur_Directory.place(x=40, y=280)

        self.path_label = Label(self.frame_1, text='/',
                                font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_3)
        self.path_label.place(x=40, y=310)

        select_loc_btn = Button(self.frame_1, text="Selecionar local",
                                font=(self.font_1, 8, 'bold'), command=self.Select_Directory)
        select_loc_btn.place(x=40, y=340)

        self.PDF_List = Listbox(self.frame_1, font=(self.font_2, 10, 'bold'),
                                bg=self.color_2, fg=self.color_1,
                                selectbackground=self.color_1, selectmode=MULTIPLE)
        self.PDF_List.place(x=40, y=130, width=320, height=140)

        scrollbar = Scrollbar(self.PDF_List, orient="vertical")
        scrollbar.config(command=self.PDF_List.yview)
        scrollbar.pack(side="right", fill="y")

        merge_button = Button(self.frame_1, text="Mesclar",
                              font=(self.font_3, 15, 'bold'), bg=self.color_4, fg=self.color_1,
                              width=12, command=self.Merge_PDF)
        merge_button.place(x=380, y=390)

        delete_button = Button(self.frame_1, text="Excluir",
                               font=(self.font_3, 15, 'bold'), bg=self.color_4, fg=self.color_1,
                               width=12, command=self.delete_list_items)
        delete_button.place(x=180, y=390)

    # Tela para Girar os PDFs
    def Rotate_PDFs_Data(self):
        self.ClearScreen()
        home_btn = Button(self.frame_1, text="Inicio",
                          font=(self.font_1, 10, 'bold'), command=self.Home_Page)
        home_btn.place(x=10, y=10)

        header = Label(self.frame_1, text="Girar PDF",
                       font=(self.font_3, 25, "bold"), bg=self.color_2, fg=self.color_1)
        header.place(x=265, y=15)

        Cur_Directory = Label(self.frame_1, text="Local de Armazenamento",
                              font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_1)
        Cur_Directory.place(x=40, y=70)

        self.path_label = Label(self.frame_1, text='/',
                                font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_3)
        self.path_label.place(x=40, y=100)

        select_loc_btn = Button(self.frame_1, text="Selecionar Local",
                                font=(self.font_1, 8, 'bold'), command=self.Select_Directory)
        select_loc_btn.place(x=40, y=130)

        select_degree = Label(self.frame_1, text="Selecione o grau",
                              font=(self.font_2, 16, 'bold'), bg=self.color_2, fg=self.color_1)
        select_degree.place(x=40, y=180)

        self.degree_combo = ttk.Combobox(self.frame_1, font=(self.font_2, 14, 'bold'),
                                         values=["90", "180", "270"], width=8)
        self.degree_combo.place(x=40, y=210)

        rotate_button = Button(self.frame_1, text="Girar",
                               font=(self.font_3, 15, 'bold'), bg=self.color_4, fg=self.color_1,
                               width=12, command=self.Rotate_PDF)
        rotate_button.place(x=280, y=390)

    def delete_list_items(self):
        selected_items = self.PDF_List.curselection()
        for index in selected_items[::-1]:
            self.PDF_List.delete(index)

    # Operation Functions
    # Salavando os arquivos: Dividido
    def Split_PDF(self):
        try:
            start = int(self.From_Entry.get())
            end = int(self.To_Entry.get())
            pdfReader = PdfReader(self.PDF_path)
            if start < 1 or end > len(pdfReader.pages) or start > end:
                raise ValueError

        except ValueError:
            messagebox.showerror("Error", "Número de Página Inválido")
            return

        if not self.saving_location:
            messagebox.showerror("Error", "Selecione Local de Armazenamento")
            return

        pdf_writer = PdfWriter()

        for page_num in range(start - 1, end):
            pdf_writer.add_page(pdfReader.pages[page_num])

        output_file_path = os.path.join(self.saving_location, "split_document.pdf")
        with open(output_file_path, 'wb') as output_file:
            pdf_writer.write(output_file)

        messagebox.showinfo("Atenção", "PDF Dividido com sucesso!")

    # Salavando os arquivos: Mesclados
    def Merge_PDF(self):
        try:
            if not self.saving_location:
                messagebox.showerror("Error", "Selecione Local de Armazenamento")
                return

            pdf_writer = PdfWriter()
            for idx in range(self.PDF_List.size()):
                pdf_path = self.PDF_List.get(idx)
                pdf_reader = PdfReader(pdf_path)
                for page_num in range(len(pdf_reader.pages)):
                    pdf_writer.add_page(pdf_reader.pages[page_num])

            output_file_path = os.path.join(self.saving_location, "merged_document.pdf")
            with open(output_file_path, 'wb') as output_file:
                pdf_writer.write(output_file)

            messagebox.showinfo("Info", "PDFs Mesclados com sucesso!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while merging PDFs: {e}")
            print(f"Error in Merge_PDF: {e}")

    # Salavando os arquivos: Girados
    def Rotate_PDF(self):
        try:
            degree = int(self.degree_combo.get())
            if degree not in [90, 180, 270]:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Selecione os graus")
            return

        if not self.saving_location:
            messagebox.showerror("Error", "Selecione Local de Armazenamento")
            return

        try:
            pdf_reader = PdfReader(self.PDF_path)
            pdf_writer = PdfWriter()

            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                page.rotate(degree)  # Substitua rotate_clockwise por rotate
                pdf_writer.add_page(page)

            output_file_path = os.path.join(self.saving_location, "rotated_document.pdf")
            with open(output_file_path, 'wb') as output_file:
                pdf_writer.write(output_file)

            messagebox.showinfo("Info", "PDF Girado com Sucesso!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while rotating the PDF: {e}")
            print(f"Error in Rotate_PDF: {e}")
            
if __name__ == "__main__":
    root = Tk()
    obj = PDF_Editor(root)
    root.mainloop()
