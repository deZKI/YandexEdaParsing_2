import threading
from tkinter import Frame, Button, Entry, OptionMenu, StringVar, Label, Checkbutton, IntVar, Tk
from core import start_search


ADDRESSES = {
    'Москва – Красная площадь': 'qnolu',
    'Москва – Проспект мира': 'pocwr',
    'Cанкт-Петербург – Невский': 'rpvqe',
}

CATEGORIES = {
    'Красота и гигиена': '3374',
    'Для детей':'21575',
}
class Example(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent, background='white')
        self.parent = parent
        self.clickedAddress = StringVar()
        self.clickedCategory = StringVar()
        self.clickedAddress.set('Москва – Красная площадь')
        self.clickedCategory.set('Красота и гигиена')
        self.start_search = start_search
        self.create_topic_menu()

    def create_topic_menu(self):
        self.text_address = Label(text='Адрес', font="ARIAL 15")
        self.text_address.place(x=0, y=0, width=350)

        self.address = OptionMenu(self.parent, self.clickedAddress, *ADDRESSES.keys())
        self.address.place(x=0, y=24, width=350, height=30)

        self.text_category = Label(text='Категория', font="ARIAL 15")
        self.text_category.place(x=0, y=50, width=350)

        self.category = OptionMenu(self.parent, self.clickedCategory, *CATEGORIES.keys())
        self.category.place(x=0, y=80, width=350, height=30)

        self.text_font_size = Label(text='Размер шрифта', font="ARIAL 15")
        self.text_font_size.place(x=0, y=110, width=350)

        self.font_size = Entry()
        self.font_size.place(x=0, y=130, width=350, height=35)

        self.text_root = Label(text='Выберите папку', font="ARIAL 15")
        self.text_root.place(x=0, y=180, width=350)

        self.root = Entry()
        self.root.place(x=0, y=210, width=350, height=35)


        self.btn_create_excel = Button(text='Найти', command=self.create_excel)
        self.btn_create_excel.place(x=0, y=250, height=50, width=350)


    def create_excel(self):
        a = threading.Thread(target=start_search,
                             args=(
                                 ADDRESSES[self.clickedAddress.get()],
                                CATEGORIES[self.clickedCategory.get()],
                                 self.font_size.get(),
                                 self.root.get()
                        ), daemon=True)
        a.start()


def startapp():
    # get_markers()
    root = Tk()
    root.title('ANTI5.com')
    root.geometry("350x310+200+200")
    root.resizable(False, False)
    app = Example(root)
    root.mainloop()
