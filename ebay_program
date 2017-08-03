from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter import ttk
from PIL import Image, ImageTk
import nri_ebay
import bs4, xlsxwriter, xlrd
import time
from ebaysdk.finding import Connection as Finding

class Window(Frame):
    
    def __init__(self, master = None):
        Frame.__init__(self, master)
        self.master = master
        self.init_window()
       
        
    def init_window(self):
        self.master.title("EBAY SCANNER")
        self.pack(fill=BOTH, expand=1)
        menu = Menu(self.master)
        self.master.config(menu=menu)
        
        file = Menu(menu)
        file.add_command(label='Exit', command=self.client_exit)
        menu.add_cascade(label='File', menu=file)
    
    
    def home_page(self):
        run_button = Button(root, height = 1, width = 12,text='Run scan',command=self.run_scan)
        run_button.place(x=3, y=0)
    
        optionlist = ['New', 'Preset']
        self.setting = ttk.Combobox(self, width = 8)
        self.setting.place(x=100, y=0)
        self.setting.set(optionlist[0])
        self.setting['values'] = (optionlist)
              
        
    def run_scan(self):
        filename = (str(time.strftime('%d-%m-%Y')) + '_ebay_scrape' + '.xlsx')
        self.workbook = nri_ebay.Workbook(filename)
        variable = self.setting.get()
        if variable == 'Preset':
            self.keywords = self.import_xlsx()
            self.filter(self.keywords)
        else:
            self.new_query()
            
            
    def filter(self, keywords, filters=''):
        if not filters:
            filters = ['substitute', 'used', 'old']
        nri_ebay.ebay_search(keywords, filters, self.workbook, root)
        self.workbook.close()
    
    
    def file_select(self):
        Tk().withdraw() 
        filename = askopenfilename()
        return filename
        
        
    def import_xlsx(self):
        filename = self.file_select()
        workbook = xlrd.open_workbook(filename)
        worksheet = workbook.sheet_by_index(0)
        
        num_rows = worksheet.nrows
        curr_row = 1
        row_array = {}
        
        while curr_row < num_rows:
            remote = worksheet.cell_value(curr_row, 0)
            x = worksheet.cell_value(curr_row, 1)
            y = worksheet.cell_value(curr_row, 2)
            row_array[remote] = {'price_param':float(x), 'stock_param':int(y)} 
            curr_row += 1
        self.row_array = row_array
        return row_array
        
        
    def new_query(self):
        self.row_array = {}
        window = Tk()
        window.geometry('500x500')
        window = Window(window)
        
        text = Label(window, text='Remote Part')
        text.place(x=25, y=0)
        text2 = Label(window, text='Price Limit')
        text2.place(x=188, y=0)
        text3 = Label(window, text='Stock Limit')
        text3.place(x=350, y=0)
        self.formcell = Entry(window)
        self.formcell.place(x=25, y=30)
        self.formcell2 = Entry(window)
        self.formcell2.place(x=188, y =30)
        self.formcell3 = Entry(window)
        self.formcell3.place(x=350, y =30)
        
        query_button = Button(window, height = 1, width = 9,text='Add/Update',command=self.query_grab)
        query_button.place(x=25, y=60)
        remove_button = Button(window, height = 1, width = 9,text='Remove',command=self.remove_query)
        remove_button.place(x=188, y=60)
        done_button = Button(window, height = 1, width = 9,text='Done',command=self.query_results)
        done_button.place(x=350, y=60)
        self.T = Text(window, state=DISABLED, height=23, width=55)
        self.T.place(x=25, y=100)
        self.T.insert(END, self.row_array)

        
    def remove_query(self):
        self.part = self.formcell.get()
        del self.row_array[self.part]
        self.T['state'] = NORMAL
        self.T.delete("1.0", END)
        for row in self.row_array:
            self.T.insert(END, row + '    price:' + str(self.row_array[row]['price_param']) + '    stock:' + str(self.row_array[row]['stock_param']) + '\n')
        self.T['state'] = DISABLED
        
        
    def query_grab(self):
        self.part = self.formcell.get()
        self.price = self.formcell2.get()
        self.stock = self.formcell3.get()
        self.row_array[self.part] = {'price_param':float(self.price), 'stock_param':int(self.stock)}
        self.T['state'] = NORMAL
        self.T.delete("1.0", END)
        for row in self.row_array:
            self.T.insert(END, row + '    price:' + str(self.row_array[row]['price_param']) + '    stock:' + str(self.row_array[row]['stock_param']) + '\n')
        self.T['state'] = DISABLED

    
    def query_results(self):
        row_array = self.row_array
        self.save_settings()
        self.filter(row_array)
    
    
    def save_settings(self):
        row_array = self.row_array
        filename =  str(time.strftime('%d-%m-%Y')) + '_preset' + '.xlsx'
        filename = 'C:/Users/dougl/Desktop/EBAY BOT/Presets/' + filename
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()
        cell = 0
        worksheet.write(cell, 0, 'item')
        worksheet.write(cell, 1, 'price_param')
        worksheet.write(cell, 2, 'stock_param')
        for item in row_array.keys():
            cell += 1
            worksheet.write(cell, 0, item)
            worksheet.write(cell, 1, row_array[item]['price_param'])
            worksheet.write(cell, 2, row_array[item]['stock_param'])
    
    
    def client_exit(self):
        exit()
    
    
if __name__ == "__main__":
    root = Tk()
    root.geometry('178x60')
    app = Window(root)
    app.home_page()
    root.mainloop()
    
