from tkinter import *
from hospital_data_scrap import scrap_data
import datetime

def get_current_date():
    current_date = datetime.datetime.today().strftime('%Y-%m-%d')
    return str(current_date)

class Gui_Window(Tk):
    def __init__(self):
        super().__init__()
        self.title('手术数据批量爬取')
        self.geometry('500x400')
        self.main_frame = Frame(self, width = 500, height = 400, bg= 'lightgrey')
        # self.main_frame.pack()
        # label = Label(self.main_frame, text = 'hello world')
        # label.pack(side = TOP)
        
        Label(self.master, text = '开始日期：').grid(row = 0, column = 0, padx = 5, pady =5)
        Label(self.master, text = '结束日期：').grid(row = 1, column = 0, padx = 5, pady =5)
        Label(self.master, text = '存储路径：').grid(row = 2, column = 0, padx = 5, pady =5)
        Label(self.master, text = '当前进度：').grid(row = 3, column = 0, padx = 5, pady =5)
        Label(self.master, text = '预计时间：').grid(row = 4, column = 0, padx = 5, pady =5)


        v1 = StringVar(self.master, '2019-01-01')
        v2 = StringVar(self.master, get_current_date())
        v3 = StringVar(self.master, 'D:\\code\\data.xls')
        v4 = StringVar(self.master)
        v5 = StringVar(self.master)


        self.start_date_entry = Entry(self.master, textvariable = v1)
        self.end_date_entry = Entry(self.master, textvariable = v2 )
        self.stored_dir_entry = Entry(self.master, textvariable = v3)
        self.current_progress_entry = Entry(self.master, textvariable = v4 )
        self.remaining_time_entry = Entry(self.master, textvariable = v5)

        self.start_date_entry.grid(row = 0, column = 1)
        self.end_date_entry.grid(row = 1, column = 1)
        self.stored_dir_entry.grid(row = 2, column = 1)
        self.current_progress_entry.grid(row = 3, column = 1)
        self.remaining_time_entry.grid(row = 4, column = 1)

        self.start_button = Button(self.master, text="开始", height = 2, width = 10, command=self.start_query)
        self.start_button.grid(row = 5, column = 0, padx = 30, pady = 30)
        self.stop_button = Button(self.master, text="停止", height = 2, width = 10, command=self.stop_query)
        self.stop_button.grid(row = 5, column = 1, padx = 30, pady = 30)
        # self.start_button.pack(side=BOTTOM, fill=X)

    def start_query(self):
        scrap_data(self.start_date_entry.get(), self.end_date_entry.get(), self.stored_dir_entry.get())

    def stop_query(self):
        pass

    def estimate_remaing_time(self):
        pass

    def publish_current_date(self):
        pass


if __name__ == "__main__":
    gui_window = Gui_Window()
    gui_window.mainloop()