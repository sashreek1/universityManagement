from read_odf import scan_ods
from pyexcel_ods import save_data
from collections import OrderedDict
from tkinter import *

class app:
    def __init__(self):
        self.list_order = []
        self.result_dict = {}
        self.total_scr_lst = []
    def collect_data(self,data_folder):
        self.result_dict = scan_ods(data_folder)
        for dic in self.result_dict:
            total = 0
            for key in self.result_dict[dic]:
                if "academ" in key:
                    total += self.result_dict[dic][key]
            self.result_dict[dic]['ovr_academics'] = (total/4)*0.4
            self.result_dict[dic]['IELTS'] = (self.result_dict[dic]['IELTS'] / 9) * 30
            self.result_dict[dic]['interview'] = (self.result_dict[dic]['interview']) * 3
            self.result_dict[dic]['total_score'] =round(self.result_dict[dic]['ovr_academics']+self.result_dict[dic]['IELTS']+ self.result_dict[dic]['interview'],2)
            temp_tup = (dic ,self.result_dict[dic]['total_score'])
            self.total_scr_lst.append(temp_tup)

    def sorter(self,lst):
        final_lst = []
        num_list = []
        for i in lst:
            num = i[1]
            num_list.append(num)
        num_list = sorted(num_list)
        for i in num_list:
            for j in lst:
                if i == j[1]:
                    final_lst.append(j)
                    break
        return final_lst

    def main (self,final_path):

        main_list = []
        data = OrderedDict()
        self.list_order = (self.sorter(self.total_scr_lst))
        for j in range(len(self.list_order)):
            tup1 = self.list_order[j]
            main_list.append(list(tup1))

        main_list = list(reversed(main_list))
        main_list.insert(0,['Name','overall score (100)'])
        data.update({"Sheet 1": main_list})
        string = final_path+"/result.ods"
        print(string)
        save_data(string, data)

    def sub(self,data_path,final_path):
        self.collect_data(data_path)
        self.main(final_path)

    def gui(self):
        root = Tk()
        root.geometry("600x200+700+400")
        root.resizable(False, False)
        root.title("University Management")

        l = Label(root, text="Enter the path to the folder containing the ods files ")
        l.place(x=20, y=20)
        l.config(font=26)

        t = Entry(width=70)
        t.place(x=20, y=50)

        l = Label(root, text="Enter the path to the folder to place the result file ")
        l.place(x=20, y=90)
        l.config(font=26)

        u = Entry(width=70)
        u.place(x=20, y=120)

        b = Button(text="Submit", command=lambda: self.sub(t.get(), u.get()))
        b.place(x=270, y=150)

        root.mainloop()


app = app()
app.gui()
