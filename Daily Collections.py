import tkinter as tk
from tkinter import *
import Excel
import tkinter.messagebox
from tkinter import ttk
import time
from datetime import datetime
from Mail import mail
from Excel import *


root = Tk()
client = list()
update = list()
tk_var = StringVar(root)
date_var = IntVar(root)
month_var = IntVar(root)
year_var = IntVar(root)
fup_date_var = IntVar(root)
fup_month_var = IntVar(root)
fup_year_var = IntVar(root)
pre_date_var = IntVar(root)
pre_month_var = IntVar(root)
pre_year_var = IntVar(root)
choices = ['--None--', 'Yearly', 'Half-Yearly', 'Quarterly', 'Monthly']
date = range(1, 32)
month = range(1, 13)
year = range(1950, 3000)
excel = Excel.excel()

class Client():

    def clear_frame(self):
        for widget in right_frame.winfo_children():
            widget.destroy()

    def change_dropdown(self, *args):
        global tk_var, date_var, month_var, year_var, fup_year_var, fup_month_var, fup_date_var, pre_year_var, pre_month_var, pre_date_var
        try:
            self.period = tk_var.get()
            self.date = date_var.get()
            self.month = month_var.get()
            self.year = year_var.get()
            self.fup_date = fup_date_var.get()
            self.fup_month = fup_month_var.get()
            self.fup_year = fup_year_var.get()
            self.pre_date = pre_date_var.get()
            self.pre_month = pre_month_var.get()
            self.pre_year = pre_year_var.get()
        except Exception as e:
            tkinter.messagebox.showerror("Error", e)

    def add_customer_blank(self):
        global tk_var, choices
        self.clear_frame()
        self.destroy_on_reset()
        self.window = Frame(right_frame, height=500, width=500, borderwidth=2, relief=SOLID)
        self.window.grid(row=0, column=0, sticky=NE)
        self.window.grid_propagate(0)
        Label(self.window, text="NAME").grid(row=0, sticky=W)
        self.name_entry = Entry(self.window)
        self.name_entry.grid(row=0, column=1)
        Label(self.window, text="ADDRESS").grid(row=1, sticky=W)
        self.address_entry = Entry(self.window)
        self.address_entry.grid(row=1, column=1)
        Label(self.window, text="Mobile Number").grid(row=2, sticky=W)
        self.mobile_entry = Entry(self.window)
        self.mobile_entry.grid(row=2, column=1)
        Label(self.window, text="Policy Number").grid(row=3, sticky=W)
        self.policy_entry = Entry(self.window)
        self.policy_entry.grid(row=3, column=1)
        Label(self.window, text="DOC").grid(row=4, sticky=W)
        date_var.set('1')
        month_var.set('1')
        year_var.set('1950')
        year_menu = OptionMenu(self.window, year_var, *year)
        month_menu = OptionMenu(self.window, month_var, *month)
        date_menu = OptionMenu(self.window, date_var, *date)
        date_menu.grid(row=4, column=1)
        date_var.trace('w', self.change_dropdown)
        month_menu.grid(row=4, column=2)
        month_var.trace('w', self.change_dropdown)
        year_menu.grid(row=4, column=3)
        year_var.trace('w', self.change_dropdown)
        Label(self.window, text="FUP").grid(row=5, sticky=W)
        fup_date_var.set('1')
        fup_month_var.set('1')
        fup_year_var.set('1950')
        fup_year_menu = OptionMenu(self.window, fup_year_var, *year)
        fup_month_menu = OptionMenu(self.window, fup_month_var, *month)
        fup_date_menu = OptionMenu(self.window, fup_date_var, *date)
        fup_date_menu.grid(row=5, column=1)
        fup_date_var.trace('w', self.change_dropdown)
        fup_month_menu.grid(row=5, column=2)
        fup_month_var.trace('w', self.change_dropdown)
        fup_year_menu.grid(row=5, column=3)
        fup_year_var.trace('w', self.change_dropdown)
        Label(self.window, text="Premium").grid(row=6, sticky=W)
        self.premium_entry = Entry(self.window)
        self.premium_entry.grid(row=6, column=1)
        tk_var.set('--None--')
        popupMenu = OptionMenu(self.window, tk_var, *choices)
        Label(self.window, text="Period").grid(row=7, sticky=W)
        popupMenu.grid(row=7, column=1)
        tk_var.trace('w', self.change_dropdown)
        Label(self.window, text="Target").grid(row=10, sticky=W)
        self.target_entry = Entry(self.window)
        self.target_entry.grid(row=10, column=1)
        Button(self.window, text="Submit", command=self.add_to_excel, height=1, width=15, font=(None, 10), relief=RAISED).grid(row=30, columnspan=5)
        Button(self.window, text="Reset", command=self.add_customer_blank, height=1, width=15, font=(None, 10), relief=RAISED).grid(row=31, columnspan=5)

    def add_to_excel(self):
        global client
        Name = self.name_entry.get()
        if Name == "":
            try:
                self.Name_Label
            except:
                self.Name_Label = Label(self.window, text="  Please fill Name", foreground="red", font=(None, 10))
                self.Name_Label.grid(row=0, column=2)
        else:
            try:
                self.Name_Label.destroy()
                del self.Name_Label
            except Exception as e:
                pass
            client.append(Name)
        Address = self.address_entry.get()
        if Address == "":
            try:
                self.Address_Label
            except:
                self.Address_Label = Label(self.window, text="  Please fill Address", foreground="red", font=(None, 10))
                self.Address_Label.grid(row=1, column=2)
        else:
            try:
                self.Address_Label.destroy()
                del self.Address_Label
            except:
                pass
            client.append(Address)
        Mobile = self.mobile_entry.get()
        if Mobile == "":
            try:
                self.Mobile_Label
            except:
                self.Mobile_Label = Label(self.window, text="  Please fill Mobile Number", foreground="red", font=(None, 10))
                self.Mobile_Label.grid(row=2, column=2)
        else:
            try:
                self.Mobile_Label.destroy()
                del self.Mobile_Label
            except:
                pass
            client.append(Mobile)
        Policy = self.policy_entry.get()
        if Policy == "":
            try:
                self.Policy_Label
            except:
                self.Policy_Label = Label(self.window, text="  Please fill Policy Number", foreground="red", font=(None, 10))
                self.Policy_Label.grid(row=3, column=2)
        else:
            try:
                self.Policy_Label.destroy()
                del self.Policy_Label
            except Exception as e:
                pass
            client.append(Policy)
        try:
            doc = str(self.date) + "-" + str(self.month) + "-" + str(self.year)
            client.append(doc)
        except:
            doc = str(date_var.get()) + "-" + str(month_var.get()) + "-" + str(year_var.get())
            client.append(doc)
        try:
            fup = str(self.fup_date) + "-" + str(self.fup_month) + "-" + str(self.fup_year)
            client.append(fup)
        except:
            fup = str(fup_date_var.get()) + "-" + str(fup_month_var.get()) + "-" + str(fup_year_var.get())
            client.append(fup)
        try:
            Period = self.period
            if Period == "--None--":
                try:
                    self.Period_Label
                except:
                    self.Period_Label = Label(self.window, text="  Please select Time Period", foreground="red", font=(None, 10))
                    self.Period_Label.grid(row=7, column=2)
            else:
                try:
                    self.Period_Label.destroy()
                    del self.Period_Label
                except:
                    pass
                client.append(Period)
        except:
            Period = tk_var.get()
            if Period == "--None--":
                try:
                    self.Period_Label
                except:
                    self.Period_Label = Label(self.window, text="  Please select Time Period", foreground="red", font=(None, 10))
                    self.Period_Label.grid(row=7, column=2)
            else:
                try:
                    self.Period_Label.destroy()
                    del self.Period_Label
                except:
                    pass
                client.append(Period)
        Premium = self.premium_entry.get()
        if Premium == "":
            try:
                self.Premium_Label
            except:
                self.Premium_Label = Label(self.window, text="  Please fill Premium", foreground="red", font=(None, 10))
                self.Premium_Label.grid(row=6, column=2)
        else:
            try:
                self.Premium_Label.destroy()
                del self.Premium_Label
            except:
                pass
            try:
                if int(Premium):
                    client.append(Premium)
            except ValueError:
                self.Premium_Label = Label(self.window, text="  Please fill Premium correctly", foreground="red",
                                           font=(None, 10))
                self.Premium_Label.grid(row=6, column=2)
        Target = self.target_entry.get()
        if Target == "":
            try:
                self.Target_Label
            except:
                self.Target_Label = Label(self.window, text="  Please fill Target", foreground="red", font=(None, 10))
                self.Target_Label.grid(row=10, column=2)
        else:
            try:
                self.Target_Label.destroy()
                del self.Target_Label
            except:
                pass
            client.append(Target)
        if len(client) == 9:
            submit = tkinter.messagebox.askquestion("Submit Entry", "Are you sure you want to submit?")
            if submit == "yes":
                try:
                    msg = excel.add_client(client)
                    if msg == 1:
                        tkinter.messagebox.showinfo("Message", "Customer added successfully")
                    elif msg == 0:
                        tkinter.messagebox.showerror("Error", "Could not add customer. Try again")
                    else:
                        tkinter.messagebox.showerror("Error", msg)
                    self.add_customer_blank()
                except PermissionError:
                    tkinter.messagebox.showerror("Error", "Closing the application \n Please close the Excel file")
                    time.sleep(15)
                    root.quit()
        else:
            tkinter.messagebox.showerror("Error", "Error submitting the form. Please resolve all the errors")
        del client[:]

    def update_customer_blank(self):
        global choices
        try:
            self.update_window = Toplevel(right_frame, bd=2)
            self.update_window.minsize(500, 500)
            Label(self.update_window, text="NAME").grid(row=0, sticky=W)
            self.update_name = Entry(self.update_window)
            self.update_name.insert(0, self.entry['values'][0])
            self.update_name.grid(row=0, column=1)
            Label(self.update_window, text="ADDRESS").grid(row=1, sticky=W)
            self.update_address = Entry(self.update_window)
            self.update_address.insert(0, self.entry['values'][1])
            self.update_address.grid(row=1, column=1)
            Label(self.update_window, text="Mobile Number  ").grid(row=2, sticky=W)
            self.update_mobile = Entry(self.update_window)
            self.update_mobile.insert(0, self.entry['values'][2])
            self.update_mobile.grid(row=2, column=1)
            Label(self.update_window, text="Policy Number").grid(row=3, sticky=W)
            self.update_policy = Entry(self.update_window)
            self.update_policy.insert(0, self.entry['values'][3])
            self.update_policy.grid(row=3, column=1)
            self.update_policy.configure(state='disabled')
            Label(self.update_window, text="DOC").grid(row=4, sticky=W)
            doc_split = self.entry['values'][4].split('-')
            date_var.set(doc_split[0])
            month_var.set(doc_split[1])
            year_var.set(doc_split[2])
            year_menu = OptionMenu(self.update_window, year_var, *year)
            month_menu = OptionMenu(self.update_window, month_var, *month)
            date_menu = OptionMenu(self.update_window, date_var, *date)
            date_menu.grid(row=4, column=1)
            date_var.trace('w', self.change_dropdown)
            month_menu.grid(row=4, column=2)
            month_var.trace('w', self.change_dropdown)
            year_menu.grid(row=4, column=3)
            year_var.trace('w', self.change_dropdown)
            Label(self.update_window, text="FUP").grid(row=5, sticky=W)
            fup_split = self.entry['values'][5].split('-')
            fup_date_var.set(fup_split[0])
            fup_month_var.set(fup_split[1])
            fup_year_var.set(fup_split[2])
            fup_year_menu = OptionMenu(self.update_window, fup_year_var, *year)
            fup_month_menu = OptionMenu(self.update_window, fup_month_var, *month)
            fup_date_menu = OptionMenu(self.update_window, fup_date_var, *date)
            fup_date_menu.grid(row=5, column=1)
            fup_date_var.trace('w', self.change_dropdown)
            fup_month_menu.grid(row=5, column=2)
            fup_month_var.trace('w', self.change_dropdown)
            fup_year_menu.grid(row=5, column=3)
            fup_year_var.trace('w', self.change_dropdown)
            Label(self.update_window, text="Premium").grid(row=6, sticky=W)
            self.update_premium_entry = Entry(self.update_window)
            self.update_premium_entry.insert(0, self.entry['values'][7])
            self.update_premium_entry.grid(row=6, column=1)
            tk_var.set(self.entry['values'][6])
            menu = OptionMenu(self.update_window, tk_var, *choices)
            Label(self.update_window, text="Choose the Time Period of due").grid(row=7, sticky=W)
            menu.grid(row=7, column=1)
            tk_var.trace('w', self.change_dropdown)
            Label(self.update_window, text="Daily Collection").grid(row=8, sticky=W)
            self.update_daily_collection_entry = Entry(self.update_window)
            self.update_daily_collection_entry.insert(0, self.entry['values'][8])
            self.update_daily_collection_entry.configure(state='disabled')
            self.update_daily_collection_entry.grid(row=8, column=1)
            Label(self.update_window, text="Target").grid(row=10, sticky=W)
            self.update_target_entry = Entry(self.update_window)
            self.update_target_entry.insert(0, self.entry['values'][9])
            self.update_target_entry.grid(row=10, column=1)
            Button(self.update_window, text="Submit", command=self.update_to_excel, height=1, width=15, font=(None, 10),
                            relief=RAISED).grid(row=30, columnspan=5)
        except AttributeError:
            tkinter.messagebox.showerror("Error", "Please select an entry and try again")
            self.update_window.destroy()

    def update_to_excel(self):
        global update
        Name = self.update_name.get()
        update.append(Name)
        Address = self.update_address.get()
        update.append(Address)
        Mobile = self.update_mobile.get()
        update.append(Mobile)
        Policy = self.update_policy.get()
        update.append(Policy)
        try:
            doc = str(self.date) + "-" + str(self.month) + "-" + str(self.year)
            update.append(doc)
        except:
            doc = str(date_var.get()) + "-" + str(month_var.get()) + "-" + str(year_var.get())
            update.append(doc)
        try:
            fup = str(self.fup_date) + "-" + str(self.fup_month) + "-" + str(self.fup_year)
            update.append(fup)
        except:
            fup = str(fup_date_var.get()) + "-" + str(fup_month_var.get()) + "-" + str(fup_year_var.get())
            update.append(fup)
        try:
            Period = self.period
            update.append(Period)
        except:
            Period = tk_var.get()
            update.append(Period)
        Premium = self.update_premium_entry.get()
        try:
            if int(Premium):
                update.append(Premium)
        except ValueError:
            tkinter.messagebox.showerror("Error", "Please enter correct Premium")
        Daily_Collection = self.update_daily_collection_entry.get()
        try:
            int(Daily_Collection)
            update.append(Daily_Collection)
        except ValueError:
            tkinter.messagebox.showerror("Error", "Please enter correct Daily Collection value")
        Target = self.update_target_entry.get()
        update.append(Target)
        try:
            if len(update) == 10:
                submit = tkinter.messagebox.askquestion("Submit Entry", "Are you sure you want to submit?")
                if submit == "yes":
                    flag = excel.update_client(update, self.entry["text"])
                    if flag == 0:
                        tkinter.messagebox.showerror("Error", "Could not update customer details. Try again")
            self.view_from_excel()
        except PermissionError:
            tkinter.messagebox.showerror("Error", "Closing the application \n Please close the Excel file")
            root.quit()
        self.update_window.destroy()
        del update[:]

    def delete_from_excel(self):
        try:
            if self.entry:
                submit = tkinter.messagebox.askquestion("Delete Entry", "Are you sure you want to delete {}?".format(self.entry['values'][0]))
                if submit == "yes":
                    flag = excel.delete_client(self.entry['text'])
                    if flag == 0:
                        tkinter.messagebox.showerror("Error", "Could not delete customer. Try again")
                self.view_from_excel()
        except AttributeError:
            tkinter.messagebox.showerror("Error", "Please select an entry and try again")

    def clear_bottom_frame(self):
        for widget in self.right_bottom_frame.winfo_children():
            widget.destroy()
        try:
            self.ysb.destroy()
        except AttributeError:
            pass

    def clear_report_frame(self):
        for widget in self.report_bottom_frame.winfo_children():
            widget.destroy()
        try:
            self.y.destroy()
        except AttributeError:
            pass

    def select_item(self, a):
        self.entry = self.table.item(self.table.selection())

    def view_from_excel(self):
        self.clear_bottom_frame()
        try:
            self.entry
            del self.entry
        except:
            pass
        self.table = ttk.Treeview(self.right_bottom_frame)
        self.table['columns'] = ('Name', 'Address', 'Mobile', 'Policy Number', 'DOC', 'FUP', 'Period', \
                  'Premium', 'Daily Collection', 'Target', 'Balance')

        Full_Name = self.full_name_entry.get()
        Policy_entry = self.policy_number_entry.get()
        self.text = excel.view_client(Full_Name, Policy_entry)
        self.table.column("#0", width=100, minwidth=50, stretch=tk.NO)
        self.table.column("Name", width=100, minwidth=50, stretch=tk.NO)
        self.table.column("Address", width=100, minwidth=50, stretch=tk.NO)
        self.table.column("Mobile", width=100, minwidth=50, stretch=tk.NO)
        self.table.column("Policy Number", width=100, minwidth=50, stretch=tk.NO)
        self.table.column("DOC", width=100, minwidth=50, stretch=tk.NO)
        self.table.column("FUP", width=100, minwidth=50, stretch=tk.NO)
        self.table.column("Premium", width=100, minwidth=50, stretch=tk.NO)
        self.table.column("Period", width=100, minwidth=50, stretch=tk.NO)
        self.table.column("Daily Collection", width=100, minwidth=50, stretch=tk.NO)
        self.table.column("Target", width=100, minwidth=50, stretch=tk.NO)
        self.table.column("Balance", width=100, minwidth=50, stretch=tk.NO)
        ttk.Style().configure("Treeview", font=('', 10), foreground="maroon", background="yellow")
        ttk.Style().configure("Treeview.Heading", font=('', 10), foreground="navyblue")
        self.table.heading('#0', text='S.No', anchor=tk.W)
        self.table.heading('Name', text='Name', anchor=tk.W)
        self.table.heading('Address', text='Address', anchor=tk.W)
        self.table.heading('Mobile', text='Mobile', anchor=tk.W)
        self.table.heading('Policy Number', text='Policy Number', anchor=tk.W)
        self.table.heading('DOC', text='DOC', anchor=tk.W)
        self.table.heading('FUP', text='FUP', anchor=tk.W)
        self.table.heading('Premium', text='Premium', anchor=tk.W)
        self.table.heading('Period', text='Period', anchor=tk.W)
        self.table.heading('Daily Collection', text='Daily Collection', anchor=tk.W)
        self.table.heading('Target', text='Target', anchor=tk.W)
        self.table.heading('Balance', text='Balance', anchor=tk.W)
        self.ysb = ttk.Scrollbar(orient=VERTICAL, command=self.table.yview)
        self.table['yscroll'] = self.ysb.set

        try:
            row = 1
            for element in self.text:
                S_No, Name, Address, Mobile, Policy_Number, DOC, FUP, Period, Premium, Daily_collection, Target, Balance = element[:12]
                self.table.insert("", row, text=S_No, values=(Name, Address, Mobile, Policy_Number, DOC, FUP, Period, Premium, Daily_collection, Target, Balance))
                row += 1
        except ValueError:
            tkinter.messagebox.showerror("Error", "No entry with name {}, Policy {}".format(Full_Name, Policy_entry))

        self.table.bind('<ButtonRelease-1>', self.select_item)
        self.table.grid(in_=self.right_bottom_frame, sticky=NSEW)
        s = ttk.Style()
        s.configure('TButton', font=(None, 10), relief=RAISED, height=1, width=15)
        ttk.Button(self.right_bottom_frame, text="Delete Details", command=self.delete_from_excel, style='TButton').grid(sticky=S)
        ttk.Button(self.right_bottom_frame, text="Update Details", command=self.update_customer_blank).grid(sticky=S)

        self.ysb.grid(in_=self.right_bottom_frame, row=0, column=1, sticky=NS)

    def view_customer_blank(self):
        self.clear_frame()

        right_top_frame = Frame(right_frame, borderwidth=2, relief=SOLID)
        right_top_frame.pack(side=TOP, fill="both")
        self.right_bottom_frame = Frame(right_frame, borderwidth=2)
        self.right_bottom_frame.pack(side=BOTTOM, expand=True, fill="both")
        header = Label(right_top_frame, text="Please enter the Full Name or Policy Number of the customer to be deleted/updated in the text box below")
        header.grid(row=2, columnspan=6)
        Label(right_top_frame, text="FULL NAME").grid(row=5, column=0)
        self.full_name_entry = Entry(right_top_frame)
        self.full_name_entry.grid(row=5, column=1)
        Label(right_top_frame, text="POLICY NUMBER").grid(row=5, column=2)
        self.policy_number_entry = Entry(right_top_frame)
        self.policy_number_entry.grid(row=5, column=3)
        Button(right_top_frame, text="Search Details", command=self.view_from_excel, height=1, width=15,
                               font=(None, 10), relief=RAISED).grid(row=60, column=6)

    def view_report(self):
        self.clear_report_frame()
        policy = self.policy_details.get()
        name = self.name_details.get()
        try:
            daily, name, total = excel.client_report(policy, name)
            report_view = ttk.Treeview(self.report_bottom_frame)
            report_view['columns'] = 'Collection'

            ttk.Style().configure("Treeview", font=('', 10), foreground="maroon", background="yellow")
            ttk.Style().configure("Treeview.Heading", font=('', 10), foreground="navyblue")
            s = ttk.Style()
            s.configure('TLabel', font=(None, 13), height=1, width=15, foreground="navyblue", columnspan=2)
            s.configure('TButton', font=(None, 10), relief=RAISED, height=1, width=15)

            ttk.Label(self.report_bottom_frame, text=name["name"], style='TLabel').grid(row=1)
            report_view.column("#0", width=100, minwidth=50, stretch=tk.NO)
            report_view.column("Collection", width=100, minwidth=50, stretch=tk.NO)
            report_view.heading('#0', text='Date')
            report_view.heading('Collection', text='Collection')
            self.y = ttk.Scrollbar(orient=VERTICAL, command=report_view.yview)
            report_view['yscroll'] = self.y.set

            row = 2
            for i in daily:
                report_view.insert("", row, text=i, values=(daily[i], i))
                row += 1

            report_view.grid(in_=self.report_bottom_frame, sticky=NSEW)
            ttk.Label(self.report_bottom_frame, text="TOTAL: {}".format(total), style='TLabel').grid(row=row, sticky=W)
            #ttk.Label(self.report_bottom_frame, text=total, style='TLabel').grid(row=row, column=1)
            ttk.Button(self.report_bottom_frame, text="Print Report", command=excel.open_report, style='TButton').grid(
                sticky=S)
            self.y.grid(in_=self.report_bottom_frame, row=2, column=1, sticky=NS)

        except ValueError:
            tkinter.messagebox.showerror("Error", "No entry with Name {} and Policy Number {}".format(name, policy))

    def backup_mail(self):
        toaddr = self.email.get()
        if toaddr:
            submit = tkinter.messagebox.askquestion("Message", "Are you sure you want to submit?")
            if submit == "yes":
                ret = mail(toaddr)
                if ret == 1:
                    tkinter.messagebox.showinfo("Message", "Backup taken successfully")
                else:
                    tkinter.messagebox.showerror("Error", "Backup failed. Please check your internet connection")
        else:
            tkinter.messagebox.showerror("Error", "Please enter an E-Mail address")

    def backup_blank(self):
        self.clear_frame()
        backup_top_frame = Frame(right_frame, borderwidth=2, relief=SOLID)
        backup_top_frame.pack(side=TOP, fill="both", expand=True)
        Label(backup_top_frame, text="Please enter the 'TO' Email Address", font=(None, 10)).grid(row=0, column=0)
        Label(backup_top_frame, text="Enter E-Mail Address", font=(None, 10)).grid(row=1, column=0)
        self.email = Entry(backup_top_frame)
        self.email.grid(row=1, column=1)
        self.email.insert(0, "emailaddress@domin.com")
        Button(backup_top_frame, text="Submit", command=self.backup_mail, height=1, width=15,
               font=(None, 10), relief=RAISED).grid(row=2, column=1)

    def report_customer_blank(self):
        self.clear_frame()
        report_top_frame = Frame(right_frame, borderwidth=2, relief=SOLID)
        report_top_frame.pack(side=TOP, fill="both")
        self.report_bottom_frame = Frame(right_frame, relief=SOLID)
        self.report_bottom_frame.pack(side=BOTTOM, fill="both", expand=True)
        Label(report_top_frame, text="Please enter the Policy Number of the customer to view report").grid(row=2, columnspan=6)
        Label(report_top_frame, text="FULL NAME").grid(row=5, column=0)
        self.name_details = Entry(report_top_frame)
        self.name_details.grid(row=5, column=1)
        Label(report_top_frame, text="POLICY NUMBER").grid(row=6, column=0)
        self.policy_details = Entry(report_top_frame)
        self.policy_details.grid(row=6, column=1)
        Button(report_top_frame, text="View Report", command=self.view_report, height=1, width=15,
               font=(None, 10), relief=RAISED).grid(row=60, column=6)

    def pre_report_blank(self):
        self.clear_frame()
        Label(right_frame, text="Choose the previous date below").grid(row=0, column=0)
        Label(right_frame, text="Previous Date").grid(row=5, column=0)
        pre_date_var.set('1')
        pre_month_var.set('1')
        pre_year_var.set('1950')
        pre_year_menu = OptionMenu(right_frame, pre_year_var, *year)
        pre_month_menu = OptionMenu(right_frame, pre_month_var, *month)
        pre_date_menu = OptionMenu(right_frame, pre_date_var, *date)
        pre_date_menu.grid(row=5, column=1)
        pre_date_var.trace('w', self.change_dropdown)
        pre_month_menu.grid(row=5, column=2)
        pre_month_var.trace('w', self.change_dropdown)
        pre_year_menu.grid(row=5, column=3)
        pre_year_var.trace('w', self.change_dropdown)
        Button(right_frame, text="Submit", command=self.report_previous,
               height=1, width=15, font=(None, 10), relief=RAISED).grid(row=6, column=2)

    def report_previous(self):
        self.clear_frame()
        try:
            try:
                date_previous = str(self.pre_date) + "-" + str(self.pre_month) + "-" + str(self.pre_year)
            except:
                date_previous = str(pre_date_var.get()) + "-" + str(pre_month_var.get()) + "-" + str(pre_year_var.get())
            flag = excel.view_previous(date_previous)
            if len(flag) > 2:
                tkinter.messagebox.showerror("Error", flag)
                self.pre_report_blank()
            else:
                pre_frame = Frame(right_frame, relief=SOLID)
                pre_frame.pack(side=TOP, fill="both")
                pre_bottom_frame = Frame(right_frame, relief=SOLID)
                pre_bottom_frame.pack(side=BOTTOM, fill="both", expand=True)
                Label(pre_frame, text="The daily collection for date {} is displayed below".format(date_previous),
                      font=(None, 10)).grid(row=0, columnspan=6)
                pre_view = ttk.Treeview(pre_frame)
                pre_view['columns'] = ('Name', 'Address', 'Policy Number', 'Collection')
                ttk.Style().configure("Treeview", font=('', 10), foreground="maroon", background="yellow")
                ttk.Style().configure("Treeview.Heading", font=('', 10), foreground="navyblue")
                s = ttk.Style()
                s.configure('TLabel', font=(None, 13), height=1, width=15, foreground="navyblue", columnspan=2)
                s.configure('TButton', font=(None, 10), relief=RAISED, height=1, width=15)
                pre_view.column("#0", width=100, minwidth=50, stretch=tk.NO)
                pre_view.column("Name", width=100, minwidth=50, stretch=tk.NO)
                pre_view.column("Address", width=100, minwidth=50, stretch=tk.NO)
                pre_view.column("Policy Number", width=100, minwidth=50, stretch=tk.NO)
                pre_view.column("Collection", width=100, minwidth=50, stretch=tk.NO)
                pre_view.heading('#0', text='S.No')
                pre_view.heading('Name', text='Name')
                pre_view.heading('Address', text='Address')
                pre_view.heading('Policy Number', text='Policy Number')
                pre_view.heading("Collection", text="Collection")
                row = 1
                try:
                    if len(flag) == 2:
                        entries, total = flag
                        for detail in entries:
                            S_No, Name, Address, Policy_Number, Collection = detail
                            pre_view.insert("", row, text=S_No, values=(Name, Address, Policy_Number, Collection))
                            row += 1
                        pre_view.grid(in_=pre_frame, sticky=NSEW)
                        self.yaxis = ttk.Scrollbar(orient=VERTICAL, command=pre_view.yview)
                        pre_view['yscroll'] = self.yaxis.set
                        self.yaxis.grid(in_=pre_frame, row=1, column=1, sticky=NS)
                        Label(pre_bottom_frame, text="Total Collection:   ", font=(None, 15)).grid(row=0, column=0)
                        Label(pre_bottom_frame, text=total, font=(None, 15)).grid(row=0, column=1)
                except TypeError:
                    tkinter.messagebox.showerror("Error", "No customer details present in database")
                    self.pre_report_blank()

        except Exception as e:
                tkinter.messagebox.showerror("Error", e)

    def report_today(self):
        self.clear_frame()
        self.today_frame = Frame(right_frame, relief=SOLID)
        self.today_frame.pack(side=TOP, fill="both")
        today_bottom_frame = Frame(right_frame, relief=SOLID)
        today_bottom_frame.pack(side=BOTTOM, fill="both", expand=True)
        Label(self.today_frame, text="The daily collection for today is displayed below", font=(None, 10)).grid(row=0, columnspan=6)
        today_view = ttk.Treeview(self.today_frame)
        today_view['columns'] = ('Name','Address','Policy Number','Collection')
        ttk.Style().configure("Treeview", font=('', 10), foreground="maroon", background="yellow")
        ttk.Style().configure("Treeview.Heading", font=('', 10), foreground="navyblue")
        s = ttk.Style()
        s.configure('TLabel', font=(None, 13), height=1, width=15, foreground="navyblue", columnspan=2)
        s.configure('TButton', font=(None, 10), relief=RAISED, height=1, width=15)
        today_view.column("#0", width=100, minwidth=50, stretch=tk.NO)
        today_view.column("Name", width=100, minwidth=50, stretch=tk.NO)
        today_view.column("Address", width=100, minwidth=50, stretch=tk.NO)
        today_view.column("Policy Number", width=100, minwidth=50, stretch=tk.NO)
        today_view.column("Collection", width=100, minwidth=50, stretch=tk.NO)
        today_view.heading('#0', text='S.No')
        today_view.heading('Name', text='Name')
        today_view.heading('Address', text='Address')
        today_view.heading('Policy Number', text='Policy Number')
        today_view.heading("Collection", text="Collection")

        try:
            details, total = excel.view_today()
            row = 1
            for detail in details:
                S_No, Name, Address, Policy_Number, Collection = detail
                today_view.insert("", row, text=S_No, values=(Name, Address, Policy_Number, Collection))
                row += 1

            today_view.grid(in_=self.today_frame, sticky=NSEW)
            #ttk.Label(self.today_frame, text="Total", style='TLabel').grid(row=row+1, column=0, sticky=S)
            #ttk.Label(self.today_frame, text=total, style='TLabel').grid(row=row+1, column=1, sticky=S)
            #ttk.Button(self.today_frame, text="Print Report", command=excel.open_report, style='TButton').grid(
                #sticky=S)
            self.ybar = ttk.Scrollbar(orient=VERTICAL, command=today_view.yview)
            today_view['yscroll'] = self.ybar.set
            self.ybar.grid(in_=self.today_frame, row=1, column=1, sticky=NS)

            Label(today_bottom_frame, text="Total Collection:   ", font=(None, 15)).grid(row=0, column=0)
            Label(today_bottom_frame, text=total, font=(None, 15)).grid(row=0, column=1)
        except TypeError:
            tkinter.messagebox.showerror("Error", "No customers present")

    def daily_bind(self, event):
        excel.delete_column()
        self.create_daily_frames()

    def add_daily(self, event):
        frame_no = repr(event.widget).split('.!')
        try:
            row_no = int(re.search("\d{1,}", frame_no[4]).group()) + 1
        except AttributeError:
            row_no = 2
        flag = excel.add_daily_collection(row_no, self.entry_list[row_no-2].get())
        if flag:
            tkinter.messagebox.showinfo("Message", "Balance is reset to Premium again")
        self.create_daily_frames()

    def subtract_daily(self, event):
        frame_no = repr(event.widget).split('.!')
        try:
            row_no = int(re.search("\d{1,}", frame_no[4]).group()) + 1
        except AttributeError:
            row_no = 2
        excel.subtract_daily_collection(row_no, self.entry_list[row_no - 2].get())
        self.create_daily_frames()

    def myfunction(self, event):
        w,h = root.winfo_screenwidth(), root.winfo_screenheight()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"), width=w, height=h)

    def create_daily_frames(self):
        self.clear_frame()
        self.no_of_frames = excel.ws.max_row
        row = 0
        column = 0
        self.entry_list = list()
        #right_daily_frame = Frame(right_frame, borderwidth=0.5, height=150, width=200, relief=SOLID)
        #right_daily_frame.pack(side=LEFT, fill="both",expand=TRUE)
        self.canvas = Canvas(right_frame)
        right_daily_frame = Frame(self.canvas)
        myscrollbar = Scrollbar(right_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=myscrollbar.set)
        myscrollbar.pack(side="right", fill="y")
        xscroll = Scrollbar(right_frame, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(xscrollcommand=xscroll.set)
        xscroll.pack(side="bottom", fill="x")
        self.canvas.pack(side="left")
        self.canvas.create_window((0, 0), window=right_daily_frame, anchor='nw')
        right_daily_frame.bind("<Configure>", self.myfunction)
        for frames in range(1, self.no_of_frames):
            name, address,daily_collection,balance = excel.read_entries(frames)
            frame = Frame(right_daily_frame, borderwidth=9, height=250, width=250, relief=SUNKEN, bg="white")
            frame.grid(row=row,column=column)
            frame.grid_propagate(0)
            column+=1
            if column >= 5:
                row+=1
                column=0
            Label(frame, text="SNo.: ", bg="white").grid(row=0, column=0)
            Label(frame,text=frames, bg="white").grid(row=0,column=1)
            Label(frame, text="Name: ", bg="white").grid(row=1, column=0)
            Label(frame, text="Address: ", bg="white",height=3).grid(row=2, column=0)
            Label(frame, text="Balance: ", bg="white").grid(row=3, column=0)
            Label(frame, text=balance, bg="white").grid(row=3, column=1)
            daily_name = Label(frame,text=name.upper(), bg="white").grid(row=1,column=1)
            daily_policy = Label(frame, text=address, bg="white", wraplength=240, anchor=W, justify=CENTER).grid(row=2, column=1)
            Label(frame, text="Rupees ", bg="white").grid(row=4, column=0)
            self.daily_entry = Entry(frame)
            self.entry_list.append(self.daily_entry)
            self.daily_entry.insert(0, daily_collection)
            self.daily_entry.grid(row=4, column=1)
            Label(frame, text="   ", bg="white").grid(row=5, column=0)
            btn = Button(frame, text="Add", height=1, width=8, relief=RAISED)
            btn.bind("<Button-1>", self.add_daily)
            btn.grid(row=5,column=1)
            del_btn = Button(frame, text="Minus", height=1, width=8, relief=RAISED)
            del_btn.bind("<Button-1>", self.subtract_daily)
            del_btn.grid(row=6, column=1)

    def previous_collection(self):
        try:
            pre = str(self.pre_date) + "-" + str(self.pre_month) + "-" + str(self.pre_year)
            flag = excel.previous_date(pre)
            if flag:
                self.create_daily_frames()
            else:
                tkinter.messagebox.showerror("Error", "Cannot choose future date.")
        except Exception as e:
            tkinter.messagebox.showerror("Error", "Please change the date to the required date")
            #pre = str(pre_date_var.get()) + "-" + str(pre_month_var.get()) + "-" + str(pre_year_var.get())

    def previous_collection_blank(self):
        #pass # date, submit
        self.clear_frame()
        Label(right_frame, text="Choose the previous date below").grid(row=0, column=0)
        Label(right_frame, text="Previous Date").grid(row=5, column=0)
        pre_date_var.set('1')
        pre_month_var.set('1')
        pre_year_var.set('1950')
        pre_year_menu = OptionMenu(right_frame, pre_year_var, *year)
        pre_month_menu = OptionMenu(right_frame, pre_month_var, *month)
        pre_date_menu = OptionMenu(right_frame, pre_date_var, *date)
        pre_date_menu.grid(row=5, column=1)
        pre_date_var.trace('w', self.change_dropdown)
        pre_month_menu.grid(row=5, column=2)
        pre_month_var.trace('w', self.change_dropdown)
        pre_year_menu.grid(row=5, column=3)
        pre_year_var.trace('w', self.change_dropdown)
        Button(right_frame, text="Submit", command=self.previous_collection,
               height=1, width=15, font=(None, 10), relief=RAISED).grid(row=6, column=2)

    def destroy_on_reset(self):
        try:
            self.Name_Label
            self.Name_Label.destroy()
            del self.Name_Label
        except Exception as e:
            pass
            #print(e)

        try:
            self.Address_Label
            self.Address_Label.destroy()
            del self.Address_Label
        except Exception as e:
            pass
            #print(e)

        try:
            self.Mobile_Label
            self.Mobile_Label.destroy()
            del self.Mobile_Label
        except Exception as e:
            pass#print(e)

        try:
            self.Policy_Label
            self.Policy_Label.destroy()
            del self.Policy_Label
        except Exception as e:
            pass#print(e)

        try:
            self.Period_Label
            self.Period_Label.destroy()
            del self.Period_Label
        except Exception as e:
            pass#print(e)

        try:
            self.Premium_Label
            self.Premium_Label.destroy()
            del self.Premium_Label
        except Exception as e:
            pass#print(e)

        try:
            self.Daily_Collection_Label
            self.Daily_Collection_Label.destroy()
            del self.Daily_Collection_Label
        except Exception as e:
            pass#print(e)

        try:
            self.Target_Label
            self.Target_Label.destroy()
            del self.Target_Label
        except Exception as e:
            pass#print(e)


root.wm_iconbitmap('logo6.ico')
root.title("Daily Collections")
w, h = root.winfo_screenwidth(), root.winfo_screenheight()
root.geometry("%dx%d+0+0" % (w, h))
left_frame = Frame(root, borderwidth=2, relief=SUNKEN)
left_frame.pack(side=LEFT, fill="both")
right_frame = Frame(root, borderwidth=2, relief=SOLID)
right_frame.pack(side=RIGHT, expand=True, fill="both")
details = Client()
add = Button(left_frame, text="Add Customer", command=details.add_customer_blank, height=1, width=15, font=(None, 10), relief=RAISED)
add.grid(row=0, sticky=W)
view = Button(left_frame, text="Customer Details", command=details.view_customer_blank, height=1, width=15, font=(None, 10), relief=RAISED)
view.grid(row=10, sticky=W)
daily_collection_button = Button(left_frame, text="Add Daily Collection", height=1, width=15, font=(None, 10), relief=RAISED)
daily_collection_button.bind("<Button-1>", details.daily_bind)
daily_collection_button.grid(row=30, sticky=W)
today_report = Button(left_frame, text="Today's Report", command=details.report_today, height=1, width=15, font=(None, 10), relief=RAISED)
today_report.grid(row=40, sticky=W)
previous_date = Button(left_frame, text="Add Previous Collection", wraplength=100, command=details.previous_collection_blank, height=2, width=15, font=(None, 10), relief=RAISED)
previous_date.grid(row=50, sticky=W)
previous_report = Button(left_frame, text="Previous Report", wraplength=100, command=details.pre_report_blank, height=1, width=15, font=(None, 10), relief=RAISED)
previous_report.grid(row=60, sticky=W)
report = Button(left_frame, text="Customer Report", command=details.report_customer_blank, height=1, width=15, font=(None, 10), relief=RAISED)
report.grid(row=70, sticky=W)
backup = Button(left_frame, text="Backup", command=details.backup_blank, height=1, width=15, font=(None, 10), relief=RAISED)
backup.grid(row=80, sticky=W)
root.mainloop()
