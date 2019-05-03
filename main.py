import tkinter as tk
import datetime
import time
from tkinter import messagebox
import mysql.connector
from mysql.connector import errorcode
from tabulate import tabulate


# DataBase Class will handle all interactions with the database
# All SQL code should be written here
# If database fails to connect it will raise an exception that should be handled in the class that called it
class DataBase:
    connector = None
    cursor = None

    def __init__(self):
        self.connector = None
        self.cursor = None

    def connect(self):
        # Connect to remote database
        while 1:
            try:

                self.connector = mysql.connector.connect(
                    host="remotemysql.com",
                    user="3rybg59bIE",
                    password="Rb8WxAwcfD",
                    database="3rybg59bIE"
                )
                self.cursor = self.connector.cursor()
                return
            # Catch all errors
            except mysql.connector.Error as err:
                m = messagebox.askretrycancel("Lost Connection to Server",
                                              "Failed to connect to remote database. Check your internet connection.")
                if m is False:
                    raise Exception("Failed to connect to remote database")

                time.sleep(1)

    #
    #   @param: data: list containing customer_id, f_name, l_name
    #
    def search_customer(self, data):
        # all functions re-connect to prevent timeout
        try:
            self.connect()
        except Exception as e:
            raise e
        result = None
        i = data[0]
        # if given id
        if i is not "":
            # SQL to search customer using id
            sql = "SELECT * FROM customer WHERE customer_id = %s"
            # Execute
            self.cursor.execute(sql, (i,))
            result = self.cursor.fetchall()
            # Make sure data is committed to the database
            self.connector.commit()
            self.cursor.close()
        else:
            f = data[1]
            l = data[2]
            # SQL to search customer using names
            sql = "SELECT * FROM customer WHERE first_name LIKE %s AND last_name LIKE %s ORDER BY last_name ASC"
            usr_entry = (f + "%", l + "%")
            # Execute
            self.cursor.execute(sql, usr_entry)
            result = self.cursor.fetchall()
            # Make sure data is committed to the database
            self.connector.commit()
            self.cursor.close()
        return result

    #
    #   @param: data: list containing f_name, l_name, phone, street, city, state
    #
    def add_customer(self, data):
        # all functions re-connect to prevent timeout
        self.connect()
        # trim input data
        data = self.trim_data(data)
        # SQL to add customer
        add_customer = ("INSERT INTO customer "
                        "(first_name, last_name, phone, street, city, state) "
                        "VALUES (%s, %s, %s, %s, %s, %s)")
        # Execute - Insert new employee
        self.cursor.execute(add_customer, data)
        # Make sure data is committed to the database
        self.connector.commit()
        self.cursor.close()

    #
    #   @param: id: customer_id to delete
    #
    def remove_customer(self, id):
        # all functions re-connect to prevent timeout
        self.connect()
        # SQL to remove customer
        sql = "DELETE FROM customer WHERE customer_id = %s"
        # Execute
        self.cursor.execute(sql, (id,))
        # Make sure data is committed to the database
        self.connector.commit()
        self.cursor.close()

    #
    #   @param: data: list containing f_name, l_name, phone, street, city, state, customer_id to update
    #
    def update_customer(self, data):
        # all functions re-connect to prevent timeout
        self.connect()
        # trim data
        data = self.trim_data(data)
        # SQL to update customer
        sql = "UPDATE customer SET first_name = %s, last_name = %s, phone = %s, street = %s, city = %s, state = %s WHERE customer_id = %s"
        # Execute
        self.cursor.execute(sql, data)
        # Make sure data is committed to the database
        self.connector.commit()
        self.cursor.close()

        #
        #   Retrieves all available and non-available slips
        #
    def get_slip(self, data):
        # Re-establish connection
        self.connect()
        # SQL to retrieve data from slip table
        sql = "SELECT * FROM slip"
        # Execute SQL
        self.cursor.execute(sql, data)
        result = self.cursor.fetchall()
        # Commit data manipulation to database
        self.connector.commit()
        self.cursor.close()
        return result

    #
    #   Add boat slip
    #
    def add_slip(self, data):
        # Re-establish connection
        self.connect()
        # SQL to add slip
        add_slip = ("INSERT INTO slip "
                    "(current_lease, max_length, dock_id, customer_id) "
                    "VALUES (%s, %s, %d, %d)")
        # Execute - Insert new slip
        self.cursor.execute(add_slip, data)
        # Make sure data is committed to the database
        self.connector.commit()
        self.cursor.close()

    #
    #   Add service to the schedule
    #

    def add_service(self, data):
        # Re-establish connection
        self.connect()
        # SQL to add service
        add_service = ("INSERT INTO service "
                       "(boat_id, service")
        # Execute SQL
        self.cursor.execute(add_service, data)
        # Make sure data is committed to the database
        self.connector.commit()
        self.cursor.close()

        #
        #   Retrieves schedule of services
        #

    def get_services(self, data):
        # Re-establish connection
        self.connect()
        # SQL to retrieve services
        get_services = ("SELECT * FROM service")
        # Execute SQL
        self.cursor.execute(get_services, data)
        # Make sure data is committed to the database
        self.connector.commit()
        self.cursor.close()

    #
    #   @ param: data: list of values
    #   @ return: same list with all values trimmed
    #
    def trim_data(self, data):
        trimmed_data = list()
        for string in data:
            trimmed_data.append(str(string)[:20])
        return trimmed_data


# FingerLakesSystem Class
# Acts as system's Main Window, contains a menu frame and a main frame
class MainWindow(tk.Tk):
    main_frame = None

    def __init__(self):
        tk.Tk.__init__(self)
        self.winfo_toplevel().title("Finger Lakes Marinas System")
        # add menu bar to left side
        self.menu_frame = MenuFrame(self)
        self.menu_frame.pack(side="top")

        # load customer page
        self.switch_main_frame(CustomerPage)

    def switch_main_frame(self, frame_class):
        """Destroys current frame and replaces it with a new one."""
        new_frame = frame_class(self)
        if self.main_frame is not None:
            self.main_frame.destroy()
        self.main_frame = new_frame
        self.main_frame.pack(side="top")
        self.activate_menu()

    def activate_menu(self):
        self.menu_frame.activate()

    def disable_menu(self):
        self.menu_frame.disable()


#
# MenuFrame: Frame containing 4 buttons to allow for switching between Pages
# Methods: activate, disable
#
class MenuFrame(tk.Frame):
    slip_button = None
    service_button = None
    customer_button = None

    def __init__(self, master):
        x_pad = 30
        y_pad = 10

        tk.Frame.__init__(self, master)
        self.slip_button = tk.Button(self, text="Slip Rentals", padx=x_pad, pady=y_pad,
                                     command=lambda: master.switch_main_frame(SlipPage))
        self.service_button = tk.Button(self, text="  Services  ", padx=x_pad, pady=y_pad,
                                        command=lambda: master.switch_main_frame(ServicePage))
        self.customer_button = tk.Button(self, text="  Customer  ", padx=x_pad, pady=y_pad,
                                         command=lambda: master.switch_main_frame(CustomerPage))

        self.customer_button.pack(side="left")
        self.slip_button.pack(side="left")
        self.service_button.pack(side="left")

        self.activate()

    def activate(self):
        self.disable()
        if not isinstance(self.master.main_frame, SlipPage):
            self.slip_button.configure(state="normal")
        if not isinstance(self.master.main_frame, ServicePage):
            self.service_button.configure(state="normal")
        if not isinstance(self.master.main_frame, CustomerPage):
            self.customer_button.configure(state="normal")

    def disable(self):
        self.slip_button.configure(state="disabled")
        self.service_button.configure(state="disabled")
        self.customer_button.configure(state="disabled")


#
# ServicePage Class
# Not implemented
#
class ServicePage(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        tk.Label(self, text="This is the service page").pack(side="top", fill="x", pady=10)


#
# SlipPage Class
# Not implemented
#
class SlipPage(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        tk.Label(self, text="This is slip page").pack(side="top", fill="x", pady=10)


#
# CustomerPage Class
# Holds Lookup Customer entries/buttons and CustomerSearchPanel
#
class CustomerPage(tk.Frame):
    cur_search_frame = None
    entry_fname = None
    entry_lname = None
    entry_id = None

    def __init__(self, master):
        tk.Frame.__init__(self, master)
        # Lookup Customer ---------------------------------------------------------------------------------
        tk.Label(self, text="      Lookup Customer:      ").grid(row=0, column=0, columnspan=2)
        # first name
        tk.Label(self, text="First Name:").grid(row=1, column=0, sticky="e")
        self.entry_fname = tk.Entry(self, width=12)
        self.entry_fname.grid(row=1, column=1)
        # last name
        tk.Label(self, text="Last Name:").grid(row=2, column=0, sticky="e")
        self.entry_lname = tk.Entry(self, width=12)
        self.entry_lname.grid(row=2, column=1)
        # id
        tk.Label(self, text="Customer ID:").grid(row=3, column=0, sticky="e")
        self.entry_id = tk.Entry(self, width=12)
        self.entry_id.grid(row=3, column=1)
        # clear button
        tk.Button(self, text="Clear", padx=20,
                  command=lambda: self.clear_lookup_entry()).grid(row=4, column=0)
        # search button
        self.bind("<Return>",
                  (lambda event: self.update_search_panel(self.entry_fname.get(), self.entry_lname.get(),
                                                          self.entry_id.get())))

        tk.Button(self, text="Search", padx=20,
                  command=lambda: self.update_search_panel(self.entry_fname.get(), self.entry_lname.get(),
                                                           self.entry_id.get())).grid(row=4, column=1)

        # Add Customer ---------------------------------------------------------------------------------
        tk.Button(self, text="      Add Customer      ", command=lambda: self.add_customer()).grid(row=5, column=0,
                                                                                                   columnspan=2)

        # initially show all customers
        self.update_search_panel("", "", "")

    def update_search_panel(self, f, l, id):
        new_frame = CustomerSearchPanel(self, f, l, id)
        if self.cur_search_frame is not None:
            self.cur_search_frame.destroy()
        self.cur_search_frame = new_frame
        self.cur_search_frame.grid(row=0, column=2, rowspan=30)

    def clear_lookup_entry(self):
        self.entry_fname.delete(0, 'end')
        self.entry_lname.delete(0, 'end')
        self.entry_id.delete(0, 'end')

    def add_customer(self):
        AddCustomerPopup(self.master)
        self.update_search_panel(self.entry_fname.get(), self.entry_lname.get(),
                                 self.entry_id.get())


#
# AddCustomerPopup class
# Toplevel popup with entries and buttons to add a customer
#
class AddCustomerPopup(tk.Toplevel):
    f_name = None
    l_name = None
    phone = None
    street = None
    city = None
    state = None

    def __init__(self, parent, title="Add Customer"):
        self.top = tk.Toplevel.__init__(self, parent)
        self.transient(parent)
        if title:
            self.title(title)
        self.parent = parent
        self.result = None
        body = tk.Frame(self)
        self.initial_focus = self.body(body)
        body.pack(padx=5, pady=5)
        self.buttonbox()
        self.grab_set()
        if not self.initial_focus:
            self.initial_focus = self
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.geometry("+%d+%d" % (parent.winfo_rootx() + 50,
                                  parent.winfo_rooty() + 50))
        self.initial_focus.focus_set()
        self.wait_window(self)

    #
    # construction hooks
    def body(self, master):
        tk.Label(master, text="* First Name:").grid(row=1, sticky="e")
        tk.Label(master, text="* Last Name:").grid(row=2, sticky="e")
        tk.Label(master, text="* Phone #:").grid(row=3, sticky="e")
        tk.Label(master, text="Street Address:").grid(row=4, sticky="e")
        tk.Label(master, text="City:").grid(row=5, sticky="e")
        tk.Label(master, text="State:").grid(row=6, sticky="e")

        self.f_name = tk.Entry(master, width=15)
        self.l_name = tk.Entry(master, width=15)
        self.phone = tk.Entry(master, width=15)
        self.street = tk.Entry(master, width=15)
        self.city = tk.Entry(master, width=15)
        self.state = tk.Entry(master, width=15)

        self.f_name.grid(row=1, column=1)
        self.l_name.grid(row=2, column=1)
        self.phone.grid(row=3, column=1)
        self.street.grid(row=4, column=1)
        self.city.grid(row=5, column=1)
        self.state.grid(row=6, column=1)
        return self.f_name  # initial focus

    def buttonbox(self):
        # add standard button box. override if you don't want the
        # standard buttons
        box = tk.Frame(self)
        w = tk.Button(box, text="OK", width=10, command=self.ok, default=tk.ACTIVE)
        w.pack(side=tk.LEFT, padx=5, pady=5)
        w = tk.Button(box, text="Cancel", width=10, command=self.cancel)
        w.pack(side=tk.LEFT, padx=5, pady=5)
        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.cancel)
        box.pack()

    #
    # standard button semantics
    def ok(self, event=None):
        if not self.validate():
            self.initial_focus.focus_set()  # put focus back
            return
        self.withdraw()
        self.update_idletasks()
        self.apply()
        self.cancel()

    def cancel(self, event=None):
        # put focus back to the parent window
        self.parent.focus_set()
        self.destroy()

    #
    # command hooks
    def validate(self):
        return 1  # override

    def apply(self):
        if self.f_name.get() is not "" and self.l_name.get() is not "" and self.phone.get() is not "":
            try:
                db = DataBase()
                customer_data = (self.f_name.get(), self.l_name.get(), self.phone.get(), self.street.get(), self.city.get(),
                                 self.state.get())
                db.add_customer(customer_data)
            except:
                pass
        else:
            messagebox.showerror("Error", "Customer must have a first name, last name and phone number.")


#
# CustomerSearchPanel Class
# Holds Scrolling Canvas to display customers
#
class CustomerSearchPanel(tk.Frame):
    f_name = None
    l_name = None
    remove_id = None
    s = None
    t = None
    result = None

    def __init__(self, master, f, l, x):
        tk.Frame.__init__(self, master)
        self.configure(bg="#e6e6e6")
        self.s = tk.Scrollbar(self)
        self.t = tk.Text(self, height=50, width=150, relief="sunken")
        self.s.pack(side='right', fill='y')
        self.t.pack(side='left', fill='y')
        self.s.config(command=self.t.yview)
        self.t.config(yscrollcommand=self.s.set)
        # search database insert string into scrollbar
        self.t.insert(tk.INSERT, self.search(f, l, x))
        self.t.configure(state='disabled')

    # return string of search results
    def search(self, f, l, i):
        s = ""
        try:
            db = DataBase()
            self.result = db.search_customer((i, f, l))
            # if only one customer found display detailed view
            if self.result.__len__() == 1:
                CustomerDetailPopup(self.master.master, self.result)
                # update result in case of change
                self.result = db.search_customer((self.result[0][0], "", ""))
                # if customer was deleted result = all customers
                if self.result.__len__() == 0:
                    self.result = db.search_customer(("", "", ""))
                    self.master.clear_lookup_entry()
            # only display first 6 columns
            for i in range(0, self.result.__len__()):
                self.result[i] = self.result[i][:7]
            s = tabulate(self.result,
                         headers=["ID", "First Name", "Last Name", "Phone", "Street", "City", "State"],
                         tablefmt="simple")
        except Exception as e:
            s = e
        return s


#
# CustomerDetailPopup
# Toplevel popup, shows single customer's details and allows user to modify or delete customer
#
class CustomerDetailPopup(tk.Toplevel):
    customer = None
    f_name = None
    l_name = None
    phone = None
    street = None
    city = None
    state = None
    apply_button = None
    edit_button = None
    close_button = None

    def __init__(self, parent, result, title="Customer Details"):
        self.top = tk.Toplevel.__init__(self, parent)
        self.customer = result
        self.transient(parent)
        if title:
            self.title(title)
        self.parent = parent
        self.result = None
        body = tk.Frame(self)
        self.initial_focus = self.body(body)
        body.pack(padx=5, pady=5)
        self.buttonbox()
        self.grab_set()
        if not self.initial_focus:
            self.initial_focus = self
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.geometry("+%d+%d" % (parent.winfo_rootx() + 50,
                                  parent.winfo_rooty() + 50))
        self.initial_focus.focus_set()
        self.wait_window(self)

    #
    # construction hooks
    def body(self, master):
        tk.Label(master, text="First Name:").grid(row=0, sticky="e")
        tk.Label(master, text="Last Name:").grid(row=1, sticky="e")
        tk.Label(master, text="Phone #:").grid(row=2, sticky="e")
        tk.Label(master, text="Street Address:").grid(row=3, sticky="e")
        tk.Label(master, text="City:").grid(row=4, sticky="e")
        tk.Label(master, text="State:").grid(row=5, sticky="e")

        self.f_name = tk.Entry(master)
        self.f_name.insert(tk.END, self.customer[0][1])
        self.l_name = tk.Entry(master)
        self.l_name.insert(tk.END, self.customer[0][2])
        self.phone = tk.Entry(master)
        self.phone.insert(tk.END, self.customer[0][3])
        self.street = tk.Entry(master)
        self.street.insert(tk.END, self.customer[0][4])
        self.city = tk.Entry(master)
        self.city.insert(tk.END, self.customer[0][5])
        self.state = tk.Entry(master)
        self.state.insert(tk.END, self.customer[0][6])
        self.disable_entries()

        self.f_name.grid(row=0, column=1)
        self.l_name.grid(row=1, column=1)
        self.phone.grid(row=2, column=1)
        self.street.grid(row=3, column=1)
        self.city.grid(row=4, column=1)
        self.state.grid(row=5, column=1)
        return self.f_name  # initial focus

    def disable_entries(self):
        self.f_name.configure(state="disabled")
        self.l_name.configure(state="disabled")
        self.phone.configure(state="disabled")
        self.street.configure(state="disabled")
        self.city.configure(state="disabled")
        self.state.configure(state="disabled")

    def buttonbox(self):
        # add button box
        box = tk.Frame(self)
        w = tk.Button(box, text="Delete Customer", width=15, command=self.delete_customer)
        w.pack(side=tk.LEFT, padx=5, pady=5)
        self.edit_button = tk.Button(box, text="Edit Customer", width=15, command=self.enable_entries)
        self.edit_button.pack(side=tk.LEFT, padx=5, pady=5)
        self.apply_button = tk.Button(box, text="Apply Changes", width=15, command=self.ok, state=tk.DISABLED)
        self.apply_button.pack(side=tk.LEFT, padx=5, pady=5)
        self.close_button = tk.Button(box, text="Close", width=10, command=self.cancel)
        self.close_button.pack(side=tk.LEFT, padx=5, pady=5)
        self.bind("<Return>", self.cancel)
        self.bind("<Escape>", self.cancel)
        box.pack()

    #
    # standard button semantics
    def ok(self, event=None):
        if not self.validate():
            self.initial_focus.focus_set()  # put focus back
            return
        self.withdraw()
        self.update_idletasks()
        self.apply()
        self.cancel()

    def cancel(self, event=None):
        # put focus back to the parent window
        self.parent.focus_set()
        self.destroy()

    #
    # command hooks
    def validate(self):
        return 1  # override

    def apply(self):
        # confirm removal
        msg_box = tk.messagebox.askquestion("Confirm Update",
                                            "Are you sure you want to update this customer?\n",
                                            icon='warning')
        if msg_box == 'yes':
            usr_entry = (
                self.f_name.get(), self.l_name.get(), self.phone.get(), self.street.get(), self.city.get(),
                self.state.get(),
                self.customer[0][0])
            try:
                db = DataBase()
                db.update_customer(usr_entry)
            except:
                pass
            self.parent.focus_set()
            self.destroy()

    def enable_entries(self):
        self.f_name.configure(state="normal")
        self.l_name.configure(state="normal")
        self.phone.configure(state="normal")
        self.street.configure(state="normal")
        self.city.configure(state="normal")
        self.state.configure(state="normal")
        self.apply_button.configure(state="normal")
        self.bind("<Return>", self.ok)
        self.edit_button.configure(state=tk.DISABLED)

    def delete_customer(self):
        # confirm removal
        msg_box = tk.messagebox.askquestion("Confirm Removal",
                                            "Are you sure you want to remove this customer? \n\nThis action cannot be undone.\n",
                                            icon='warning')
        if msg_box == 'yes':
            try:
                db = DataBase()
                db.remove_customer(self.customer[0][0])
            except:
                pass
            self.cancel()


if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
