from tkinter import *
from tkinter import ttk
from tkinter.font import *


def get_between_dates(self):
	pass


def refresh(self):
	pass


def export_excel(self):
	pass


def update_assignment_details():
	print (cbo_assign.get())


root = Tk()
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

users = ['AMD','INTEL'],['ABC','ERT'],['DEF','KFG']

cbo_assign = ttk.Combobox(root,values=[user[0] for user in users] )
cbo_assign.grid(row=0, column=0, columnspan=5, sticky=N + S + W + E)
cbo_assign.bind("<<ComboboxSelected>>", lambda x: update_assignment_details())
root.mainloop()

