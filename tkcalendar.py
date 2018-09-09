from tkinter import	*
from tkinter import	ttk
import calendar
from datetime import datetime,date

class TkCalendar(Frame):
	def __init__(self, master=None,dt=None):
		self.status = 'Ok'
		super().__init__(master)
		self.grid(row=0, column=0, sticky=N + E + S + W)
		self['bg'] = 'black'
		self.rowconfigure(0, weight=1)
		self.columnconfigure(0, weight=1)
		if dt is None:
			self.dt = date.today()

		if isinstance(dt,(date)):
			self.dt = dt
		else:
			self.status ='Invalid parameter type {str(dt)}. Should be datetime.date'
		self.month = StringVar()
		self.year = StringVar()
		self.create_widget()


	def create_widget(self):
		self.container = Frame(self)
		self.container.grid(row=0,column=0,sticky=N+E+W+S)
		self.container.rowconfigure(0, weight=1)
		self.container.columnconfigure(0, weight=1)

		cbo_month = ttk.Combobox(self.container,values=calendar.month_abbr[1:],textvariable=self.month)
		cbo_month.grid(row=0,column=0,columnspan=4,sticky=N + E + S + W)
		cbo_month.rowconfigure(0, weight=1)
		cbo_month.columnconfigure(0, weight=1)
		cbo_month.bind("<<ComboboxSelected>>", lambda x: self.update_calendar(x, 'M'))
		self.month.set(calendar.month_abbr[self.dt.month])

		cbo_year = ttk.Combobox(self.container,values=range(self.dt.year-2, self.dt.year),textvariable=self.year)
		cbo_year.grid(row=0,column=5,columnspan=4,sticky=N + E + S + W)
		cbo_year.rowconfigure(0, weight=1)
		cbo_year.columnconfigure(5, weight=1)
		self.year.set(self.dt.year)
		cbo_year.bind("<<ComboboxSelected>>", lambda x: self.update_calendar(x,'Y'))

		self.frm_weekday = Frame(self.container)
		self.frm_weekday.grid(row=10,column=0,sticky=N+E+W+S,columnspan=8)
		self.frm_weekday.rowconfigure(0, weight=1)
		self.frm_weekday.columnconfigure(0, weight=1)

		self.create_date_widget(self.frm_weekday)

	def create_date_widget(self,frm_weekday):
		lbl_mo = ttk.Label(frm_weekday,text='MO')
		lbl_mo.grid(row=0,column=0,sticky=N+E+W+S)
		lbl_mo.columnconfigure(0,weight=1)
		lbl_mo.rowconfigure(0, weight=1)

		lbl_tu = ttk.Label(frm_weekday, text='TU')
		lbl_tu.grid(row=0, column=4, sticky=N + E + W + S)
		lbl_tu.columnconfigure(0, weight=1)
		lbl_tu.rowconfigure(0, weight=1)

		lbl_we = ttk.Label(frm_weekday, text='WE')
		lbl_we.grid(row=0, column=8, sticky=N + E + W + S)
		lbl_we.columnconfigure(0, weight=1)
		lbl_we.rowconfigure(0, weight=1)

		lbl_th = ttk.Label(frm_weekday, text='TH')
		lbl_th.grid(row=0, column=12, sticky=N + E + W + S)
		lbl_th.columnconfigure(0, weight=1)
		lbl_th.rowconfigure(0, weight=1)

		lbl_fr = ttk.Label(frm_weekday, text='FR')
		lbl_fr.grid(row=0, column=16, sticky=N + E + W + S)
		lbl_fr.columnconfigure(0, weight=1)
		lbl_fr.rowconfigure(0, weight=1)

		lbl_sa = ttk.Label(frm_weekday, text='SA')
		lbl_sa.grid(row=0, column=20, sticky=N + E + W + S)
		lbl_sa.columnconfigure(0, weight=1)
		lbl_sa.rowconfigure(0, weight=1)

		lbl_su = ttk.Label(frm_weekday, text='SU')
		lbl_su.grid(row=0, column=24, sticky=N + E + W + S)
		lbl_su.columnconfigure(0, weight=1)
		lbl_su.rowconfigure(0, weight=1)

		c = calendar.Calendar()
		row = 5
		col = 0
		print ((calendar.month_abbr[1:].index(self.month.get())))
		print ((self.year.get()))
		for item in c.itermonthdays2(int(self.year.get()),calendar.month_abbr[0:].index(self.month.get())):

			val = '' if item[0] == 0 else item[0]
			if val != '':
				val = '0' + str(val) if len(str(val)) == 1 else val
			btn_day = ttk.Button(frm_weekday,text=val,command=lambda arg=val:self.return_date(arg))
			if val == '':
				btn_day.state(["disabled"])
			btn_day.grid(row=row,column=col,sticky=N+E+S+W)
			btn_day.columnconfigure(row, weight=1)
			btn_day.rowconfigure(col, weight=1)
			col += 4
			if item[1] == 6:
				col = 0
				row += 5

	def return_date(self,arg):
		self.dt_selected = str(self.year.get()) + '-' + str(self.month.get()) + '-' + str(arg)


	def update_calendar(self,event,type):
		self.create_date_widget(self.frm_weekday)

if __name__ == '__main__':
	root = Tk()
	root.columnconfigure(0, weight=1)
	root.rowconfigure(0, weight=1)
	root.title("CALENDAR")

	app = TkCalendar(master=root)
	app.mainloop()
