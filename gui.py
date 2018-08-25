from tkinter import *
from tkinter import ttk
from tkinter.font import *
import connect_to_sqlite as sql
import queries as q
from datetime import datetime,timedelta
import main as m


class Application(Frame):
	def __init__(self, master=None):
		super().__init__(master)
		self.grid(row=0,column=0,sticky=N+E+S+W)
		self['bg'] = 'black'
		self.rowconfigure(0, weight = 1)
		self.columnconfigure(0, weight=1)
		self.conn = sql.SQLInterface('maildb.db')
		self.users = []
		self.assignee = StringVar()
		self.bold_button = Font(family="Fixedsys", size=10, weight=BOLD)
		self.normal_button = Font(family="Fixedsys", size=8, weight=NORMAL)

		self.get_user_list()
		self.create_widgets()


	def get_user_list(self):
		query = q.query_dict['get_active_users']
		results = self.conn.get_data(query)
		if self.conn.status == 'Ok':
			for result in results:
				self.users.append([result[0], result[1]])
		else:
			self.message_from_application('Error', self.conn.status)


	def identify_page_number(self,up_down_ind):

		idx = self.options.index(self.page.get())
		if idx + up_down_ind < 0:
			pass
		elif idx + up_down_ind > len(self.options):
			pass
		else:
			self.page.set(self.options[idx+up_down_ind])
			self.pagination(self.options[idx+up_down_ind])

	def create_widgets(self):
		style = ttk.Style()
		style.configure("C.TButton", background="white", font=("Fixedsys", 10, 'bold'), foreground='black')
		style.configure("C.TCombobox", background="white", font=("Fixedsys", 10, 'bold'), foreground='black')

		utility_frame = Frame(self, bg='black')
		utility_frame.grid(row=0, column=0, sticky=N+E+S+W, padx=10, pady=10, ipadx=5, ipady=5)
		utility_frame.rowconfigure(0, weight=1)
		utility_frame.columnconfigure(0, weight=1, uniform="fred")


		self.btn_between_dates = Button(utility_frame, text="FETCH BETWEEN DATES", font=self.bold_button,command=self.get_between_dates, width=30,bg='white',fg='black')
		self.btn_between_dates.grid(row=0, column=0 ,columnspan=5,padx=(2,15),pady=2,ipadx=2,ipady=2,sticky=N + E + S + W)
		self.btn_between_dates.columnconfigure(0, weight=1,uniform='third')

		self.btn_refresh = Button(utility_frame, text="REFRESH", font=self.bold_button,command=self.refresh, width=30,bg='white',fg='black')
		self.btn_refresh.grid(row=0, column=15,columnspan=5, padx=(15,15),pady=2,ipadx=2,ipady=2,sticky=N + E + S + W)
		self.btn_refresh.columnconfigure(15, weight=1,uniform='third')

		self.btn_export_to_excel = Button(utility_frame, text="EXPORT TO EXCEL", font=self.bold_button,command=self.export_excel, width=30,bg='white',fg='black')
		self.btn_export_to_excel.grid(row=0, column=30, columnspan=5,padx=(15,2),pady=2,ipadx=2,ipady=2,sticky=N + E + S + W)
		self.btn_export_to_excel.grid(row=0, column=30, columnspan=5,padx=(15,2),pady=2,ipadx=2,ipady=2,sticky=N + E + S + W)
		self.btn_export_to_excel.columnconfigure(30, weight=1,uniform='third')

		self.page = StringVar()
		mail_utility = Frame(self, bg='black')
		mail_utility.grid(row=10, column=0, sticky=N + E + S + W, padx=10,pady=10,ipadx=5,ipady=5)
		mail_utility.rowconfigure(0, weight=1)
		mail_utility.columnconfigure(0, weight=1)

		self.btn_prev = ttk.Button(mail_utility, text='PREVIOUS', command=lambda: self.identify_page_number(-1), width=30,style='C.TButton')
		self.btn_prev.grid(row=0, column=0, columnspan=5, sticky=N + S + W + E, padx=(2,15), pady=2, ipadx=0, ipady=2)
		self.btn_prev.rowconfigure(0, weight=1)
		self.btn_prev.columnconfigure(0, weight=1)

		self.cbo_page_list = ttk.Combobox(mail_utility, textvariable=self.page,style='C.TCombobox')
		self.cbo_page_list.config(width=30)
		self.cbo_page_list.grid(row=0, column=20, columnspan=5, sticky=N + S + W + E, padx=(15,15), pady=2, ipadx=2, ipady=2)
		self.cbo_page_list.bind("<<ComboboxSelected>>", lambda x: self.pagination(self.page.get()))

		self.btn_next = ttk.Button(mail_utility, text='NEXT', command=lambda: self.identify_page_number(1), width=30,style='C.TButton')
		self.btn_next.grid(row=0, column=40, columnspan=5, sticky=N + S + W + E, padx=(15,2), pady=2, ipadx=0, ipady=2)
		self.btn_next.rowconfigure(0, weight=1)
		self.btn_next.columnconfigure(40, weight=1)

		self.mail_frame = Frame(self, bg='green')
		self.mail_frame.grid(row=20,column=0,sticky=N + E + S + W,padx=10,pady=10)
		self.mail_frame.rowconfigure(0, weight=1)
		self.mail_frame.columnconfigure(0, weight=1)
		#mail_frame.pack(fill=BOTH, expand=1)
		self.initiate_mail_process()

	def get_between_dates(self):
		pass

	def refresh(self):
		pass

	def export_excel(self):
		pass

	def update_assignment_details(self):
		mail_id = 'mail[-1]'
		subject = 'mail[1]'
		sender = 'mail[5]'
		mail_dt = "mail[3].strftime('%Y-%m-%d %H:%M:%S')"
		print (f'Length {len(self.cbo_assign.get())}')
		updt_query = q.query_dict['update_user_assignment'],[self.cbo_assign.get(),mail_id]
		sel_query = q.query_dict['search_mail_assignment'],[mail_id]
		ins_query = q.query_dict['insert_mail_assignment'],[mail_id,subject,sender,mail_dt,self.cbo_assign.get()]
		self.conn.upsert(sel_query=sel_query,updt_query=updt_query,ins_query=ins_query)

		if self.conn.status == 'Ok':
			pass
		else:
			self.message_from_application('Error', self.conn.status)

	def pagination(self,page):
		page_view_start = int(page.split('-')[0].strip())
		page_view_end = int(page.split('-')[1].strip())
		row=5
		print (self.mail_list[0])
		for mail in self.mail_list[page_view_start:page_view_end]:
			col = 0
			folder_label =Label(self.mail_frame, text = mail[0],bg='white',fg='black')
			folder_label.grid(row=row,column=col,sticky=N+S+W+E)
			col += 5
			subject_label = Label(self.mail_frame, text=mail[1][:45],anchor=W,bg='white',fg='black')
			subject_label.grid(row=row, column=col, columnspan=5, sticky=N + S + W + E)
			col += 5
			sender_label = Label(self.mail_frame, text=str(mail[5]),anchor=W,bg='white',fg='black')
			sender_label.grid(row=row, column=col, sticky=N + S + W + E)
			col += 5
			send_date_label = Label(self.mail_frame, text=str(mail[3])[:16],bg='white',fg='black')
			send_date_label.grid(row=row, column=col, columnspan=5, sticky=N + S + W + E)
			col += 5
			btn_details = Button(self.mail_frame,text='Details',command=lambda : self.details(),font=self.normal_button,bg='azure',fg='black' )
			btn_details.grid(row=row, column=col, columnspan=5, sticky=N + S + W + E)
			col += 5
			user_names = [user[0] for user in self.users]
			self.cbo_assign = ttk.Combobox(self.mail_frame,values=user_names )
			self.cbo_assign.grid(row=row, column=col, columnspan=5, sticky=N + S + W + E)
			self.cbo_assign.bind("<<ComboboxSelected>>", lambda x: self.update_assignment_details())
			col += 5
			row += 5



	def message_from_application(self,type,message):
		window = Toplevel(self)

		style = ttk.Style()
		style.configure("E.TButton",background="salmon",font=("Fixedsys", 10, 'bold'), foreground='red')

		topFrame = Frame(window)
		topFrame.grid(row=0, column=0, sticky=N + E + W + S)
		topFrame.rowconfigure(0, weight=1)
		topFrame.columnconfigure(0, weight=1)

		if type == 'Error':
			window.title("Error Details")
			lbl_policy_number = ttk.Label(topFrame, text=message, width=30)
			lbl_policy_number.grid(row=0, column=0, sticky=N + E + W + S)
			lbl_policy_number.rowconfigure(0, weight=1)
			lbl_policy_number.columnconfigure(0, weight=1)

			btn_ok = ttk.Button(topFrame, text='Ok', command=sys.exit, width=30,style='E.TButton')
			btn_ok.grid(row=10,column=0,sticky=N+E+S+W)

	def details(self):
		window = Toplevel(self)
		window.title("Additional Details")

		topFrame = Frame(window)
		topFrame.grid(row=0,column=0,sticky=N+E+W+S)
		topFrame.rowconfigure(0, weight=1)
		topFrame.columnconfigure(0, weight=1)
		self.policy_number = StringVar()
		self.cusomer_type = StringVar()
		self.system = StringVar()
		self.inquiry_cat = StringVar()
		self.inquiry = StringVar()

		lbl_policy_number = ttk.Label(topFrame,text='Policy Number',width=30)
		lbl_policy_number.grid(row=0,column=0,sticky=N+E+W+S)
		lbl_policy_number.rowconfigure(0, weight=1)
		lbl_policy_number.columnconfigure(0, weight=1)

		ent_policy_number = ttk.Entry (topFrame,textvariable=self.policy_number)
		ent_policy_number.grid(row=0, column=40, sticky=N + E + W + S)
		ent_policy_number.rowconfigure(0, weight=1)
		ent_policy_number.columnconfigure(0, weight=1)

		lbl_customer_type = ttk.Label(topFrame,text='Customer Type',width=30)
		lbl_customer_type.grid(row=10, column=0, sticky=N + E + W + S)
		lbl_customer_type.rowconfigure(10, weight=1)
		lbl_customer_type.columnconfigure(0, weight=1)

		ent_customer_type = ttk.Entry (topFrame,textvariable=self.cusomer_type)
		ent_customer_type.grid(row=10, column=40, sticky=N + E + W + S)
		ent_customer_type.rowconfigure(10, weight=1)
		ent_customer_type.columnconfigure(0, weight=1)

		lbl_system = ttk.Label(topFrame,text='System',width=30)
		lbl_system.grid(row=20, column=0, sticky=N + E + W + S)
		lbl_system.rowconfigure(20, weight=1)
		lbl_system.columnconfigure(0, weight=1)

		ent_system = ttk.Entry (topFrame,textvariable=self.system)
		ent_system.grid(row=20, column=40, sticky=N + E + W + S)
		ent_system.rowconfigure(20, weight=1)
		ent_system.columnconfigure(0, weight=1)

		lbl_inquiry_category = ttk.Label(topFrame,text='Inquiry Category',width=30)
		lbl_inquiry_category.grid(row=30, column=0, sticky=N + E + W + S)
		lbl_inquiry_category.rowconfigure(30, weight=1)
		lbl_inquiry_category.columnconfigure(0, weight=1)

		ent_inquiry_category = ttk.Entry (topFrame,textvariable=self.inquiry_cat)
		ent_inquiry_category.grid(row=30, column=40, sticky=N + E + W + S)
		ent_inquiry_category.rowconfigure(30, weight=1)
		ent_inquiry_category.columnconfigure(0, weight=1)

		lbl_inquiry = ttk.Label(topFrame,text='Inquiry',width=30)
		lbl_inquiry.grid(row=40, column=0, sticky=N + E + W + S)
		lbl_inquiry.rowconfigure(40, weight=1)
		lbl_inquiry.columnconfigure(0, weight=1)

		ent_inquiry = ttk.Entry (topFrame,textvariable=self.inquiry)
		ent_inquiry.grid(row=40, column=40, sticky=N + E + W + S)
		ent_inquiry.rowconfigure(40, weight=1)
		ent_inquiry.columnconfigure(0, weight=1)

		mail_body = Text(topFrame)
		mail_body.grid(row=60, column=0, sticky=N + E + W + S,columnspan=100)
		mail_body.rowconfigure(60, weight=1)
		mail_body.columnconfigure(0, weight=1)

		btn_save = ttk.Button(topFrame,text='SAVE',width=30)
		btn_save.grid(row=150,column=0)
		btn_cancel = ttk.Button(topFrame, text='CANCEL', width=30)
		btn_cancel.grid(row=150,column=40)

	def initiate_mail_process(self):
		last_date_from_sql = self.get_last_run()

		if last_date_from_sql[0] is None:
			start_date =  (datetime.now() - timedelta(days=1)).strftime('%m/%d/%Y')
			end_date = datetime.now().strftime('%m/%d/%Y')
		else:
			start_date = datetime.strptime(last_date_from_sql[0],'%Y-%m-%d %H:%M:%S').strftime('%m/%d/%Y')
			end_date = datetime.now().strftime('%m/%d/%Y')

		mails = m.ReadOutLookMails()
		print (f'Mail Process Started {datetime.now()}')
		self.mail_list = mails.get_mails(start_date, end_date)

		print(f'Mail Process Ender {datetime.now()}')
		total_number_of_mail = len(self.mail_list)

		self.options = [f'{x} - {x + 20}' for x in range(1,total_number_of_mail,20)]
		self.cbo_page_list['values'] = self.options
		self.page.set('1 - 21')

		self.pagination(self.page.get())

	def get_last_run(self):
		query = q.query_dict['get_last_run']
		results = self.conn.get_data(query)
		if self.conn.status == 'Ok':
			return results[0]
		else:
			return None




root = Tk()
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
root.title("CHUBB TICKET ASSIGNMENT")
app = Application(master=root)
app.mainloop()
