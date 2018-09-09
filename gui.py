from tkinter import *
from tkinter import ttk
from tkinter.font import *
import connect_to_sqlite as sql
import queries as q
from datetime import datetime,timedelta
import main as m
import threading
import pythoncom
import re
import json
import tkcalendar as cal
import os
import pandas as pd
import queue
from operator import itemgetter
from fuzzywuzzy import fuzz

class Application(Frame):
	def __init__(self, master=None):
		self.mail_list = None
		self.q = queue.Queue()
		self.SESSION = os.getlogin() + '_' + str(datetime.now())
		self.conn = sql.SQLInterface('maildb.db')
		self.initiate_thread_for_mail_extraction()
		self.initiate_thread_for_mail_details()
		super().__init__(master)
		self.grid(row=0,column=0,sticky=N+E+S+W)
		self['bg'] = 'black'
		self.rowconfigure(0, weight = 1)
		self.columnconfigure(0, weight=1)

		self.users = []
		self.assignee = StringVar()
		self.bold_button = Font(family="Fixedsys", size=10, weight=BOLD)
		self.normal_button = Font(family="Fixedsys", size=8, weight=NORMAL)

		self.get_user_list()
		self.create_widgets()

	def get_mails(self):
		pythoncom.CoInitialize()
		mail_det_conn = sql.SQLInterface('maildb.db')
		last_date_from_sql = self.get_last_run(mail_det_conn)
		print (f'The last date is{last_date_from_sql}Sesion is {self.SESSION}')

		if last_date_from_sql[0] is None:
			start_date = (datetime.now() - timedelta(hours=24)).strftime('%m/%d/%Y %I:%M %p')
			end_date = datetime.now().strftime('%m/%d/%Y %I:%M %p')
		else:
			start_date = last_date_from_sql[0]
			end_date = datetime.now().strftime('%m/%d/%Y %I:%M %p')

		print (f'{start_date}-{end_date}')
		mails = m.ReadOutLookMails()
		print(f'Mail Process Started {datetime.now()}')
		self.q.put(mails.get_mails('dt',start_date, end_date))

		last_run_query = q.query_dict['insert_session_data'],[self.SESSION,end_date]

		mail_det_conn.insert_data(last_run_query)
		if mail_det_conn.status == 'Ok':
			pass
		else:
			self.message_from_application('Error', mail_det_conn.status)
		mail_det_conn.close_conn()



	def get_mail_data(self):
		mail_det_conn = sql.SQLInterface('maildb.db',ro=True)
		mail_details_query = q.query_dict['get_mail_details'], [datetime.now() - timedelta(days=90)]
		results = mail_det_conn.get_data(mail_details_query)
		self.mail_details = []
		if mail_det_conn.status == 'Ok':
			for result in results:
				self.mail_details.append([result[0], result[1]])
		else:
			self.message_from_application('Error', mail_det_conn.status)
		mail_det_conn.close_conn()

	def initiate_thread_for_mail_details(self):
		t = threading.Thread(target=self.get_mail_data)
		t.start()
		self.threads['mail_details'] = t

	def initiate_thread_for_mail_extraction(self):
		self.threads = {}
		t = threading.Thread(target=self.get_mails)
		self.threads['mail_extraction'] = t
		t.start()

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
		style = ttk.Style(self)

		style.configure("C.TButton", background="white", font=("Fixedsys", 10, 'bold'), foreground='black')
		style.configure("C.TCombobox", background="white", font=("Fixedsys", 10, 'bold'), foreground='black')

		utility_frame = Frame(self, bg='black')
		utility_frame.grid(row=0, column=0, sticky=N+E+S+W, padx=10, pady=10, ipadx=5, ipady=5)
		utility_frame.rowconfigure(0, weight=1)
		utility_frame.columnconfigure(0, weight=1, uniform="fred")


		self.btn_between_dates = Button(utility_frame, text="FETCH BETWEEN DATES", font=self.bold_button,command=self.get_between_dates, width=30,bg='white',fg='black')
		self.btn_between_dates.grid(row=0, column=0 ,columnspan=15,padx=(2,15),pady=2,ipadx=2,ipady=2,sticky=N + E + S + W)
		self.btn_between_dates.columnconfigure(0, weight=1,uniform='third')

		self.btn_refresh = Button(utility_frame, text="REFRESH", font=self.bold_button,command=self.refresh, width=30,bg='white',fg='black')
		self.btn_refresh.grid(row=0, column=15,columnspan=15, padx=(15,15),pady=2,ipadx=2,ipady=2,sticky=N + E + S + W)
		self.btn_refresh.columnconfigure(15, weight=1,uniform='third')

		self.btn_export_to_excel = Button(utility_frame, text="EXPORT TO EXCEL", font=self.bold_button,command=lambda : self.export_excel(self.mail_list), width=30,bg='white',fg='black')
		self.btn_export_to_excel.grid(row=0, column=30, columnspan=15,padx=(15,2),pady=2,ipadx=2,ipady=2,sticky=N + E + S + W)
		self.btn_export_to_excel.columnconfigure(30, weight=1,uniform='third')

		self.page = StringVar()
		mail_utility = Frame(self, bg='black')
		mail_utility.grid(row=10, column=0, sticky=N + E + S + W, padx=10,pady=10,ipadx=5,ipady=5)
		mail_utility.rowconfigure(0, weight=1)
		mail_utility.columnconfigure(0, weight=1,uniform="fred")

		self.btn_prev = Button(mail_utility, text='PREVIOUS', font=self.bold_button,command=lambda: self.identify_page_number(-1), width=30,bg='white',fg='black')
		self.btn_prev.grid(row=0, column=0, columnspan=15, sticky=N + S + W + E, padx=(2,15), pady=2,ipadx=2,  ipady=2)
		self.btn_prev.columnconfigure(0, weight=1,uniform='third')

		self.cbo_page_list = ttk.Combobox(mail_utility, textvariable=self.page,style='C.TCombobox')
		self.cbo_page_list.config(width=45)
		self.cbo_page_list.grid(row=0, column=15, columnspan=15, sticky=N + S + W + E, padx=(15,15), pady=2, ipadx=2, ipady=2)
		self.cbo_page_list.bind("<<ComboboxSelected>>", lambda x: self.pagination(self.page.get()))
		self.cbo_page_list.columnconfigure(0, weight=1,uniform='third')

		self.btn_next = Button(mail_utility, text='NEXT', font=self.bold_button,command=lambda: self.identify_page_number(1), width=30,bg='white',fg='black')
		self.btn_next.grid(row=0, column=30, columnspan=15, sticky=N + S + W + E, padx=(15,2), pady=2, ipadx=0, ipady=2)
		self.btn_next.columnconfigure(40, weight=1,uniform='third')

		self.mail_frame = Frame(self, bg='green')
		self.mail_frame.grid(row=20,column=0,sticky=N + E + S + W,padx=2,pady=2)
		self.mail_frame.rowconfigure(0, weight=1)
		self.mail_frame.columnconfigure(0, weight=1,uniform="fred")
		self.initiate_mail_process()

	def get_between_dates(self):
		style = ttk.Style()
		style.configure("E.TButton", background="salmon", font=("Fixedsys", 10, 'bold'), foreground='red')
		app = cal.TkCalendar(master=Toplevel(self))
		app.create_widget()

	def refresh(self):
		self.initiate_thread_for_mail_extraction()
		self.threads['mail_extraction'].join()
		self.mail_list = self.q.get() + self.mail_list
		print(f'Mail Process Ender {datetime.now()}')
		total_number_of_mail = len(self.mail_list)

		self.options = [f'{x} - {x + 20}' for x in range(1,total_number_of_mail,20)]
		self.cbo_page_list['values'] = self.options
		self.page.set('1 - 21')

		self.pagination(self.page.get())

	def export_excel(self,l):
		p = os.path.join('.','FILES')
		file_name = f"download_{datetime.now().strftime('%Y_%m_%d_%H_%M_%S')}.xlsx"
		print (file_name)
		filepath = p
		if os.path.exists(p):
			pass
		else:
			os.mkdir(p)
		filepath = os.path.join(p,file_name)
		print (filepath)

		data = []
		sent_time_pattern = re.compile(r'Sent: (.* [AM|PM]+).*')
		for item in l:

			find_all = sent_time_pattern.findall(item[2])
			dates = []
			if find_all:
				try:
					dates.append(datetime.strptime(find_all[0], '%A, %B %d, %Y %I:%M %p'))
				except Exception as ex:
					pass
			first_mail_date = None
			if dates:
				first_mail_date = min(dates)

			data.append([item[0],item[1],item[5],item[3],first_mail_date,item[6]])
		#print (data[:3])
		df = pd.DataFrame(data=data,
			columns=['Folder Name', 'Subject', 'Sender Name', 'Time Received', 'First Mail',
						'Conversation ID'])
		#df.to_excel('test.xlsx')
		#print(f'The length of dataframe is {len(df)}')
		df.sort_values(['Conversation ID', 'Time Received'], ascending=[True, False], inplace=True)
		df['Aging'] = df['Time Received'].subtract(df['First Mail']).dt.days
		df.to_excel(filepath, index=False)

		window = Toplevel(self)
		window.title('File Download Information')
		style = ttk.Style()
		style.configure("E.TButton", background="salmon", font=("Fixedsys", 10, 'bold'), foreground='red')

		topFrame = Frame(window)
		topFrame.grid(row=0, column=0, sticky=N + E + W + S)
		topFrame.rowconfigure(0, weight=1)
		topFrame.columnconfigure(0, weight=1)

		lbl_msg = ttk.Label(topFrame,text='File is downloaded. Open File or Folder')
		lbl_msg.grid(row=0,column=0,sticky=N + E + W + S,columnspan=4)
		lbl_msg.rowconfigure(0, weight=1)
		lbl_msg.columnconfigure(0, weight=1)

		btn_file = ttk.Button(topFrame,text='OPEN FILE',command= lambda: os.startfile(filepath))
		btn_file.grid(row=10,column=0,sticky=N + E + W + S)
		btn_file.rowconfigure(0, weight=1)
		btn_file.columnconfigure(0, weight=1)

		btn_folder = ttk.Button(topFrame,text='OPEN FOLDER',command=lambda: os.startfile(p))
		btn_folder.grid(row=10,column=10,sticky=N + E + W + S)
		btn_folder.rowconfigure(0, weight=1)
		btn_folder.columnconfigure(0, weight=1)


	def update_assignment_details(self,*args):
		mail_id = args[2][6]
		subject = args[2][1]
		sender = args[2][5]
		mail_dt = args[2][3].strftime('%Y-%m-%d %H:%M:%S')
		print (f'{args[1].get()} {mail_id}')
		updt_query = q.query_dict['update_user_assignment'],[args[1].get(),mail_id]
		sel_query = q.query_dict['search_mail_assignment'],[mail_id]
		ins_query = q.query_dict['insert_mail_assignment'],[mail_id,subject,sender,mail_dt,args[1].get()]
		self.conn.upsert(sel_query=sel_query,updt_query=updt_query,ins_query=ins_query)

		if self.conn.status == 'Ok':
			pass
		else:
			self.message_from_application('Error', self.conn.status)

	def pagination(self,page):
		page_view_start = int(page.split('-')[0].strip())
		page_view_end = int(page.split('-')[1].strip())
		row=5
		print (self.unique_mail[0])
		for mail in self.unique_mail[page_view_start:page_view_end]:
			col = 0
			folder_label =Label(self.mail_frame,anchor=W, text = str(mail[1]),bg='alice blue',fg='black',width=10)
			folder_label.grid(row=row,column=col,sticky=N+S+W+E)
			col += 1
			subject_label = Label(self.mail_frame, text=mail[5],anchor=W,bg='gainsboro',fg='black')
			subject_label.grid(row=row, column=col,  sticky=N + S + W + E)
			col += 5
			sender_label = Label(self.mail_frame, text=str(mail[3])[:16],anchor=W,bg='alice blue',fg='black')
			sender_label.grid(row=row, column=col, sticky=N + S + W + E)
			col += 5
			send_date_label = Label(self.mail_frame, text=str(mail[0]),bg='gainsboro',fg='black')
			send_date_label.grid(row=row, column=col,  sticky=N + S + W + E)
			col += 5
			btn_details = Button(self.mail_frame,text='Details',command=lambda mail=mail : self.details(mail),font=self.normal_button,bg='azure',fg='black' )
			btn_details.grid(row=row, column=col,  sticky=N + S + W + E)
			col += 5
			user_names = [user[0] for user in self.users]
			self.cbo_assign = ttk.Combobox(self.mail_frame,values=user_names )
			self.cbo_assign.grid(row=row, column=col,  sticky=N + S + W + E)
			self.cbo_assign.bind("<<ComboboxSelected>>", lambda x,cbo=self.cbo_assign,mail=mail: self.update_assignment_details(x,cbo,mail))
			col += 5
			btn_open_conversation = Button(self.mail_frame, text='Conversation', command=lambda mail=mail: self.open_conversation(mail),
								 font=self.normal_button, bg='azure', fg='black')
			btn_open_conversation.grid(row=row, column=col, sticky=N + S + W + E)
			row += 5

	def generate_form(self,cnv_frame,row,mails,disable='False'):

		for item in mails:
			col = 0
			folder_label = Label(cnv_frame, anchor=W, text=str(item[1]), bg='alice blue', fg='black')
			folder_label.grid(row=row, column=col, sticky=N + S + W + E)
			folder_label.rowconfigure(row,weight=1)
			folder_label.columnconfigure(col,weight=1)
			col += 1
			subject_label = Label(cnv_frame, text=item[5], anchor=W, bg='gainsboro', fg='black')
			subject_label.grid(row=row, column=col, sticky=N + S + W + E)
			subject_label.rowconfigure(row, weight=1)
			subject_label.columnconfigure(col, weight=1)
			col += 5
			sender_label = Label(cnv_frame, text=str(item[3])[:16], anchor=W, bg='alice blue', fg='black')
			sender_label.grid(row=row, column=col, sticky=N + S + W + E)
			sender_label.rowconfigure(row, weight=1)
			sender_label.columnconfigure(col, weight=1)
			col += 5
			send_date_label = Label(cnv_frame, text=str(item[0]), bg='gainsboro', fg='black')
			send_date_label.grid(row=row, column=col, sticky=N + S + W + E)
			send_date_label.rowconfigure(row, weight=1)
			send_date_label.columnconfigure(col, weight=1)
			col += 5
			btn_details = Button(cnv_frame, text='DETAILS',  font=self.normal_button, bg='azure', fg='black')
			btn_details.grid(row=row, column=col, sticky=N + S + W + E)
			btn_details.rowconfigure(row, weight=1)
			btn_details.columnconfigure(col, weight=1)
			col += 5
			btn_details = Button(cnv_frame, text='LINK/UNLINK', font=self.normal_button, bg='azure', fg='black',state='disabled')
			btn_details.grid(row=row, column=col, sticky=N + S + W + E)
			btn_details.rowconfigure(row, weight=1)
			btn_details.columnconfigure(col, weight=1)
			col += 5
			row += 5

	def open_conversation(self,mail):
		window = Toplevel(self)
		topFrame = Frame(window)
		window.title('Conversation Details')
		topFrame.grid(row=0, column=0, sticky=N + E + W + S)
		topFrame.rowconfigure(0, weight=1)
		topFrame.columnconfigure(0, weight=1)

		rd_mail = m.ReadOutLookMails()
		mails = rd_mail.get_mails('ct',mail[7])
		row = 0

		self.generate_form(topFrame,row,mails)
		row += len(mails)*5
		mail_details = []
		for item in self.mail_list:
			if mail[6] != item[6]:
				if fuzz.ratio(item[1],mail[1]) > 90:
					mail_details.append(item)

		self.generate_form(topFrame, row, mail_details,True)

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

	def unique(self,l):
		return list(set(l))


	def analyze_mail_body(self,mail_id,mail_body):
		for detail in self.mail_details:
			if detail[0] == mail_id:
				return json.loads(detail[1])


		policy_number_pattern = re.compile(r'(D[0-9][0-9][0-9][0-9][0-9][0-9][0-9][A-Za-z0-9])')

		find_all = policy_number_pattern.findall(mail_body)
		print (str(find_all))
		return {'policy':self.unique(find_all)}

	def reset_textvar(self,window):
		self.policy_number.set(None)
		self.cusomer_type.set(None)
		self.system.set(None)
		self.inquiry_cat.set(None)
		self.inquiry.set(None)
		window.destroy()

	def details(self,mail):

		mail_details = self.analyze_mail_body(mail[6],mail[2])
		print (mail_details)
		window = Toplevel(self)
		window.title('Additional Details')
		window.protocol('WM_DELETE_WINDOW',lambda x=window:self.reset_textvar(x))
		topFrame = Frame(window)
		topFrame.grid(row=0,column=0,sticky=N+E+W+S)
		topFrame.rowconfigure(0, weight=1)
		topFrame.columnconfigure(0, weight=1)
		self.policy_number = StringVar()
		self.cusomer_type = StringVar()
		self.system = StringVar()
		self.inquiry_cat = StringVar()
		self.inquiry = StringVar()


		scrollbar = Scrollbar(topFrame)
		scrollbar.grid(row=60,column=100,sticky=N+S+E)

		self.policy_number.set(mail_details['policy'])
		lbl_policy_number = ttk.Label(topFrame,text='Policy Number',width=30)
		lbl_policy_number.grid(row=0,column=0,sticky=N+E+W+S)
		lbl_policy_number.rowconfigure(0, weight=1)
		lbl_policy_number.columnconfigure(0, weight=1)

		ent_policy_number = ttk.Entry (topFrame,textvariable=self.policy_number)
		ent_policy_number.grid(row=0, column=40, sticky=N + E + W + S)
		ent_policy_number.rowconfigure(0, weight=1)
		ent_policy_number.columnconfigure(40, weight=1)

		lbl_customer_type = ttk.Label(topFrame,text='Customer Type',width=30)
		lbl_customer_type.grid(row=10, column=0, sticky=N + E + W + S)
		lbl_customer_type.rowconfigure(10, weight=1)
		lbl_customer_type.columnconfigure(0, weight=1)

		ent_customer_type = ttk.Entry (topFrame,textvariable=self.cusomer_type)
		ent_customer_type.grid(row=10, column=40, sticky=N + E + W + S)
		ent_customer_type.rowconfigure(10, weight=1)
		ent_customer_type.columnconfigure(40, weight=1)

		lbl_system = ttk.Label(topFrame,text='System',width=30)
		lbl_system.grid(row=20, column=0, sticky=N + E + W + S)
		lbl_system.rowconfigure(20, weight=1)
		lbl_system.columnconfigure(0, weight=1)

		ent_system = ttk.Entry (topFrame,textvariable=self.system)
		ent_system.grid(row=20, column=40, sticky=N + E + W + S)
		ent_system.rowconfigure(20, weight=1)
		ent_system.columnconfigure(40, weight=1)

		lbl_inquiry_category = ttk.Label(topFrame,text='Inquiry Category',width=30)
		lbl_inquiry_category.grid(row=30, column=0, sticky=N + E + W + S)
		lbl_inquiry_category.rowconfigure(30, weight=1)
		lbl_inquiry_category.columnconfigure(0, weight=1)

		ent_inquiry_category = ttk.Entry (topFrame,textvariable=self.inquiry_cat)
		ent_inquiry_category.grid(row=30, column=40, sticky=N + E + W + S)
		ent_inquiry_category.rowconfigure(30, weight=1)
		ent_inquiry_category.columnconfigure(40, weight=1)

		lbl_inquiry = ttk.Label(topFrame,text='Inquiry',width=30)
		lbl_inquiry.grid(row=40, column=0, sticky=N + E + W + S)
		lbl_inquiry.rowconfigure(40, weight=1)
		lbl_inquiry.columnconfigure(0, weight=1)

		ent_inquiry = ttk.Entry (topFrame,textvariable=self.inquiry)
		ent_inquiry.grid(row=40, column=40, sticky=N + E + W + S)
		ent_inquiry.rowconfigure(40, weight=1)
		ent_inquiry.columnconfigure(40, weight=1)

		lbl_status = ttk.Label(topFrame, text='Status', width=30)
		lbl_status.grid(row=50, column=0, sticky=N + E + W + S)
		lbl_status.rowconfigure(50, weight=1)
		lbl_status.columnconfigure(0, weight=1)

		ent_inquiry = ttk.Combobox (topFrame,values=['OPEN','CLOSED','REOPEN'])
		ent_inquiry.grid(row=50, column=40, sticky=N + E + W + S)
		ent_inquiry.rowconfigure(50, weight=1)
		ent_inquiry.columnconfigure(40, weight=1)

		mail_body = Text(topFrame)
		mail_body.grid(row=60, column=0, sticky=N + E + W + S,columnspan=100)
		mail_body.rowconfigure(60, weight=1)
		mail_body.columnconfigure(0, weight=1)
		mail_body.insert(END, mail[2])

		mail_body.config(yscrollcommand=scrollbar.set)
		scrollbar.config(command=mail_body.yview)

		btn_save = Button(topFrame,text='SAVE',width=30)
		btn_save.grid(row=150,column=0)
		btn_cancel = Button(topFrame, text='CANCEL', width=30)
		btn_cancel.grid(row=150,column=40)

	def get_unique_list(self,mail=None):
		if mail is None:
			mail = self.mail_list

		self.unique_mail = []
		prev_id = None
		for m in mail:
			if m[6] != prev_id:
				self.unique_mail.append(m)
			prev_id = m[6]

		self.unique_mail.sort(key=itemgetter(3),reverse=True)


	def initiate_mail_process(self):

		self.threads['mail_extraction'].join()
		original_mail_list = self.q.get()
		print (f'The total mail received {len(original_mail_list)}')
		self.mail_list = sorted(original_mail_list,key=itemgetter(6,3))
		self.get_unique_list()
		total_number_of_mail = len(self.unique_mail)
		print(f'The Unique mail received {len(self.unique_mail)}')

		self.options = [f'{x} - {x + 20}' for x in range(1,total_number_of_mail,20)]
		self.cbo_page_list['values'] = self.options
		self.page.set('1 - 21')

		self.pagination(self.page.get())

	def get_last_run(self,mail_det_conn):
		query = q.query_dict['get_last_run'],[self.SESSION]
		results = mail_det_conn.get_data(query)
		print (f'The results returned {results}. The session data {self.SESSION}')
		if mail_det_conn.status == 'Ok':
			return results[0]
		else:
			return None




root = Tk()
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
root.title("CHUBB TICKET ASSIGNMENT")
app = Application(master=root)
app.mainloop()
