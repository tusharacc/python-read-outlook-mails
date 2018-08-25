import win32com.client
import win32com
from datetime import datetime,timedelta
import pandas as pd
import re
import cProfile, pstats, io


class ReadOutLookMails:
	def __init__(self):
		self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
		self.accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts

	#@profile
	def read_mails(self,items,mail_filter,folder_name):
		mails = []
		sent_time = re.compile(r'Sent: (.* [AM|PM]+).*')
		myRestrictItems = items.Restrict(mail_filter)
		print(f'The number of items in restrict collection {len(myRestrictItems)}')
		error = []
		# pr = cProfile.Profile()
		# pr.enable()
		for item in myRestrictItems:
			try:
				#body = item.Body
				dates = []
				# for line in body.splitlines():
				#
				# 	find_all = sent_time.findall(line)
				# 	if find_all:
				# 		try:
				# 			dates.append(datetime.strptime(find_all[0], '%A, %B %d, %Y %I:%M %p'))
				# 		except Exception as ex:
				# 			error.append([line, str(ex)])
				if dates:
					min_date = min(dates)
				else:
					min_date = None
				mails.append([folder_name, item.Subject, item.Body, item.ReceivedTime, min_date, item.SenderName,
							  item.ConversationID])
			except AttributeError as ex:
				print(f'Attribute Error {item.Subject}')
		# pr.disable()
		# ps = pstats.Stats(pr)
		# ps.print_stats()
		return mails

	def get_mails(self,start_date_str,end_date_str):
		for account in self.accounts:
			global inbox
			inbox = self.outlook.Folders(account.DeliveryStore.DisplayName)
			folders = inbox.Folders

			mail_filter = "[ReceivedTime] > '{0} 12:00 AM' And [ReceivedTime] < '{1} 12:00 AM' ".format(start_date_str,
																										end_date_str)
			for folder in folders:
				if folder.Name == 'Inbox':
					items = folder.Items
					inbox_mails = self.read_mails(items,mail_filter,'INBOX')
					subfolders = folder.Folders
					for subfolder in subfolders:
						items = subfolder.Items
						subfolder_mails = self.read_mails(items, mail_filter, subfolder.Name.upper())

				elif folder.Name == 'Sent Items':
					items = subfolder.Items
					sent_mails = self.read_mails(items, mail_filter, 'SENT ITEMS')
			return inbox_mails + subfolder_mails + sent_mails
if __name__ == '__main__':
	mails = ReadOutLookMails()
	start_date = (datetime.now() - timedelta(days=30)).strftime('%m/%d/%Y')
	end_date = datetime.now().strftime('%m/%d/%Y')
	mail_list = mails.get_mails(start_date,end_date)

	df = pd.DataFrame(data=mail_list,columns=['Folder Name','Subject','Body','Time Received','First Mail', 'Sender Name','Conversation ID'])
	print (f'The length of dataframe is {len(df)}')
	df.sort_values(['Conversation ID','Time Received'],ascending=[True, False],inplace=True)
	#df['Aging'] = df['Time Received'] - df['First Mail']
	df['Aging'] = df['Time Received'].subtract(df['First Mail']).dt.days
	#err_df = pd.DataFrame(data=error,columns=['Line','Message'])
	df.to_excel('mail.xlsx',index=False)
	#err_df.to_excel('err.xlsx',index=False)

