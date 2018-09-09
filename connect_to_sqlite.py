import sqlite3
from sqlite3 import Error

class SQLInterface:
	def __init__(self,filename,ro=False):

		try:
			self.status = 'Ok'
			if ro:
				self.conn = sqlite3.connect('file:./maildb.db?mode=ro',uri=True)
			else:
				self.conn = sqlite3.connect(filename)
		except Error as e:
			self.status = str(e)
			print(e)

	def upsert(self,sel_query,updt_query,ins_query):
		try:
			self.status = 'Ok'
			results = self.get_data(sel_query)

			if self.status == 'Ok':
				if len(results) == 1:
					self.update_data(updt_query)
				elif len(results) == 0:
					self.insert_data(ins_query)
		except Error as e:
			self.status = str(e)
			print(e)

	def update_data(self,query):
		try:
			cur = self.conn.cursor()
			cur.execute(*query)
			self.conn.commit()
			self.status = 'Ok'
		except Error as e:
			self.status = str(e)
			print(e)

	def insert_data(self,query):

		try:
			cur = self.conn.cursor()
			cur.execute(*query)
			self.conn.commit()
			self.status = 'Ok'
		except Error as e:
			self.status = str(e)
			print(e)

	def get_data(self,query):
		try:
			self.status = 'Ok'
			cur = self.conn.cursor()
			if isinstance(query, tuple):
				cur.execute(*query)
			else:
				cur.execute(query)
			return cur.fetchall()
		except Error as e:
			self.status = str(e)


	def close_conn(self):
		self.conn.close()