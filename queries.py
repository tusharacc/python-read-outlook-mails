query_dict = {

	'get_last_run': 'SELECT max(LastRun) from LastRun  ',
	'get_active_users': 'SELECT User,Role FROM Users WHERE Status = \'ACTIVE\'',
	'update_user_assignment' : 'UPDATE MailDetails SET AssignedTo = ? WHERE MailId = ?',
	'search_mail_assignment' : 'SELECT MailId FROM MailDetails WHERE MailId = ?',
	'insert_mail_assignment' : 'INSERT INTO MailDetails(MailId,Subject,Sender,SentDate,AssignedTo) VALUES(?,?,?,?,?)'
}