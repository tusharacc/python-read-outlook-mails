import PySimpleGUI as sg

def get_details():
	print ('Got details')

def get_mails():
	print ('Got details')

with sg.FlexForm('MAIL AUTOMATION', auto_size_text=True,auto_size_buttons=True, default_element_size=(30,1)) as form:
	layout = [[sg.SimpleButton('Get Mails', button_color=('white', 'green'),size=(35,1),auto_size_button=False),
			   sg.SimpleButton('Refresh', button_color=('white', 'green'),size=(35,1),auto_size_button=False),
			   sg.SimpleButton('Export to Excel', button_color=('white', 'green'),size=(35,1),auto_size_button=False)],
			  [sg.SimpleButton('Previous', button_color=('white', 'blue'),size=(35,1),auto_size_button=False),
			   sg.InputCombo(['choice 1', 'choice 2'],size=(33,1)),
			   sg.SimpleButton('Next', button_color=('white', 'blue'),size=(35,1),auto_size_button=False)],
			  [sg.Text('FOLDER', size=(10, 1),  text_color='blue'),
			   sg.Text('SUBJECT OF MAIL', size=(30, 1), text_color='blue'),
			   sg.Text('SENDER OF MAIL', size=(15, 1), text_color='blue'),
			   sg.Text('DATE MAIL SENT', size=(15, 1), text_color='blue'),
			   sg.ReadFormButton('DETAILS', button_color=('white', 'blue'),bind_return_key=True)]]
	form.Layout(layout)

while True:
	# Read the form
	button, value = form.Read()
	print (button)
	# Take appropriate action based on button
	if button == 'DETAILS':
		get_details()
	elif button == 'Get Mails':
		get_mails()
	elif button =='Quit' or button is None:
		break