import Tkinter as tk
import ttk
from ttk import *
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from collections import defaultdict
import pandas as pd
import numpy as np
import datetime
import holidays
from bdateutil import relativedelta
from bdateutil import isbday
from icalendar import Calendar, Event, Alarm
import os

# --- functions ---

def search_command(event=None):

    # to compare lower case
    text = e1.get().lower()

    list1.delete(0, 'end')

    if text: # search only if text is not empty
        for word in liststuff:
            if word.lower().startswith(text):
                list1.insert('end', word)


def callback1(selection):
	global b, liststuff, real_test
	
	b=selection
	test=data[(data['Rule Set']==b) & (data['Trigger'])]
	real_test=test['Trigger'].unique()
	liststuff = list(real_test)
	
		

def callback2(evt):
	
	global date_entry_var, date_entry, date_submit, last_day_events, num_days, case_var, case_num
	
	date_entry_var=tk.StringVar()
	date_entry_lb=tk.Label(root, text="Type in a Trigger Date like 01/01/2017")
	date_entry_lb.grid(row=11, column=1)
	date_entry=tk.Entry(root, textvar=date_entry_var)
	date_entry.grid(row=12, column=1)
	case_var=tk.StringVar()
	case_lbl=tk.Label(root, text="Case Name/Matter Number")
	case_num=tk.Entry(root, textvariable=case_var)
	case_lbl.grid(row=13, column=1)
	case_num.grid(row=14, column=1)
	
	date_submit=tk.Button(root, text="Submit Date and Case/Matter Number", command=callback3)
	date_submit.grid(row=15, column=1)
	
	w=evt.widget
	index=int(w.curselection()[0])
	value=w.get(index)
	c=value
	new_test=data[(data['Rule Set']==b) & (data['Trigger']==c)]
	last_day_events=list(new_test['Last day event'].unique())
	
	
	num_days=list(new_test['Number of Days'].unique())
	
def quit_widget():
	root.quit()



def callback4():
	
	
	make_ical_file.config(state="disabled")
	cal=Calendar()
	cal.add('prodid', '-//My calendar product//mxm.dk//')
	cal.add('version', '2.0')
	event = Event()
	last_day_date_selected_begin=last_day_date_selected+"/16/30"
	last_day_date_selected_end=last_day_date_selected+"/17/00"
	event.add('summary', last_day_selected+": "+case_var.get())
	event.add('dtstart', datetime.datetime.strptime(last_day_date_selected_begin, '%m/%d/%Y/%H/%M'))
	event.add('dtend', datetime.datetime.strptime(last_day_date_selected_end, '%m/%d/%Y/%H/%M'))
	alarm= Alarm()
	alarm.add("action", "DISPLAY")
	alarm.add("description", "Reminder")
	alarm.add("TRIGGER;RELATED=START", "-PT{0}H".format(24))
	event.add_component(alarm)
	cal.add_component(event)
	m=os.path.join(os.path.expanduser('~'),'Desktop',"Last Day.ics")
	with open(m, 'wb') as f:
		f.write(cal.to_ical())
		f.close()
	quit_button=tk.Button(root, text="Quit", command=quit_widget)
	quit_button.grid(row=21, column=1)
	

def on_select1(evt):
	global last_day_selected, last_day_date_selected, make_ical_file
	w=evt.widget
	index=int(w.curselection()[0])
	list3.selection_set(index)
	
	last_day_selected=list2.get(index)
	
	last_day_date_selected=list3.get(index)
	
	
	make_ical_file=tk.Button(root, text="Download iCal file", command=callback4)
	make_ical_file.grid(row=20, column=1)
	
	
def on_select2(evt):
	global last_day_selected, last_day_date_selected, make_ical_file
	w=evt.widget
	index=int(w.curselection()[0])
	list2.selection_set(index)
	
	last_day_selected=list2.get(index)
	last_day_date_selected=list3.get(index) 
	
	make_ical_file=tk.Button(root, text="Download iCal file", command=callback4)
	make_ical_file.grid(row=20, column=1)

def callback3():
	global last_day_events, list2, list3
	
	date_entry.config(state='disabled')
	date_submit.config(state='disabled')
	case_num.config(state='disabled')
	new_lbl=tk.Label(root, text="Last Day Event(s): ")
	new_lbl.grid(row=17, column=1)
	
	list2=tk.Listbox(root, height=0, width=75, selectmode='SINGLE')
	list2.bind('<<ListboxSelect>>', on_select1, )
	list2.grid(row=18, column=1)
	list3_lbl=tk.Label(root, text="Last Day Date")
	list3_lbl.grid(row=17, column=2)
	list3=tk.Listbox(root, height=0, width=20, selectmode='SINGLE')
	list3.bind('<<ListboxSelect>>', on_select2)
	list3.grid(row=18, column=2)
	for item in last_day_events:
		list2.insert('end', item)
	
	date=date_entry_var.get()
	us_holidays=holidays.UnitedStates()
	for item in num_days:
		
		if(item<7):
			d = datetime.datetime.strptime(date, '%m/%d/%Y') + relativedelta(bdays=+item)
			if(isbday(d, holidays=holidays.US())==False):
				d=d+relativedelta(bdays=+1)
				list3.insert('end', d.strftime('%m/%d/%Y'))
			elif(isbday(d, holidays=holidays.US())==True):
				list3.insert('end', d.strftime('%m/%d/%Y'))
		elif(item>=7):
			d = datetime.datetime.strptime(date, '%m/%d/%Y') + datetime.timedelta(days=item)
			if(isbday(d, holidays=holidays.US())==False):
				d=d+relativedelta(bdays=+1)
				list3.insert('end', d.strftime('%m/%d/%Y'))
			elif(isbday(d, holidays=holidays.US())==True):
				list3.insert('end', d.strftime('%m/%d/%Y'))

# --- main ---


# GUI

root = tk.Tk()
#root.geometry("1000x1000")
root.title("Last Day iCal Widget")
scope=['https://spreadsheets.google.com/feeds']
creds=ServiceAccountCredentials.from_json_keyfile_name("ical_json.json", scope)
client=gspread.authorize(creds)

sheet = client.open('Legal Calendar Worksheet').sheet1


data=pd.DataFrame(sheet.get_all_records())

data['Trigger ID'].replace('', np.nan, inplace=True)
data.dropna(subset=['Trigger ID'], inplace=True)


OPTIONS=['MN Rules of Civil Procedure', 'MN Rules of Criminal Procedure','MN Rules of Civil Appellate Procedure','Federal Rules of Civil Appellate Procedure','Federal Rules of Civil Procedure'] 	
main_var=tk.StringVar(root)
main_var.set('')
rule_set_drop_down_menu=tk.OptionMenu(root,main_var, *OPTIONS, command=callback1)
lbl=tk.Label(root, text="Choose Rule Set: ")
lbl.grid(row=0, column=1)
rule_set_drop_down_menu.grid(row=1, column=1)		
l1 = tk.Label(root, text='Search Trigger Events')
l1.grid(row=2, column=1)

title_text = tk.StringVar()

e1 = tk.Entry(root, textvariable=title_text)
e1.grid(row=3, column=1)
e1.bind('<KeyRelease>', search_command)
search_result_text=tk.Label(root, text="Search Results")
search_result_text.grid(row=4,column=1)
list1 = tk.Listbox(root, height=0, width=100, exportselection=False)

list1.grid(row=5, rowspan=5, columnspan=2)
list1.bind('<<ListboxSelect>>', callback2)







root.mainloop()