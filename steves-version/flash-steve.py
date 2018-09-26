from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import os
import xlrd
import json, ast, os, string, random, urllib
import xml.etree.cElementTree as ET
import datetime
import dateparser



window = Tk()
img = PhotoImage (file = 'AS.gif')
imgLb1 = Label (window, image = img)
browseLabel = Label (window, width=30)
fileBtn = Button (window, padx=10, pady=20)
marnel = StringVar()



except_1_variable = StringVar()
except_2_variable = StringVar()
except_3_variable = StringVar()
except_4_variable = StringVar()
except_5_variable = StringVar()
except_6_variable = StringVar()
except_7_variable = StringVar()
except_8_variable = StringVar()
except_9_variable = StringVar()
except_10_variable = StringVar()
except_11_variable = StringVar()
except_12_variable = StringVar()

except_13_variable = StringVar()
except_14_variable = StringVar()
except_15_variable = StringVar()
except_16_variable = StringVar()
except_17_variable = StringVar()
except_18_variable = StringVar()
except_19_variable = StringVar()
except_20_variable = StringVar()


#entry_travel_start = Entry(window)
#entry_start_label = Label(window, width=1, text="Travel Start")

radio_1 = Radiobutton(window, text="Flash Sale ", variable=marnel, value="FlashSale")

except_1_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_1_variable, value="travel_valid_value")
except_1_servicestart = Radiobutton(window, text="Service Start", variable=except_1_variable, value="service_start_value")
except_2_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_2_variable, value="travel_valid_value")
except_2_servicestart = Radiobutton(window, text="Service Start", variable=except_2_variable, value="service_start_value")
except_3_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_3_variable, value="travel_valid_value")
except_3_servicestart = Radiobutton(window, text="Service Start", variable=except_3_variable, value="service_start_value")
except_4_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_4_variable, value="travel_valid_value")
except_4_servicestart = Radiobutton(window, text="Service Start", variable=except_4_variable, value="service_start_value")
except_5_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_5_variable, value="travel_valid_value")
except_5_servicestart = Radiobutton(window, text="Service Start", variable=except_5_variable, value="service_start_value")
except_6_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_6_variable, value="travel_valid_value")
except_6_servicestart = Radiobutton(window, text="Service Start", variable=except_6_variable, value="service_start_value")
except_7_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_7_variable, value="travel_valid_value")
except_7_servicestart = Radiobutton(window, text="Service Start", variable=except_7_variable, value="service_start_value")
except_8_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_8_variable, value="travel_valid_value")
except_8_servicestart = Radiobutton(window, text="Service Start", variable=except_8_variable, value="service_start_value")
except_9_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_9_variable, value="travel_valid_value")
except_9_servicestart = Radiobutton(window, text="Service Start", variable=except_9_variable, value="service_start_value")
except_10_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_10_variable, value="travel_valid_value")
except_10_servicestart = Radiobutton(window, text="Service Start", variable=except_10_variable, value="service_start_value")
except_11_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_11_variable, value="travel_valid_value")
except_11_servicestart = Radiobutton(window, text="Service Start", variable=except_11_variable, value="service_start_value")
except_12_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_12_variable, value="travel_valid_value")
except_12_servicestart = Radiobutton(window, text="Service Start", variable=except_12_variable, value="service_start_value")

except_13_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_12_variable, value="travel_valid_value")
except_13_servicestart = Radiobutton(window, text="Service Start", variable=except_12_variable, value="service_start_value")
except_14_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_12_variable, value="travel_valid_value")
except_14_servicestart = Radiobutton(window, text="Service Start", variable=except_12_variable, value="service_start_value")
except_15_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_12_variable, value="travel_valid_value")
except_15_servicestart = Radiobutton(window, text="Service Start", variable=except_12_variable, value="service_start_value")
except_16_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_12_variable, value="travel_valid_value")
except_16_servicestart = Radiobutton(window, text="Service Start", variable=except_12_variable, value="service_start_value")
except_17_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_12_variable, value="travel_valid_value")
except_17_servicestart = Radiobutton(window, text="Service Start", variable=except_12_variable, value="service_start_value")
except_18_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_12_variable, value="travel_valid_value")
except_18_servicestart = Radiobutton(window, text="Service Start", variable=except_12_variable, value="service_start_value")
except_19_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_12_variable, value="travel_valid_value")
except_19_servicestart = Radiobutton(window, text="Service Start", variable=except_12_variable, value="service_start_value")
except_20_travelvalid = Radiobutton(window, text="Travel Valid", variable=except_12_variable, value="travel_valid_value")
except_20_servicestart = Radiobutton(window, text="Service Start", variable=except_12_variable, value="service_start_value")


weekly_label = Label(window, text="Flash Sale Type", width=40)


exception_codes_label1 = Label(window, text="EXCEPTIONS:", width=10)
exception_codes_label2 = Label(window, text="Codes", width=8)
exception_codes_label3 = Label(window, text="Days of Travel", width=25)
exception_codes_label4 = Label(window, text="startdate", width=8)
exception_codes_label5 = Label(window, text="enddate", width=8)


example_label = Label(window, text="SERVICE STARTS:", width=13)
example_label2 = Label(window, text="TRAVEL IS VALID:", width=13)
example_codes_label = Label(window, text="MCOSFO", width=8)
example_days_label = Label(window, text="Thursday through Monday", width=25)
example_start_label = Label(window, text="2018-09-22", width=8)
example_end_label = Label(window, text="2018-09-30", width=8)


example_label.configure(fg="gray")
example_label2.configure(fg="gray")
example_codes_label.configure(fg="gray")
example_days_label.configure(fg="gray")
example_start_label.configure(fg="gray")
example_end_label.configure(fg="gray")


weekly_label.configure(fg="gray")



exception_codes1 = Entry(window, width=8)
exception_days1 = Entry(window, width=25)
exception_start1 = Entry(window, width=8)
exception_end1 = Entry(window, width=8)

exception_codes2 = Entry(window, width=8)
exception_days2 = Entry(window, width=25)
exception_start2 = Entry(window, width=8)
exception_end2 = Entry(window, width=8)


exception_codes3 = Entry(window, width=8)
exception_days3 = Entry(window, width=25)
exception_start3 = Entry(window, width=8)
exception_end3 = Entry(window, width=8)

exception_codes4 = Entry(window, width=8)
exception_days4 = Entry(window, width=25)
exception_start4 = Entry(window, width=8)
exception_end4 = Entry(window, width=8)


exception_codes5 = Entry(window, width=8)
exception_days5 = Entry(window, width=25)
exception_start5 = Entry(window, width=8)
exception_end5 = Entry(window, width=8)


exception_codes6 = Entry(window, width=8)
exception_days6 = Entry(window, width=25)
exception_start6 = Entry(window, width=8)
exception_end6 = Entry(window, width=8)


exception_codes7 = Entry(window, width=8)
exception_days7 = Entry(window, width=25)
exception_start7 = Entry(window, width=8)
exception_end7 = Entry(window, width=8)

exception_codes8 = Entry(window, width=8)
exception_days8 = Entry(window, width=25)
exception_start8 = Entry(window, width=8)
exception_end8 = Entry(window, width=8)

exception_codes9 = Entry(window, width=8)
exception_days9 = Entry(window, width=25)
exception_start9 = Entry(window, width=8)
exception_end9 = Entry(window, width=8)

exception_codes10 = Entry(window, width=8)
exception_days10 = Entry(window, width=25)
exception_start10 = Entry(window, width=8)
exception_end10 = Entry(window, width=8)

exception_codes11 = Entry(window, width=8)
exception_days11 = Entry(window, width=25)
exception_start11 = Entry(window, width=8)
exception_end11 = Entry(window, width=8)

exception_codes12 = Entry(window, width=8)
exception_days12 = Entry(window, width=25)
exception_start12 = Entry(window, width=8)
exception_end12 = Entry(window, width=8)

exception_codes13 = Entry(window, width=8)
exception_days13 = Entry(window, width=25)
exception_start13 = Entry(window, width=8)
exception_end13 = Entry(window, width=8)

exception_codes14 = Entry(window, width=8)
exception_days14 = Entry(window, width=25)
exception_start14 = Entry(window, width=8)
exception_end14 = Entry(window, width=8)

exception_codes15 = Entry(window, width=8)
exception_days15 = Entry(window, width=25)
exception_start15 = Entry(window, width=8)
exception_end15 = Entry(window, width=8)

exception_codes16 = Entry(window, width=8)
exception_days16 = Entry(window, width=25)
exception_start16 = Entry(window, width=8)
exception_end16 = Entry(window, width=8)

exception_codes17 = Entry(window, width=8)
exception_days17 = Entry(window, width=25)
exception_start17 = Entry(window, width=8)
exception_end17 = Entry(window, width=8)

exception_codes18 = Entry(window, width=8)
exception_days18 = Entry(window, width=25)
exception_start18 = Entry(window, width=8)
exception_end18 = Entry(window, width=8)

exception_codes19 = Entry(window, width=8)
exception_days19 = Entry(window, width=25)
exception_start19 = Entry(window, width=8)
exception_end19 = Entry(window, width=8)

exception_codes20 = Entry(window, width=8)
exception_days20 = Entry(window, width=25)
exception_start20 = Entry(window, width=8)
exception_end20 = Entry(window, width=8)


runBtn = Button (window, padx=10, pady=20)
resBtn = Button (window, padx=20, pady=20)


resBtn.configure(fg="red",bg="red")
runBtn.configure(fg="green",bg="green")

radio_1.select()

imgLb1.grid(row=1, column=1, rowspan=1, columnspan = 6)
browseLabel.grid(row=3, column=1)

fileBtn.grid(row=3, column=2, columnspan = 1)

radio_1.grid(row=4, column=1, columnspan = 1)

#weekly_label.grid(row=4, column=2, columnspan = 2)


except_1_travelvalid.select()
except_1_travelvalid.grid(row=11, column=1, columnspan = 1)
except_1_servicestart.grid(row=11, column=2, columnspan = 1)

except_2_travelvalid.select()
except_2_travelvalid.grid(row=12, column=1, columnspan = 1)
except_2_servicestart.grid(row=12, column=2, columnspan = 1)

except_3_travelvalid.select()
except_3_travelvalid.grid(row=13, column=1, columnspan = 1)
except_3_servicestart.grid(row=13, column=2, columnspan = 1)

except_4_travelvalid.select()
except_4_travelvalid.grid(row=14, column=1, columnspan = 1)
except_4_servicestart.grid(row=14, column=2, columnspan = 1)

except_5_travelvalid.select()
except_5_travelvalid.grid(row=15, column=1, columnspan = 1)
except_5_servicestart.grid(row=15, column=2, columnspan = 1)

except_6_travelvalid.select()
except_6_travelvalid.grid(row=16, column=1, columnspan = 1)
except_6_servicestart.grid(row=16, column=2, columnspan = 1)

except_7_travelvalid.select()
except_7_travelvalid.grid(row=17, column=1, columnspan = 1)
except_7_servicestart.grid(row=17, column=2, columnspan = 1)

except_8_travelvalid.select()
except_8_travelvalid.grid(row=18, column=1, columnspan = 1)
except_8_servicestart.grid(row=18, column=2, columnspan = 1)

except_9_travelvalid.select()
except_9_travelvalid.grid(row=19, column=1, columnspan = 1)
except_9_servicestart.grid(row=19, column=2, columnspan = 1)

except_10_travelvalid.select()
except_10_travelvalid.grid(row=20, column=1, columnspan = 1)
except_10_servicestart.grid(row=20, column=2, columnspan = 1)

except_11_travelvalid.select()
except_11_travelvalid.grid(row=21, column=1, columnspan = 1)
except_11_servicestart.grid(row=21, column=2, columnspan = 1)

except_12_travelvalid.select()
except_12_travelvalid.grid(row=22, column=1, columnspan = 1)
except_12_servicestart.grid(row=22, column=2, columnspan = 1)


except_13_travelvalid.select()
except_13_travelvalid.grid(row=23, column=1, columnspan = 1)
except_13_servicestart.grid(row=23, column=2, columnspan = 1)

except_14_travelvalid.select()
except_14_travelvalid.grid(row=24, column=1, columnspan = 1)
except_14_servicestart.grid(row=24, column=2, columnspan = 1)

except_15_travelvalid.select()
except_15_travelvalid.grid(row=25, column=1, columnspan = 1)
except_15_servicestart.grid(row=25, column=2, columnspan = 1)

except_16_travelvalid.select()
except_16_travelvalid.grid(row=26, column=1, columnspan = 1)
except_16_servicestart.grid(row=26, column=2, columnspan = 1)

except_17_travelvalid.select()
except_17_travelvalid.grid(row=27, column=1, columnspan = 1)
except_17_servicestart.grid(row=27, column=2, columnspan = 1)

except_18_travelvalid.select()
except_18_travelvalid.grid(row=28, column=1, columnspan = 1)
except_18_servicestart.grid(row=28, column=2, columnspan = 1)

except_19_travelvalid.select()
except_19_travelvalid.grid(row=29, column=1, columnspan = 1)
except_19_servicestart.grid(row=29, column=2, columnspan = 1)

except_20_travelvalid.select()
except_20_travelvalid.grid(row=30, column=1, columnspan = 1)
except_20_servicestart.grid(row=30, column=2, columnspan = 1)


#exception_codes_label1.grid(row=11, column=1)

exception_codes_label2.grid(row=9, column=3)
exception_codes_label3.grid(row=9, column=4)
exception_codes_label4.grid(row=9, column=5)
exception_codes_label5.grid(row=9, column=6)


example_label.grid(row=10, column=2)
example_label2.grid(row=10, column=1)
example_codes_label.grid(row=10, column=3)
example_days_label.grid(row=10, column=4)
example_start_label.grid(row=10, column=5)
example_end_label.grid(row=10, column=6)


exception_codes1.grid(row=11, column=3)
exception_days1.grid(row=11, column=4)
exception_start1.grid(row=11, column=5)
exception_end1.grid(row=11, column=6)

exception_codes2.grid(row=12, column=3)
exception_days2.grid(row=12, column=4)
exception_start2.grid(row=12, column=5)
exception_end2.grid(row=12, column=6)

exception_codes3.grid(row=13, column=3)
exception_days3.grid(row=13, column=4)
exception_start3.grid(row=13, column=5)
exception_end3.grid(row=13, column=6)

exception_codes4.grid(row=14, column=3)
exception_days4.grid(row=14, column=4)
exception_start4.grid(row=14, column=5)
exception_end4.grid(row=14, column=6)


exception_codes5.grid(row=15, column=3)
exception_days5.grid(row=15, column=4)
exception_start5.grid(row=15, column=5)
exception_end5.grid(row=15, column=6)


exception_codes6.grid(row=16, column=3)
exception_days6.grid(row=16, column=4)
exception_start6.grid(row=16, column=5)
exception_end6.grid(row=16, column=6)


exception_codes7.grid(row=17, column=3)
exception_days7.grid(row=17, column=4)
exception_start7.grid(row=17, column=5)
exception_end7.grid(row=17, column=6)

exception_codes8.grid(row=18, column=3)
exception_days8.grid(row=18, column=4)
exception_start8.grid(row=18, column=5)
exception_end8.grid(row=18, column=6)


exception_codes9.grid(row=19, column=3)
exception_days9.grid(row=19, column=4)
exception_start9.grid(row=19, column=5)
exception_end9.grid(row=19, column=6)

exception_codes10.grid(row=20, column=3)
exception_days10.grid(row=20, column=4)
exception_start10.grid(row=20, column=5)
exception_end10.grid(row=20, column=6)

exception_codes11.grid(row=21, column=3)
exception_days11.grid(row=21, column=4)
exception_start11.grid(row=21, column=5)
exception_end11.grid(row=21, column=6)

exception_codes12.grid(row=22, column=3)
exception_days12.grid(row=22, column=4)
exception_start12.grid(row=22, column=5)
exception_end12.grid(row=22, column=6)

exception_codes13.grid(row=23, column=3)
exception_days13.grid(row=23, column=4)
exception_start13.grid(row=23, column=5)
exception_end13.grid(row=23, column=6)

exception_codes14.grid(row=24, column=3)
exception_days14.grid(row=24, column=4)
exception_start14.grid(row=24, column=5)
exception_end14.grid(row=24, column=6)

exception_codes15.grid(row=25, column=3)
exception_days15.grid(row=25, column=4)
exception_start15.grid(row=25, column=5)
exception_end15.grid(row=25, column=6)

exception_codes16.grid(row=26, column=3)
exception_days16.grid(row=26, column=4)
exception_start16.grid(row=26, column=5)
exception_end16.grid(row=26, column=6)

exception_codes17.grid(row=27, column=3)
exception_days17.grid(row=27, column=4)
exception_start17.grid(row=27, column=5)
exception_end17.grid(row=27, column=6)

exception_codes18.grid(row=28, column=3)
exception_days18.grid(row=28, column=4)
exception_start18.grid(row=28, column=5)
exception_end18.grid(row=28, column=6)

exception_codes19.grid(row=29, column=3)
exception_days19.grid(row=29, column=4)
exception_start19.grid(row=29, column=5)
exception_end19.grid(row=29, column=6)

exception_codes20.grid(row=30, column=3)
exception_days20.grid(row=30, column=4)
exception_start20.grid(row=30, column=5)
exception_end20.grid(row=30, column=6)

runBtn.grid(row=1, column=5, columnspan = 1)
resBtn.grid(row=1, column=4, columnspan = 1)


window.title('FLASH SALE')
window.resizable(0,0)
browseLabel.configure(text='Choose a file ....')
fileBtn.configure(text='Browse')
runBtn.configure(text='Run Program', state=DISABLED)
resBtn.configure(text='Reset ')

#wrkdirectory = '/Users/mmangruban/Desktop/github/tkinter-gui.py/data-to-read'

#Function to run when the Browse button gets clicked
def getfile() :
    #/Users/mmangruban/Desktop/github/tkinter-gui.py/data-to-read
    #//seavvfile1/Market_SAIntMktg/_Offers/5. In Work/WeeklyFlightDeals/temp/testing
    #window.fileName =  tkFileDialog.askopenfilename(initialdir = wrkdirectory,title = "Select file",filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))
    window.fileName =  filedialog.askopenfilename(title = "Select file",filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))
    #window.fileName = tkFileDialog.askopenfilename(filetypes = (("Excel files", "*.xlsx"), ("All files", "*.*")))
    path, filename = os.path.split(window.fileName)
    if len(window.fileName) > 0:
        browseLabel.configure(text="../"+filename)
        runBtn.configure(state = NORMAL)
        resBtn.configure(state=NORMAL)
        return window.fileName



#Function to run when the Reset button gets clicked
def reset() :
    browseLabel.configure(text="Select a file ....")
    window.fileName = ''
    exception_codes1.delete(0, 'end')
    exception_codes2.delete(0, 'end')
    exception_codes3.delete(0, 'end')
    exception_codes4.delete(0, 'end')
    exception_codes5.delete(0, 'end')
    exception_codes6.delete(0, 'end')
    exception_codes7.delete(0, 'end')
    exception_codes8.delete(0, 'end')
    exception_codes9.delete(0, 'end')
    exception_codes10.delete(0, 'end')
    exception_codes11.delete(0, 'end')
    exception_codes12.delete(0, 'end')
    exception_codes13.delete(0, 'end')
    exception_codes14.delete(0, 'end')
    exception_codes15.delete(0, 'end')
    exception_codes16.delete(0, 'end')
    exception_codes17.delete(0, 'end')
    exception_codes18.delete(0, 'end')
    exception_codes19.delete(0, 'end')
    exception_codes20.delete(0, 'end')
    exception_days1.delete(0, 'end')
    exception_days2.delete(0, 'end')
    exception_days3.delete(0, 'end')
    exception_days4.delete(0, 'end')
    exception_days5.delete(0, 'end')
    exception_days6.delete(0, 'end')
    exception_days7.delete(0, 'end')
    exception_days8.delete(0, 'end')
    exception_days9.delete(0, 'end')
    exception_days10.delete(0, 'end')
    exception_days11.delete(0, 'end')
    exception_days12.delete(0, 'end')
    exception_days13.delete(0, 'end')
    exception_days14.delete(0, 'end')
    exception_days15.delete(0, 'end')
    exception_days16.delete(0, 'end')
    exception_days17.delete(0, 'end')
    exception_days18.delete(0, 'end')
    exception_days19.delete(0, 'end')
    exception_days20.delete(0, 'end')
    fileBtn.configure(state = NORMAL)
    resBtn.configure(state=DISABLED)  
    runBtn.configure(state = DISABLED)
    radio_1.select()



#THIS FUNCTION WILL DETERMINE THE NEXT TUESDAY THAT IS COMING UP SO THAT IT CAN CREATE THE FILE NAME STRUCTURE
def coming_tuesday(d, weekday):
    days_ahead = weekday - d.weekday()
    if days_ahead <= 0: # Target day already happened this week
        days_ahead += 7
    return d + datetime.timedelta(days_ahead)



def find_two_tuesday(d, weekday, span):
    days_ahead = weekday - d.weekday()
    if days_ahead <= 0: # Target day already happened this week
        days_ahead += span
    return d + datetime.timedelta(days_ahead)






def automate() :
    print(window.fileName)
    print(marnel.get())
    runBtn.configure(state = DISABLED)
    resBtn.configure(state=NORMAL)



    def f(x):
        return {
            1: 'January',
            2: 'February',
            3: 'March',
            4: 'April',
            5: 'May',
            6: 'June',
            7: 'July',
            8: 'August',
            9: 'September',
            10: 'October',
            11: 'November',
            12: 'December',
        }[x]

    def m(y):
        return {
            'ABQ' : 'Albuquerque',
            'ACV' : 'Eureka',
            'ADK' : 'Adak Island',
            'ADQ' : 'Kodiak',
            'AKN' : 'King Salmon',
            'ALW' : 'Walla Walla',
            'ANC' : 'Anchorage',
            'ATL' : 'Atlanta',
            'AUS' : 'Austin',
            'BET' : 'Bethel',
            'BIL' : 'Billings',
            'BLI' : 'Bellingham',
            'BNA' : 'Nashville',
            'BOI' : 'Boise',
            'BOS' : 'Boston',
            'BRW' : 'Barrow',
            'BUR' : 'Burbank',
            'BWI' : 'Baltimore',
            'BZN' : 'Bozeman',
            'CDV' : 'Cordova',
            'CHS' : 'Charleston',
            'COS' : 'Colorado Springs',
            'CUN' : 'Cancun',
            'DCA' : 'Washington - Reagan',
            'DEN' : 'Denver',
            'DFW' : 'Dallas',
            'DLG' : 'Dillingham',
            'DTW' : 'Detroit',
            'DUT' : 'Dutch Harbor',
            'EAT' : 'Wenatchee',
            'EUG' : 'Eugene',
            'EWR' : 'New York - Newark',
            'FAI' : 'Fairbanks',
            'FAT' : 'Fresno',
            'FCA' : 'Kalispell',
            'FLG' : 'Flagstaff',
            'FLL' : 'Ft Lauderdale',
            'GDL' : 'Guadalajara',
            'GEG' : 'Spokane',
            'GST' : 'Glacier Bay',
            'GST' : 'Gustavus',
            'GTF' : 'Great Falls',
            'GUC' : 'Gunnison County / Crested Butte',
            'HDN' : 'Steamboat Springs',
            'HLN' : 'Helena',
            'HNL' : 'Honolulu',
            'IAD' : 'Washington - Dulles',
            'IAH' : 'Houston',
            'IDA' : 'Idaho Falls',
            'JFK' : 'New York - JFK',
            'JNU' : 'Juneau',
            'KOA' : 'Kona',
            'KTN' : 'Ketchikan',
            'LAP' : 'La Paz',
            'LAS' : 'Las Vegas',
            'LAX' : 'Los Angeles',
            'LGB' : 'Long Beach',
            'LIH' : 'Kauai',
            'LIR' : 'Liberia, Costa Rica',
            'LTO' : 'Loreto',
            'MCI' : 'Kansas City',
            'MCO' : 'Orlando',
            'MEX' : 'Mexico City',
            'MFR' : 'Medford',
            'MIA' : 'Miami',
            'MMH' : 'Mammoth Lakes',
            'MRY' : 'Monterey',
            'MSO' : 'Missoula',
            'MSP' : 'Minneapolis',
            'MSY' : 'New Orleans',
            'MZT' : 'Mazatlan',
            'OAK' : 'Oakland',
            'OGG' : 'Maui',
            'OKC' : 'Oklahoma City',
            'OMA' : 'Omaha',
            'OME' : 'Nome',
            'ONT' : 'Ontario',
            'ORD' : 'Chicago',
            'OTZ' : 'Kotzebue',
            'PDX' : 'Portland',
            'PHL' : 'Philadelphia',
            'PHX' : 'Phoenix',
            'PRC' : 'Prescott',
            'PSC' : 'Pasco',
            'PSG' : 'Petersburg',
            'PSP' : 'Palm Springs',
            'PUW' : 'Pullman',
            'PVR' : 'Puerto Vallarta',
            'RDD' : 'Redding',
            'RDM' : 'Redmond',
            'RDU' : 'Raleigh-Durham',
            'RNO' : 'Reno',
            'SAN' : 'San Diego',
            'SAT' : 'San Antonio',
            'SBA' : 'Santa Barbara',
            'SCC' : 'Prudhoe Bay',
            'SEA' : 'Seattle',
            'SFO' : 'San Francisco',
            'SIT' : 'Sitka',
            'SJC' : 'San Jose',
            'SJD' : 'Los Cabos',
            'SJO' : 'San Jose, Costa Rica',
            'SLC' : 'Salt Lake City',
            'SMF' : 'Sacramento',
            'SNA' : 'Orange County',
            'STL' : 'St Louis',
            'STS' : 'Santa Rosa',
            'SUN' : 'Sun Valley',
            'TPA' : 'Tampa',
            'TUS' : 'Tucson',
            'WRG' : 'Wrangell',
            'YAK' : 'Yakutat',
            'YEG' : 'Edmonton',
            'YKM' : 'Yakima',
            'YLW' : 'Kelowna',
            'YVR' : 'Vancouver',
            'YYC' : 'Calgary',
            'YYJ' : 'Victoria',
            'ZIH' : 'Ixtapa',
            'ZLO' : 'Manzanillo',
            'MKE' : 'Milwaukee',
            'HAV' : 'Havana',
            'ICT' : 'Wichita',
            'IND' : 'Indianapolis',
            'SBP' : 'San Luis Obispo',
            'DAL' : 'Dallas - Love Field',
            'LGA' : 'New York - LaGuardia',
            'IAD' : 'Washington D.C. - Dulles',
            'PIT' : 'Pittsburgh',
        }[y]


    def changeDaysFont(x):
        set_val = {
            'Sunday, Monday, Tuesday': 'Sunday through Tuesday',
            'Monday, Tuesday, Wednesday, Thursday, Saturday': 'Monday through Thursday and Saturday',
        }
        for key in set_val.keys():
            if key == x:
                return set_val[key]


    def getYear(this_date):
        value_int = xlrd.xldate_as_tuple(int(this_date), 0)
        parsed_date = datetime.date(value_int[0], value_int[1], value_int[2])
        my_year = str(parsed_date)
        my_year = my_year.split("-",1)[0]
        print(my_year)
        return int(my_year)



    def getMonth(this_date):
        value_int = xlrd.xldate_as_tuple(int(this_date), 0)
        parsed_date = datetime.date(value_int[0], value_int[1], value_int[2])
        my_month = str(parsed_date)
        my_month = my_month.split("-",2)[1]
        print(my_month)
        return int(my_month)


    def getDay(this_date):
        value_int = xlrd.xldate_as_tuple(int(this_date), 0)
        parsed_date = datetime.date(value_int[0], value_int[1], value_int[2])
        my_day = str(parsed_date)
        my_day = my_day.split("-",3)[2]
        print(my_day)
        return int(my_day)




    def parseDates(this_date):
        value_int = xlrd.xldate_as_tuple(int(this_date), 0)
        parsed_date = datetime.date(value_int[0], value_int[1], value_int[2])
        return parsed_date




    def dateInEnglish(readable_date):
        value_int = xlrd.xldate_as_tuple(int(readable_date), 0)
        parsed_date = datetime.date(value_int[0], value_int[1], value_int[2])
        return f(value_int[1])+" "+str(value_int[2])+ ", "+ str(value_int[0])




    def getStringCoordinates(string_to_search_for):
        for row_index in xrange(1, sheet_one.nrows):
            if sheet_one.cell(row_index, 0).value.strip() == string_to_search_for:
                return row_index+1
            

    def getValueToTheRightOfString(string_to_search_for):
        for row_index in xrange(1, getStringCoordinates(string_to_search_for)):
            if sheet_one.cell(row_index, 0).value.strip() == string_to_search_for:
                if sheet_one.cell(row_index, 1).value:
                    pulled_date_number = sheet_one.cell(row_index, 1).value
                    return pulled_date_number


    def getSpecificTravelStartDates(string_to_search_for):
        for row_index in xrange(getStringCoordinates("Travel Start:"), getStringCoordinates("Complete Travel By:")):
            if sheet_one.cell(row_index, 0).value.strip() == string_to_search_for:
                if sheet_one.cell(row_index, 1).value:
                    pulled_date_number = sheet_one.cell(row_index, 1).value
                    return pulled_date_number


    def getSpecificTravelEndDates(string_to_search_for):
        for row_index in xrange(getStringCoordinates("Complete Travel By:"), getStringCoordinates("Advance Purchase:")):
            if sheet_one.cell(row_index, 0).value.strip() == string_to_search_for:
                if sheet_one.cell(row_index, 1).value:
                    pulled_date_number = sheet_one.cell(row_index, 1).value
                    return pulled_date_number                




    def getBlackoutDates(string_to_search_for):
        for row_index in xrange(getStringCoordinates("Blackouts:"), getStringCoordinates("Service Exceptions:")):
            if sheet_one.cell(row_index, 0).value.strip() == string_to_search_for:
                if sheet_one.cell(row_index, 1).value:
                    pulled_date_number = sheet_one.cell(row_index, 1).value
                    return pulled_date_number   


    def getTravelStart(string_to_search_for):
        for row_index in xrange(1, getStringCoordinates("Complete Travel By:")):
            if sheet_one.cell(row_index, 0).value.strip() == string_to_search_for:
                if sheet_one.cell(row_index, 1).value:
                    pulled_date_number = sheet_one.cell(row_index, 1).value
                    return pulled_date_number


    def getTravelEnd(string_to_search_for):
        for row_index in xrange(getStringCoordinates("Complete Travel By:"), getStringCoordinates("Advance Purchase:")):
            if sheet_one.cell(row_index, 0).value.strip() == string_to_search_for:
                if sheet_one.cell(row_index, 1).value:
                    pulled_date_number = sheet_one.cell(row_index, 1).value
                    return pulled_date_number


    def getAvailability(string_to_search_for):
        for row_index in xrange(getStringCoordinates("Advance Purchase:")+1, 53):
            if sheet_one.cell(row_index, 0).value.strip() == string_to_search_for:
                if sheet_one.cell(row_index, 1).value:
                    pulled_date_number = sheet_one.cell(row_index, 1).value
                    return pulled_date_number        


    #Check to see if there are any Fares with International DEPARTURES
    def internationalDepartureCheck(airline_type):
        international_codes = ["MEX","CUN","GDL","LTO","SJD","ZLO","MZT","PVR","ZIH","LIR","SJO","HAV"]
        list_of_violating_departures = []
        for col in range(5,7):
            for row in range(1, sheet.nrows):
                if sheet.cell_value(row, 7) in international_codes:
                    list_of_violating_departures.append(row)
                else:
                    all_violation = list_of_violating_departures
            return all_violation


    def removeDuplicates(original_list, total_exceptions_list):
        i = 0
        while i < len(total_exceptions_list):
            if total_exceptions_list[i] in original_list:
                original_list.remove(total_exceptions_list[i])
            i+=1
        return original_list   


    def serviceExceptionFares(airline_type, origin, destination):
        exception_list = []
        for col in range(5,7):
            for row in range(2, sheet.nrows):
                if sheet.cell_value(row, 5) == airline_type:
                    if sheet.cell_value(row, 7) == origin and sheet.cell_value(row, 9) == destination:
                        exception_list.append(row)
                    else:
                        continue
                else:
                    my_exception_fares = exception_list
                    continue
            return my_exception_fares


    def alaskaToFromHawaiiFares(airline_type, total_exceptions_list):
        just_alaska_codes = ["ADK","ANC","BRW","BET","CDV","DLG","DUT","FAI","GST","JNU","KTN","AKN","ADQ","OTZ","OME","PSG","SCC","SIT","WRG","YAK"]
        just_hawaii_codes = ["OGG","LIH","KOA","HNL"]
        alaska_hawaii_list = []
        for col in range(5,7):
            for row in range(2, sheet.nrows):
                if sheet.cell_value(row, 5) == airline_type:
                    if sheet.cell_value(row, 7) in just_hawaii_codes or sheet.cell_value(row, 9) in just_hawaii_codes:
                        if sheet.cell_value(row, 7) in just_alaska_codes or sheet.cell_value(row, 9) in just_alaska_codes:
                            alaska_hawaii_list.append(row)
                        else:
                            continue
                    else:
                        continue
                else:
                    my_alaska_hawaii_fares = alaska_hawaii_list
                    continue
            my_alaska_hawaii_fares = removeDuplicates(my_alaska_hawaii_fares, total_exceptions_list)
            return my_alaska_hawaii_fares


   
    def hawaiiFares(airline_type, total_exceptions_list):
        alaska_codes = ["ADK","ANC","BRW","BET","CDV","DLG","DUT","FAI","GST","JNU","KTN","AKN","ADQ","OTZ","OME","PSG","SCC","SIT","WRG","YAK"]
        hawaii_codes = ["OGG","LIH","KOA","HNL"]
        hawaii_list = []
        for col in range(5,7):
            for row in range(2, sheet.nrows):
                if sheet.cell_value(row, 5) == airline_type:
                    if sheet.cell_value(row, 7) in hawaii_codes or sheet.cell_value(row, 9) in hawaii_codes:
                        if sheet.cell_value(row, 7) in alaska_codes or sheet.cell_value(row, 9) in alaska_codes:
                            continue
                        else:
                            hawaii_list.append(row)
                    else:
                        continue
                else:
                    my_hawaii_fares = hawaii_list
                    continue
            my_hawaii_fares = removeDuplicates(my_hawaii_fares, total_exceptions_list)
            return my_hawaii_fares




    def floridaFares(airline_type, total_exceptions_list):
        florida_codes = ["FLL","MCO","MIA","TPA"]
        florida_list = []
        for col in range(5,7):
            for row in range(2, sheet.nrows):
                if sheet.cell_value(row, 5) == airline_type:
                    if sheet.cell_value(row, 7) in florida_codes or sheet.cell_value(row, 9) in florida_codes:
                        florida_list.append(row)
                    else:
                        continue
                else:
                    my_florida_fares = florida_list
                    continue
            my_florida_fares = removeDuplicates(my_florida_fares, total_exceptions_list)
            return my_florida_fares




        


    def mexicoFares(airline_type, total_exceptions_list):
        mexico_codes = ["MEX","CUN","GDL","LTO","SJD","ZLO","MZT","PVR","ZIH"]
        mexico_list = []
        for col in range(5,7):
            for row in range(2, sheet.nrows):
                if sheet.cell_value(row, 5) == airline_type:
                    if sheet.cell_value(row, 9) in mexico_codes:
                        mexico_list.append(row)
                    else:
                        continue
                else:
                    my_mexico_fares = mexico_list
                    continue
            my_mexico_fares = removeDuplicates(my_mexico_fares, total_exceptions_list)
            return my_mexico_fares

    def costaricaFares(airline_type, total_exceptions_list):
        costarica_codes = ["LIR","SJO","HAV"]
        costarica_list = []
        for col in range(5,7):
            for row in range(2, sheet.nrows):
                if sheet.cell_value(row, 5) == airline_type:
                    if sheet.cell_value(row, 9) in costarica_codes:
                        costarica_list.append(row)
                    else:
                        continue
                else:
                    my_costarica_fares = costarica_list
                    continue
            my_costarica_fares = removeDuplicates(my_costarica_fares, total_exceptions_list)
            return my_costarica_fares



    #Saves and returns LIST of NON-HAWAII/VIRGIN or ALASKA Rows depending on the passed parameter of 'AS' or 'VX'
    def allOtherRows(airline_type, total_exceptions_list):
        combined_hawaii_and_international = ["FLL","MCO","MIA","TPA","OGG","HNL","LIH","KOA","MEX","CUN","GDL","LTO","SJD","ZLO","MZT","PVR","ZIH","LIR","SJO","HAV"]
        others_list = []
        for col in range(5,7):
            for row in range(1, sheet.nrows):
                if sheet.cell_value(row, 5) == airline_type:
                    if sheet.cell_value(row, 7) not in combined_hawaii_and_international and sheet.cell_value(row, 9) not in combined_hawaii_and_international:
                        others_list.append(row)
                    else:
                        continue
                else:
                    all_other_fares = others_list
                    continue
            all_other_fares = removeDuplicates(all_other_fares, total_exceptions_list)
            return all_other_fares





    def sortkeypicker(keynames):
        negate = set()
        for i, k in enumerate(keynames):
            if k[:1] == '-':
                keynames[i] = k[1:]
                negate.add(k[1:])
        def getit(adict):
            composite = [adict[k] for k in keynames]
            for i, (k, v) in enumerate(zip(keynames, composite)):
                if k in negate:
                    composite[i] = -v
            return composite
        return getit





    #This is pulling all fares with the green background and creates then returns a list of dictionaries
    def pullFaresAndSaveInList(list_being_passed):
        #This sets the name of all keys for the list of dictionary  
        keys = ["oCode","oCity","dCode","dCity","fare"]
        my_dictionary_list = []
        # this selects how many rows to read
        for row in range(1, sheet.nrows):
            if row in list_being_passed:
                my_dictionary_list.append({keys[0]: sheet.cell(row, 7).value,keys[1]: sheet.cell(row, 8).value,keys[2]: sheet.cell(row, 9).value,keys[3]: sheet.cell(row, 10).value,keys[4]: int(sheet.cell(row, 11).value)})
        # saves the list into a variable
        #my_fares = sorted(my_dictionary_list, key=itemgetter('fare'), key=itemgetter('oCity'), key=itemgetter('dCity'))
        my_fares = sorted(my_dictionary_list, key=sortkeypicker(['fare', 'oCity', 'dCity']))
        return my_fares

    





    tree = ET.parse('flash-steve.xml')
    root = tree.getroot()  # now get the root
    root.attrib['xmlns:ss']="urn:schemas-microsoft-com:office:spreadsheet"



    #CREATE FLASH DEALSETS
    def flashDealSet(which_rows, advance_purchase, upper_or_lower):
        
        dealset = ET.SubElement(root, "DealSet")
        dealset.attrib['from']= str(parseDates(getValueToTheRightOfString("Sale Start Date:")))+'T00:00:01'
        dealset.attrib['to']= str(parseDates(getValueToTheRightOfString("Purchase By:")))+'T23:59:59'


        dealinfo = ET.SubElement(dealset, "DealInfo")
        dealinfo.attrib['url']=''
        dealinfo.attrib['dealType']='Standard' #MileagePlan || Standard || Saver
        makecode = str(parseDates(getValueToTheRightOfString("Sale Start Date:"))).replace('-', '')

        if upper_or_lower == 'alaskahawaii':
            dealinfo.attrib['code']=makecode+'_SALE_AS_AKHI'

        if upper_or_lower == 'hawaii':
            dealinfo.attrib['code']=makecode+'_SALE_AS-HI'
        
        if upper_or_lower == 'mexico':
            dealinfo.attrib['code']=makecode+'_SALE_AS-MX'

        if upper_or_lower == 'costarica':
            dealinfo.attrib['code']=makecode+'_SALE_AS-CR'
        
        if upper_or_lower == 'florida':
            dealinfo.attrib['code']=makecode+'_SALE_AS-FL'

        if upper_or_lower == 'others':
            dealinfo.attrib['code']=makecode+'_SALE_AS'




        traveldates = ET.SubElement(dealinfo, "TravelDates")
        if upper_or_lower == 'hawaii' or upper_or_lower == 'alaskahawaii':
            traveldates.attrib['startdate']= str(parseDates(getValueToTheRightOfString("Proposed AS.com")))+'T00:00:01'  
            traveldates.attrib['enddate']= str(parseDates(getValueToTheRightOfString("Calendar Dates - Hawaii")))+'T23:59:59'
        else:
            traveldates.attrib['startdate']= str(parseDates(getValueToTheRightOfString("Proposed AS.com")))+'T00:00:01'  
            traveldates.attrib['enddate']= str(parseDates(getValueToTheRightOfString("Calendar Dates - Others")))+'T23:59:59'


        dealtitle = ET.SubElement(dealinfo, "DealTitle")
        
        dealdescription = ET.SubElement(dealinfo, "DealDescrip").text = "<![CDATA[Purchase by "+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+".]]>"


        if upper_or_lower == 'alaskahawaii':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel between Alaska and Hawaii is valid '+getAvailability("To/From Hawaii")+' from '+str(dateInEnglish(getSpecificTravelStartDates("Alaska to/from Hawaii")))+' - '+str(dateInEnglish(getSpecificTravelEndDates("Alaska to/from Hawaii")))+'. Blackout dates are from '+getBlackoutDates("Alaska to/from Hawaii")+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'


        if upper_or_lower == 'hawaii':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel to and from Hawaii is valid '+getAvailability("To/From Hawaii")+' from '+str(dateInEnglish(getSpecificTravelStartDates("Hawaii")))+' - '+str(dateInEnglish(getSpecificTravelEndDates("Hawaii")))+'. Blackout dates are from '+getBlackoutDates("To Hawaii")+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'
       
        if upper_or_lower == 'mexico':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel to Mexico is valid '+getAvailability("To Mexico")+' from '+str(dateInEnglish(getSpecificTravelStartDates("Mexico")))+' - '+str(dateInEnglish(getSpecificTravelEndDates("Mexico")))+'. Blackout dates are from '+getBlackoutDates("Mexico")+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'

        if upper_or_lower == 'costarica':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel to Costa Rica is valid '+getAvailability("To Costa Rica")+' from '+str(dateInEnglish(getSpecificTravelStartDates("Costa Rica")))+' - '+str(dateInEnglish(getSpecificTravelEndDates("Costa Rica")))+'. Blackout dates are from '+getBlackoutDates("Costa Rica")+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'
        
        if upper_or_lower == 'florida':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel from Florida is valid '+getAvailability("From Florida")+' from '+str(dateInEnglish(getSpecificTravelStartDates("All Other Markets")))+' - '+str(dateInEnglish(getSpecificTravelEndDates("All Other Markets")))+'. Travel to Florida is valid '+getAvailability("To Florida")+' from '+str(dateInEnglish(getSpecificTravelStartDates("All Other Markets")))+' - '+str(dateInEnglish(getSpecificTravelEndDates("All Other Markets")))+'. Blackout dates are from '+getBlackoutDates("All Others")+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'


        if upper_or_lower == 'others':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel is valid '+getAvailability("All Others")+' from '+str(dateInEnglish(getSpecificTravelStartDates("All Other Markets")))+' - '+str(dateInEnglish(getSpecificTravelEndDates("All Other Markets")))+'. Blackout dates are from '+getBlackoutDates("All Others")+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'

        fares = ET.SubElement(dealset, "Fares")

        #This for loop will create each Row and Cell of XML for each item/dictionary in the list
        #pullFaresAndSaveInList(1, discoverSeparator()) RETURNS list of dictionaries
        for a in pullFaresAndSaveInList(which_rows):
            # print a['oCode'], a['oCity'], a['dCode'], a['dCity'],a['fare']
            row = ET.SubElement(fares, "Row") #showAsDefault="true"
            row.set('fareType', "Main") #Awards || Main || Saver
            cell = ET.SubElement(row, "Cell")
            ET.SubElement(cell, "Data").text = a['oCode']
            cell = ET.SubElement(row, "Cell")
            ET.SubElement(cell, "Data").text = a['oCity']
            cell = ET.SubElement(row, "Cell")
            ET.SubElement(cell, "Data").text = a['dCode']
            cell = ET.SubElement(row, "Cell")
            ET.SubElement(cell, "Data").text = a['dCity']
            cell = ET.SubElement(row, "Cell")
            ET.SubElement(cell, "Data").text = str(a['fare'])

        return dealset




    def exceptionDealSet(which_rows, advance_purchase, origin_code, destination_code, travel_valid, travel_start, travel_end, radiovalue):
        dealset = ET.SubElement(root, "DealSet")
        dealset.attrib['from']= str(parseDates(getValueToTheRightOfString("Sale Start Date:")))+'T00:00:01'
        dealset.attrib['to']= str(parseDates(getValueToTheRightOfString("Purchase By:")))+'T23:59:59'

        dealinfo = ET.SubElement(dealset, "DealInfo")
        dealinfo.attrib['url']=''
        dealinfo.attrib['dealType']='Standard' #MileagePlan || Standard || Saver
        makecode = str(parseDates(getValueToTheRightOfString("Sale Start Date:"))).replace('-', '')

        dealinfo.attrib['code']=makecode+'_SALE_AS-'+str(origin_code)+str(destination_code)
        
        traveldates = ET.SubElement(dealinfo, "TravelDates")


        if len(travel_start) > 0:
            traveldates.attrib['startdate']= str(travel_start)+'T00:00:01'  
        else:
            traveldates.attrib['startdate']= str(parseDates(getValueToTheRightOfString("Proposed AS.com")))+'T00:00:01'  


        if len(travel_end) > 0:
            traveldates.attrib['enddate']= str(travel_end)+'T23:59:59'
        else:
            traveldates.attrib['enddate']= str(parseDates(getValueToTheRightOfString("Calendar Dates - Others")))+'T23:59:59'

        dealtitle = ET.SubElement(dealinfo, "DealTitle")
        dealdescription = ET.SubElement(dealinfo, "DealDescrip").text = "<![CDATA[Purchase by "+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+".]]>"

        if radiovalue == 'travel_valid_value':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel from '+str(m(origin_code))+'('+str(origin_code)+') to '+str(m(destination_code))+'('+str(destination_code)+')'+' is valid '+str(travel_valid)+' from '+str(dateInEnglish(getSpecificTravelStartDates("All Other Markets")))+' - '+str(dateInEnglish(getSpecificTravelEndDates("All Other Markets")))+'. Blackout dates are from '+getBlackoutDates("All Others")+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'
        
        if radiovalue == 'service_start_value':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel from '+str(m(origin_code))+'('+str(origin_code)+') to '+str(m(destination_code))+'('+str(destination_code)+')'+' is valid from '+str(dateInEnglish(getSpecificTravelStartDates("All Other Markets")))+' - '+str(dateInEnglish(getSpecificTravelEndDates("All Other Markets")))+'.'+str(travel_valid)+'. Blackout dates are from '+getBlackoutDates("All Others")+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'

        fares = ET.SubElement(dealset, "Fares")

        #This for loop will create each Row and Cell of XML for each item/dictionary in the list
        #pullFaresAndSaveInList(1, discoverSeparator()) RETURNS list of dictionaries
        for a in pullFaresAndSaveInList(which_rows):
            # print a['oCode'], a['oCity'], a['dCode'], a['dCity'],a['fare']
            row = ET.SubElement(fares, "Row") #showAsDefault="true"
            row.set('fareType', "Main") #Awards || Main || Saver
            cell = ET.SubElement(row, "Cell")
            ET.SubElement(cell, "Data").text = a['oCode']
            cell = ET.SubElement(row, "Cell")
            ET.SubElement(cell, "Data").text = a['oCity']
            cell = ET.SubElement(row, "Cell")
            ET.SubElement(cell, "Data").text = a['dCode']
            cell = ET.SubElement(row, "Cell")
            ET.SubElement(cell, "Data").text = a['dCity']
            cell = ET.SubElement(row, "Cell")
            ET.SubElement(cell, "Data").text = str(a['fare'])

        return dealset




    #Books and Sheets
    book = xlrd.open_workbook(window.fileName)
    sheet_one = book.sheet_by_index(0)
    sheet = book.sheet_by_index(3)





    #NORMAL WEEKLY DEALS HANDLER
    if(marnel.get() == 'FlashSale'):

        #SERVICE EXCEPTION HANDLER
        if len(exception_codes1.get()) > 0 or len(exception_codes2.get()) > 0 or len(exception_codes3.get()) > 0 or len(exception_codes4.get()) > 0 or len(exception_codes5.get()) > 0 or len(exception_codes6.get()) > 0 or len(exception_codes7.get()) > 0 or len(exception_codes8.get()) > 0 or len(exception_codes9.get()) > 0 or len(exception_codes10.get()) > 0 or len(exception_codes11.get()) > 0 or len(exception_codes12.get()) > 0 or len(exception_codes13.get()) > 0 or len(exception_codes14.get()) > 0 or len(exception_codes15.get()) > 0 or len(exception_codes16.get()) > 0 or len(exception_codes17.get()) > 0 or len(exception_codes18.get()) > 0 or len(exception_codes19.get()) > 0 or len(exception_codes20.get()) > 0:
            total_exceptions_list = []

            if len(exception_codes1.get()) > 0:
                if len(exception_start1.get()) > 0:
                    travel_start = exception_start1.get()
                else:
                    travel_start = ''

                if len(exception_end1.get()) > 0:
                    travel_end = exception_end1.get()
                else:
                    travel_end = ''

                ex_code_1 =  str(exception_codes1.get()).strip()
                ex_code_1_origin = ex_code_1[0]+ex_code_1[1]+ex_code_1[2]
                ex_code_1_destination = ex_code_1[3]+ex_code_1[4]+ex_code_1[5]
                ex_code_1_days = str(exception_days1.get())
       
                total_exceptions_1 = serviceExceptionFares("AS", ex_code_1_origin, ex_code_1_destination)
                if total_exceptions_1:
                    total_exceptions_list.append(total_exceptions_1[0])
                else:
                    print(exception_codes1.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION ONE: "+str(len(serviceExceptionFares("AS", ex_code_1_origin, ex_code_1_destination))))
                print(serviceExceptionFares("AS", ex_code_1_origin, ex_code_1_destination))
                serviceExceptionFares("AS", ex_code_1_origin, ex_code_1_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_1_origin, ex_code_1_destination),pass_AdvancePurchase,ex_code_1_origin,ex_code_1_destination,ex_code_1_days,travel_start,travel_end, except_1_variable.get())


            if len(exception_codes2.get()) > 0:

                if len(exception_start2.get()) > 0:
                    travel_start = exception_start2.get()
                else:
                    travel_start = ''

                if len(exception_end2.get()) > 0:
                    travel_end = exception_end2.get()
                else:
                    travel_end = ''

                ex_code_2 =  str(exception_codes2.get()).strip()
                ex_code_2_origin = ex_code_2[0]+ex_code_2[1]+ex_code_2[2]
                ex_code_2_destination = ex_code_2[3]+ex_code_2[4]+ex_code_2[5]
                ex_code_2_days = str(exception_days2.get())
       
                total_exceptions_2 = serviceExceptionFares("AS", ex_code_2_origin, ex_code_2_destination)
                if total_exceptions_2:
                    total_exceptions_list.append(total_exceptions_2[0])
                else:
                    print(exception_codes2.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION TWO: "+str(len(serviceExceptionFares("AS", ex_code_2_origin, ex_code_2_destination))))
                print(serviceExceptionFares("AS", ex_code_2_origin, ex_code_2_destination))
                serviceExceptionFares("AS", ex_code_2_origin, ex_code_2_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_2_origin, ex_code_2_destination),pass_AdvancePurchase,ex_code_2_origin,ex_code_2_destination,ex_code_2_days,travel_start,travel_end, except_2_variable.get())


            if len(exception_codes3.get()) > 0:
                if len(exception_start3.get()) > 0:
                    travel_start = exception_start3.get()
                else:
                    travel_start = ''

                if len(exception_end3.get()) > 0:
                    travel_end = exception_end3.get()
                else:
                    travel_end = ''

                ex_code_3 =  str(exception_codes3.get()).strip()
                ex_code_3_origin = ex_code_3[0]+ex_code_3[1]+ex_code_3[2]
                ex_code_3_destination = ex_code_3[3]+ex_code_3[4]+ex_code_3[5]
                ex_code_3_days = str(exception_days3.get())
       
                total_exceptions_3 = serviceExceptionFares("AS", ex_code_3_origin, ex_code_3_destination)
                if total_exceptions_3:
                    total_exceptions_list.append(total_exceptions_3[0])
                else:
                    print(exception_codes3.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION THREE: "+str(len(serviceExceptionFares("AS", ex_code_3_origin, ex_code_3_destination))))
                print(serviceExceptionFares("AS", ex_code_3_origin, ex_code_3_destination))
                serviceExceptionFares("AS", ex_code_3_origin, ex_code_3_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_3_origin, ex_code_3_destination),pass_AdvancePurchase,ex_code_3_origin,ex_code_3_destination,ex_code_3_days,travel_start,travel_end, except_3_variable.get())


            if len(exception_codes4.get()) > 0:
                if len(exception_start4.get()) > 0:
                    travel_start = exception_start4.get()
                else:
                    travel_start = ''

                if len(exception_end4.get()) > 0:
                    travel_end = exception_end4.get()
                else:
                    travel_end = ''

                ex_code_4 =  str(exception_codes4.get()).strip()
                ex_code_4_origin = ex_code_4[0]+ex_code_4[1]+ex_code_4[2]
                ex_code_4_destination = ex_code_4[3]+ex_code_4[4]+ex_code_4[5]
                ex_code_4_days = str(exception_days4.get())
       
                total_exceptions_4 = serviceExceptionFares("AS", ex_code_4_origin, ex_code_4_destination)
                if total_exceptions_4:
                    total_exceptions_list.append(total_exceptions_4[0])
                else:
                    print(exception_codes4.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION FOUR: "+str(len(serviceExceptionFares("AS", ex_code_4_origin, ex_code_4_destination))))
                print(serviceExceptionFares("AS", ex_code_4_origin, ex_code_4_destination))
                serviceExceptionFares("AS", ex_code_4_origin, ex_code_4_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_4_origin, ex_code_4_destination),pass_AdvancePurchase,ex_code_4_origin,ex_code_4_destination,ex_code_4_days,travel_start,travel_end, except_4_variable.get())


            if len(exception_codes5.get()) > 0:
                if len(exception_start5.get()) > 0:
                    travel_start = exception_start5.get()
                else:
                    travel_start = ''

                if len(exception_end5.get()) > 0:
                    travel_end = exception_end5.get()
                else:
                    travel_end = ''

                ex_code_5 =  str(exception_codes5.get()).strip()
                ex_code_5_origin = ex_code_5[0]+ex_code_5[1]+ex_code_5[2]
                ex_code_5_destination = ex_code_5[3]+ex_code_5[4]+ex_code_5[5]
                ex_code_5_days = str(exception_days5.get())
       
                total_exceptions_5 = serviceExceptionFares("AS", ex_code_5_origin, ex_code_5_destination)
                if total_exceptions_5:
                    total_exceptions_list.append(total_exceptions_5[0])
                else:
                    print(exception_codes5.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION FIVE: "+str(len(serviceExceptionFares("AS", ex_code_5_origin, ex_code_5_destination))))
                print(serviceExceptionFares("AS", ex_code_5_origin, ex_code_5_destination))
                serviceExceptionFares("AS", ex_code_5_origin, ex_code_5_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_5_origin, ex_code_5_destination),pass_AdvancePurchase,ex_code_5_origin,ex_code_5_destination,ex_code_5_days,travel_start,travel_end, except_5_variable.get())


            if len(exception_codes6.get()) > 0:
                if len(exception_start6.get()) > 0:
                    travel_start = exception_start6.get()
                else:
                    travel_start = ''

                if len(exception_end6.get()) > 0:
                    travel_end = exception_end6.get()
                else:
                    travel_end = ''

                ex_code_6 =  str(exception_codes6.get()).strip()
                ex_code_6_origin = ex_code_6[0]+ex_code_6[1]+ex_code_6[2]
                ex_code_6_destination = ex_code_6[3]+ex_code_6[4]+ex_code_6[5]
                ex_code_6_days = str(exception_days6.get())
       
                total_exceptions_6 = serviceExceptionFares("AS", ex_code_6_origin, ex_code_6_destination)
                if total_exceptions_6:
                    total_exceptions_list.append(total_exceptions_6[0])
                else:
                    print(exception_codes6.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION SIX: "+str(len(serviceExceptionFares("AS", ex_code_6_origin, ex_code_6_destination))))
                print(serviceExceptionFares("AS", ex_code_6_origin, ex_code_6_destination))
                serviceExceptionFares("AS", ex_code_6_origin, ex_code_6_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_6_origin, ex_code_6_destination),pass_AdvancePurchase,ex_code_6_origin,ex_code_6_destination,ex_code_6_days,travel_start,travel_end, except_6_variable.get())


            if len(exception_codes7.get()) > 0:
                if len(exception_start7.get()) > 0:
                    travel_start = exception_start7.get()
                else:
                    travel_start = ''

                if len(exception_end7.get()) > 0:
                    travel_end = exception_end7.get()
                else:
                    travel_end = ''

                ex_code_7 =  str(exception_codes7.get()).strip()
                ex_code_7_origin = ex_code_7[0]+ex_code_7[1]+ex_code_7[2]
                ex_code_7_destination = ex_code_7[3]+ex_code_7[4]+ex_code_7[5]
                ex_code_7_days = str(exception_days7.get())
       
                total_exceptions_7 = serviceExceptionFares("AS", ex_code_7_origin, ex_code_7_destination)
                if total_exceptions_7:
                    total_exceptions_list.append(total_exceptions_7[0])
                else:
                    print(exception_codes7.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION SEVEN: "+str(len(serviceExceptionFares("AS", ex_code_7_origin, ex_code_7_destination))))
                print(serviceExceptionFares("AS", ex_code_7_origin, ex_code_7_destination))
                serviceExceptionFares("AS", ex_code_7_origin, ex_code_7_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_7_origin, ex_code_7_destination),pass_AdvancePurchase,ex_code_7_origin,ex_code_7_destination,ex_code_7_days,travel_start,travel_end, except_7_variable.get())


            if len(exception_codes8.get()) > 0:
                if len(exception_start8.get()) > 0:
                    travel_start = exception_start8.get()
                else:
                    travel_start = ''

                if len(exception_end8.get()) > 0:
                    travel_end = exception_end8.get()
                else:
                    travel_end = ''

                ex_code_8 =  str(exception_codes8.get()).strip()
                ex_code_8_origin = ex_code_8[0]+ex_code_8[1]+ex_code_8[2]
                ex_code_8_destination = ex_code_8[3]+ex_code_8[4]+ex_code_8[5]
                ex_code_8_days = str(exception_days8.get())
       
                total_exceptions_8 = serviceExceptionFares("AS", ex_code_8_origin, ex_code_8_destination)
                if total_exceptions_8:
                    total_exceptions_list.append(total_exceptions_8[0])
                else:
                    print(exception_codes8.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION EIGHT: "+str(len(serviceExceptionFares("AS", ex_code_8_origin, ex_code_8_destination))))
                print(serviceExceptionFares("AS", ex_code_8_origin, ex_code_8_destination))
                serviceExceptionFares("AS", ex_code_8_origin, ex_code_8_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_8_origin, ex_code_8_destination),pass_AdvancePurchase,ex_code_8_origin,ex_code_8_destination,ex_code_8_days,travel_start,travel_end, except_8_variable.get())


            if len(exception_codes9.get()) > 0:
                if len(exception_start9.get()) > 0:
                    travel_start = exception_start9.get()
                else:
                    travel_start = ''

                if len(exception_end9.get()) > 0:
                    travel_end = exception_end9.get()
                else:
                    travel_end = ''

                ex_code_9 =  str(exception_codes9.get()).strip()
                ex_code_9_origin = ex_code_9[0]+ex_code_9[1]+ex_code_9[2]
                ex_code_9_destination = ex_code_9[3]+ex_code_9[4]+ex_code_9[5]
                ex_code_9_days = str(exception_days9.get())
       
                total_exceptions_9 = serviceExceptionFares("AS", ex_code_9_origin, ex_code_9_destination)
                if total_exceptions_9:
                    total_exceptions_list.append(total_exceptions_9[0])
                else:
                    print(exception_codes9.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION NINE: "+str(len(serviceExceptionFares("AS", ex_code_9_origin, ex_code_9_destination))))
                print(serviceExceptionFares("AS", ex_code_9_origin, ex_code_9_destination))
                serviceExceptionFares("AS", ex_code_9_origin, ex_code_9_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_9_origin, ex_code_9_destination),pass_AdvancePurchase,ex_code_9_origin,ex_code_9_destination,ex_code_9_days,travel_start,travel_end, except_9_variable.get())


            if len(exception_codes10.get()) > 0:
                if len(exception_start10.get()) > 0:
                    travel_start = exception_start10.get()
                else:
                    travel_start = ''

                if len(exception_end10.get()) > 0:
                    travel_end = exception_end10.get()
                else:
                    travel_end = ''

                ex_code_10 =  str(exception_codes10.get()).strip()
                ex_code_10_origin = ex_code_10[0]+ex_code_10[1]+ex_code_10[2]
                ex_code_10_destination = ex_code_10[3]+ex_code_10[4]+ex_code_10[5]
                ex_code_10_days = str(exception_days10.get())
       
                total_exceptions_10 = serviceExceptionFares("AS", ex_code_10_origin, ex_code_10_destination)
                if total_exceptions_10:
                    total_exceptions_list.append(total_exceptions_10[0])
                else:
                    print(exception_codes10.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION TEN: "+str(len(serviceExceptionFares("AS", ex_code_10_origin, ex_code_10_destination))))
                print(serviceExceptionFares("AS", ex_code_10_origin, ex_code_10_destination))
                serviceExceptionFares("AS", ex_code_10_origin, ex_code_10_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_10_origin, ex_code_10_destination),pass_AdvancePurchase,ex_code_10_origin,ex_code_10_destination,ex_code_10_days,travel_start,travel_end, except_10_variable.get())


            if len(exception_codes11.get()) > 0:
                if len(exception_start11.get()) > 0:
                    travel_start = exception_start11.get()
                else:
                    travel_start = ''

                if len(exception_end11.get()) > 0:
                    travel_end = exception_end11.get()
                else:
                    travel_end = ''

                ex_code_11 =  str(exception_codes11.get()).strip()
                ex_code_11_origin = ex_code_11[0]+ex_code_11[1]+ex_code_11[2]
                ex_code_11_destination = ex_code_11[3]+ex_code_11[4]+ex_code_11[5]
                ex_code_11_days = str(exception_days11.get())
       
                total_exceptions_11 = serviceExceptionFares("AS", ex_code_11_origin, ex_code_11_destination)
                if total_exceptions_11:
                    total_exceptions_list.append(total_exceptions_11[0])
                else:
                    print(exception_codes11.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION Elevel: "+str(len(serviceExceptionFares("AS", ex_code_11_origin, ex_code_11_destination))))
                print(serviceExceptionFares("AS", ex_code_11_origin, ex_code_11_destination))
                serviceExceptionFares("AS", ex_code_11_origin, ex_code_11_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_11_origin, ex_code_11_destination),pass_AdvancePurchase,ex_code_11_origin,ex_code_11_destination,ex_code_11_days,travel_start,travel_end, except_11_variable.get())


            if len(exception_codes12.get()) > 0:
                if len(exception_start12.get()) > 0:
                    travel_start = exception_start12.get()
                else:
                    travel_start = ''

                if len(exception_end12.get()) > 0:
                    travel_end = exception_end12.get()
                else:
                    travel_end = ''

                ex_code_12 =  str(exception_codes12.get()).strip()
                ex_code_12_origin = ex_code_12[0]+ex_code_12[1]+ex_code_12[2]
                ex_code_12_destination = ex_code_12[3]+ex_code_12[4]+ex_code_12[5]
                ex_code_12_days = str(exception_days12.get())
       
                total_exceptions_12 = serviceExceptionFares("AS", ex_code_12_origin, ex_code_12_destination)
                if total_exceptions_12:
                    total_exceptions_list.append(total_exceptions_12[0])
                else:
                    print(exception_codes12.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION Twelve: "+str(len(serviceExceptionFares("AS", ex_code_12_origin, ex_code_12_destination))))
                print(serviceExceptionFares("AS", ex_code_12_origin, ex_code_12_destination))
                serviceExceptionFares("AS", ex_code_12_origin, ex_code_12_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_12_origin, ex_code_12_destination),pass_AdvancePurchase,ex_code_12_origin,ex_code_12_destination,ex_code_12_days,travel_start,travel_end, except_12_variable.get())


            if len(exception_codes13.get()) > 0:
                if len(exception_start13.get()) > 0:
                    travel_start = exception_start13.get()
                else:
                    travel_start = ''

                if len(exception_end13.get()) > 0:
                    travel_end = exception_end13.get()
                else:
                    travel_end = ''

                ex_code_13 =  str(exception_codes13.get()).strip()
                ex_code_13_origin = ex_code_13[0]+ex_code_13[1]+ex_code_13[2]
                ex_code_13_destination = ex_code_13[3]+ex_code_13[4]+ex_code_13[5]
                ex_code_13_days = str(exception_days13.get())
       
                total_exceptions_13 = serviceExceptionFares("AS", ex_code_13_origin, ex_code_13_destination)
                if total_exceptions_13:
                    total_exceptions_list.append(total_exceptions_13[0])
                else:
                    print(exception_codes13.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION Thirteen: "+str(len(serviceExceptionFares("AS", ex_code_13_origin, ex_code_13_destination))))
                print(serviceExceptionFares("AS", ex_code_13_origin, ex_code_13_destination))
                serviceExceptionFares("AS", ex_code_13_origin, ex_code_13_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_13_origin, ex_code_13_destination),pass_AdvancePurchase,ex_code_13_origin,ex_code_13_destination,ex_code_13_days,travel_start,travel_end, except_13_variable.get())


            if len(exception_codes14.get()) > 0:
                if len(exception_start14.get()) > 0:
                    travel_start = exception_start14.get()
                else:
                    travel_start = ''

                if len(exception_end14.get()) > 0:
                    travel_end = exception_end14.get()
                else:
                    travel_end = ''

                ex_code_14 =  str(exception_codes14.get()).strip()
                ex_code_14_origin = ex_code_14[0]+ex_code_14[1]+ex_code_14[2]
                ex_code_14_destination = ex_code_14[3]+ex_code_14[4]+ex_code_14[5]
                ex_code_14_days = str(exception_days14.get())
       
                total_exceptions_14 = serviceExceptionFares("AS", ex_code_14_origin, ex_code_14_destination)
                if total_exceptions_14:
                    total_exceptions_list.append(total_exceptions_14[0])
                else:
                    print(exception_codes14.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION Fourteen: "+str(len(serviceExceptionFares("AS", ex_code_14_origin, ex_code_14_destination))))
                print(serviceExceptionFares("AS", ex_code_14_origin, ex_code_14_destination))
                serviceExceptionFares("AS", ex_code_14_origin, ex_code_14_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_14_origin, ex_code_14_destination),pass_AdvancePurchase,ex_code_14_origin,ex_code_14_destination,ex_code_14_days,travel_start,travel_end, except_14_variable.get())


            if len(exception_codes15.get()) > 0:
                if len(exception_start15.get()) > 0:
                    travel_start = exception_start15.get()
                else:
                    travel_start = ''

                if len(exception_end15.get()) > 0:
                    travel_end = exception_end15.get()
                else:
                    travel_end = ''

                ex_code_15 =  str(exception_codes15.get()).strip()
                ex_code_15_origin = ex_code_15[0]+ex_code_15[1]+ex_code_15[2]
                ex_code_15_destination = ex_code_15[3]+ex_code_15[4]+ex_code_15[5]
                ex_code_15_days = str(exception_days15.get())
       
                total_exceptions_15 = serviceExceptionFares("AS", ex_code_15_origin, ex_code_15_destination)
                if total_exceptions_15:
                    total_exceptions_list.append(total_exceptions_15[0])
                else:
                    print(exception_codes15.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION Fifteen: "+str(len(serviceExceptionFares("AS", ex_code_15_origin, ex_code_15_destination))))
                print(serviceExceptionFares("AS", ex_code_15_origin, ex_code_15_destination))
                serviceExceptionFares("AS", ex_code_15_origin, ex_code_15_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_15_origin, ex_code_15_destination),pass_AdvancePurchase,ex_code_15_origin,ex_code_15_destination,ex_code_15_days,travel_start,travel_end, except_15_variable.get())


            if len(exception_codes16.get()) > 0:
                if len(exception_start16.get()) > 0:
                    travel_start = exception_start16.get()
                else:
                    travel_start = ''

                if len(exception_end16.get()) > 0:
                    travel_end = exception_end16.get()
                else:
                    travel_end = ''

                ex_code_16 =  str(exception_codes16.get()).strip()
                ex_code_16_origin = ex_code_16[0]+ex_code_16[1]+ex_code_16[2]
                ex_code_16_destination = ex_code_16[3]+ex_code_16[4]+ex_code_16[5]
                ex_code_16_days = str(exception_days16.get())
       
                total_exceptions_16 = serviceExceptionFares("AS", ex_code_16_origin, ex_code_16_destination)
                if total_exceptions_16:
                    total_exceptions_list.append(total_exceptions_16[0])
                else:
                    print(exception_codes16.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION Sixteen: "+str(len(serviceExceptionFares("AS", ex_code_16_origin, ex_code_16_destination))))
                print(serviceExceptionFares("AS", ex_code_16_origin, ex_code_16_destination))
                serviceExceptionFares("AS", ex_code_16_origin, ex_code_16_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_16_origin, ex_code_16_destination),pass_AdvancePurchase,ex_code_16_origin,ex_code_16_destination,ex_code_16_days,travel_start,travel_end, except_16_variable.get())


            if len(exception_codes17.get()) > 0:
                if len(exception_start17.get()) > 0:
                    travel_start = exception_start17.get()
                else:
                    travel_start = ''

                if len(exception_end17.get()) > 0:
                    travel_end = exception_end17.get()
                else:
                    travel_end = ''

                ex_code_17 =  str(exception_codes17.get()).strip()
                ex_code_17_origin = ex_code_17[0]+ex_code_17[1]+ex_code_17[2]
                ex_code_17_destination = ex_code_17[3]+ex_code_17[4]+ex_code_17[5]
                ex_code_17_days = str(exception_days17.get())
       
                total_exceptions_17 = serviceExceptionFares("AS", ex_code_17_origin, ex_code_17_destination)
                if total_exceptions_17:
                    total_exceptions_list.append(total_exceptions_17[0])
                else:
                    print(exception_codes17.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION Seventeen: "+str(len(serviceExceptionFares("AS", ex_code_17_origin, ex_code_17_destination))))
                print(serviceExceptionFares("AS", ex_code_17_origin, ex_code_17_destination))
                serviceExceptionFares("AS", ex_code_17_origin, ex_code_17_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_17_origin, ex_code_17_destination),pass_AdvancePurchase,ex_code_17_origin,ex_code_17_destination,ex_code_17_days,travel_start,travel_end, except_17_variable.get())


            if len(exception_codes18.get()) > 0:
                if len(exception_start18.get()) > 0:
                    travel_start = exception_start18.get()
                else:
                    travel_start = ''

                if len(exception_end18.get()) > 0:
                    travel_end = exception_end18.get()
                else:
                    travel_end = ''

                ex_code_18 =  str(exception_codes18.get()).strip()
                ex_code_18_origin = ex_code_18[0]+ex_code_18[1]+ex_code_18[2]
                ex_code_18_destination = ex_code_18[3]+ex_code_18[4]+ex_code_18[5]
                ex_code_18_days = str(exception_days18.get())
       
                total_exceptions_18 = serviceExceptionFares("AS", ex_code_18_origin, ex_code_18_destination)
                if total_exceptions_18:
                    total_exceptions_list.append(total_exceptions_18[0])
                else:
                    print(exception_codes18.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION Eighteen: "+str(len(serviceExceptionFares("AS", ex_code_18_origin, ex_code_18_destination))))
                print(serviceExceptionFares("AS", ex_code_18_origin, ex_code_18_destination))
                serviceExceptionFares("AS", ex_code_18_origin, ex_code_18_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_18_origin, ex_code_18_destination),pass_AdvancePurchase,ex_code_18_origin,ex_code_18_destination,ex_code_18_days,travel_start,travel_end, except_18_variable.get())


            if len(exception_codes19.get()) > 0:
                if len(exception_start19.get()) > 0:
                    travel_start = exception_start19.get()
                else:
                    travel_start = ''

                if len(exception_end19.get()) > 0:
                    travel_end = exception_end19.get()
                else:
                    travel_end = ''

                ex_code_19 =  str(exception_codes19.get()).strip()
                ex_code_19_origin = ex_code_19[0]+ex_code_19[1]+ex_code_19[2]
                ex_code_19_destination = ex_code_19[3]+ex_code_19[4]+ex_code_19[5]
                ex_code_19_days = str(exception_days19.get())
       
                total_exceptions_19 = serviceExceptionFares("AS", ex_code_19_origin, ex_code_19_destination)
                if total_exceptions_19:
                    total_exceptions_list.append(total_exceptions_19[0])
                else:
                    print(exception_codes19.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION Nineteen: "+str(len(serviceExceptionFares("AS", ex_code_19_origin, ex_code_19_destination))))
                print(serviceExceptionFares("AS", ex_code_19_origin, ex_code_19_destination))
                serviceExceptionFares("AS", ex_code_19_origin, ex_code_19_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_19_origin, ex_code_19_destination),pass_AdvancePurchase,ex_code_19_origin,ex_code_19_destination,ex_code_19_days,travel_start,travel_end, except_19_variable.get())


            if len(exception_codes20.get()) > 0:

                if len(exception_start20.get()) > 0:
                    travel_start = exception_start20.get()
                else:
                    travel_start = ''

                if len(exception_end20.get()) > 0:
                    travel_end = exception_end20.get()
                else:
                    travel_end = ''

                ex_code_20 =  str(exception_codes20.get()).strip()
                ex_code_20_origin = ex_code_20[0]+ex_code_20[1]+ex_code_20[2]
                ex_code_20_destination = ex_code_20[3]+ex_code_20[4]+ex_code_20[5]
                ex_code_20_days = str(exception_days20.get())
       
                total_exceptions_20 = serviceExceptionFares("AS", ex_code_20_origin, ex_code_20_destination)
                if total_exceptions_20:
                    total_exceptions_list.append(total_exceptions_20[0])
                else:
                    print(exception_codes20.get()+' fares do not exist in this spreadsheet.')
                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print("EXCEPTION Twenty: "+str(len(serviceExceptionFares("AS", ex_code_20_origin, ex_code_20_destination))))
                print(serviceExceptionFares("AS", ex_code_20_origin, ex_code_20_destination))
                serviceExceptionFares("AS", ex_code_20_origin, ex_code_20_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_20_origin, ex_code_20_destination),pass_AdvancePurchase,ex_code_20_origin,ex_code_20_destination,ex_code_20_days,travel_start,travel_end, except_20_variable.get())


        else:
            total_exceptions_list = []

        print(total_exceptions_list)

        if len(alaskaToFromHawaiiFares("AS", total_exceptions_list)) > 0:
            pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
            print("ALASKA TO/FROM HAWAII: "+str(len(alaskaToFromHawaiiFares("AS", total_exceptions_list))))
            print(alaskaToFromHawaiiFares("AS", total_exceptions_list))
            flashDealSet(alaskaToFromHawaiiFares("AS", total_exceptions_list),pass_AdvancePurchase,"alaskahawaii")

        if len(hawaiiFares("AS", total_exceptions_list)) > 0:
            pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
            print("HAWAII: "+str(len(hawaiiFares("AS", total_exceptions_list))))
            print(hawaiiFares("AS", total_exceptions_list))
            flashDealSet(hawaiiFares("AS", total_exceptions_list),pass_AdvancePurchase,"hawaii")

        if len(mexicoFares("AS", total_exceptions_list)) > 0:
            pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
            print("MEXICO: "+str(len(mexicoFares("AS", total_exceptions_list))))
            print(mexicoFares("AS", total_exceptions_list))
            flashDealSet(mexicoFares("AS", total_exceptions_list),pass_AdvancePurchase,"mexico")

        if len(costaricaFares("AS", total_exceptions_list)) > 0:
            pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
            print("COSTA RICA: "+str(len(costaricaFares("AS", total_exceptions_list))))
            print(costaricaFares("AS", total_exceptions_list))
            flashDealSet(costaricaFares("AS", total_exceptions_list),pass_AdvancePurchase,"costarica")

        if len(floridaFares("AS", total_exceptions_list)) > 0:
            pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
            print("FLORIDA: "+str(len(floridaFares("AS", total_exceptions_list))))
            print(floridaFares("AS", total_exceptions_list))
            flashDealSet(floridaFares("AS", total_exceptions_list),pass_AdvancePurchase,"florida")

        if len(allOtherRows("AS", total_exceptions_list)) > 0:
            pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
            print("ALL OTHERS: "+str(len(allOtherRows("AS", total_exceptions_list))))
            print(allOtherRows("AS", total_exceptions_list))
            flashDealSet(allOtherRows("AS", total_exceptions_list),pass_AdvancePurchase,"others")




    tree.write("flash-steve.xml")


fileBtn.configure(command=getfile)
runBtn.configure(command=automate)
resBtn.configure(command=reset)

window.mainloop()
