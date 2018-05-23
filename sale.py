from Tkinter import *
import tkFileDialog
import os
import xlrd
import json, ast, os, string, random, urllib
import xml.etree.cElementTree as ET
import datetime
import dateparser
import tkMessageBox
 


window = Tk()
img = PhotoImage (file = 'AS.gif')
imgLb1 = Label (window, image = img)
browseLabel = Label (window, width=30)
fileBtn = Button (window, padx=10, pady=20)
marnel = StringVar()

#entry_travel_start = Entry(window)
#entry_start_label = Label(window, width=1, text="Travel Start")

radio_1 = Radiobutton(window, text="Weekly ", variable=marnel, value="AS")
radio_2 = Radiobutton(window, text="Club 49", variable=marnel, value="C49")
radio_3 = Radiobutton(window, text="Last Minute", variable=marnel, value="LastMinute")
radio_4 = Radiobutton(window, text="Award Sale", variable=marnel, value="AwardSale")

weekly_label = Label(window, text="Normal Deals. Can be used for larger sales.", width=40)
club_49_label = Label(window, text="Specific to Club 49 fares.", width=40)
last_minute_label = Label(window, text="For deals that have travel dates coming up.", width=40)


exception_codes_label1 = Label(window, text="EXCEPTIONS:", width=10)
exception_codes_label2 = Label(window, text="Codes", width=8)
exception_codes_label3 = Label(window, text="Days of Travel", width=25)
exception_codes_label4 = Label(window, text="startdate", width=8)
exception_codes_label5 = Label(window, text="enddate", width=8)

example_label = Label(window, text="Format Examples:", width=10)
example_codes_label = Label(window, text="MCOSFO", width=8)
example_days_label = Label(window, text="Thursday through Monday", width=25)
example_start_label = Label(window, text="2018-09-22", width=8)
example_end_label = Label(window, text="2018-09-30", width=8)



example_label.configure(fg="gray")
example_codes_label.configure(fg="gray")
example_days_label.configure(fg="gray")
example_start_label.configure(fg="gray")
example_end_label.configure(fg="gray")


weekly_label.configure(fg="gray")
club_49_label.configure(fg="gray")
last_minute_label.configure(fg="gray")


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


runBtn = Button (window, padx=10, pady=20)
resBtn = Button (window, padx=20, pady=20)


resBtn.configure(fg="red",bg="red")
runBtn.configure(fg="green",bg="green")

radio_1.select()

imgLb1.grid(row=1, column=1, rowspan=1, columnspan = 6)
browseLabel.grid(row=3, column=1)

fileBtn.grid(row=3, column=2, columnspan = 1)

radio_1.grid(row=4, column=1, columnspan = 1)
radio_2.grid(row=5, column=1, columnspan = 1)
radio_3.grid(row=6, column=1, columnspan = 1)
radio_4.grid(row=7, column=1, columnspan = 1)


weekly_label.grid(row=5, column=2, columnspan = 2)
club_49_label.grid(row=6, column=2, columnspan = 2)
last_minute_label.grid(row=7, column=2, columnspan = 2)



exception_codes_label1.grid(row=10, column=1)
exception_codes_label2.grid(row=8, column=2)
exception_codes_label3.grid(row=8, column=3)
exception_codes_label4.grid(row=8, column=4)
exception_codes_label5.grid(row=8, column=5)


example_label.grid(row=9, column=1)
example_codes_label.grid(row=9, column=2)
example_days_label.grid(row=9, column=3)
example_start_label.grid(row=9, column=4)
example_end_label.grid(row=9, column=5)


exception_codes1.grid(row=10, column=2)
exception_days1.grid(row=10, column=3)
exception_start1.grid(row=10, column=4)
exception_end1.grid(row=10, column=5)

exception_codes2.grid(row=11, column=2)
exception_days2.grid(row=11, column=3)
exception_start2.grid(row=11, column=4)
exception_end2.grid(row=11, column=5)

exception_codes3.grid(row=12, column=2)
exception_days3.grid(row=12, column=3)
exception_start3.grid(row=12, column=4)
exception_end3.grid(row=12, column=5)

exception_codes4.grid(row=13, column=2)
exception_days4.grid(row=13, column=3)
exception_start4.grid(row=13, column=4)
exception_end4.grid(row=13, column=5)


runBtn.grid(row=15, column=2, columnspan = 1)
resBtn.grid(row=1, column=4, columnspan = 1)


window.title('Flight Deals Automation')
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
    window.fileName =  tkFileDialog.askopenfilename(title = "Select file",filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))
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
    exception_days1.delete(0, 'end')
    exception_days2.delete(0, 'end')
    exception_days3.delete(0, 'end')
    exception_days4.delete(0, 'end')
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
    print window.fileName
    print marnel.get()
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
            'LWS' : 'Lewiston',
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
        print my_year
        return int(my_year)



    def getMonth(this_date):
        value_int = xlrd.xldate_as_tuple(int(this_date), 0)
        parsed_date = datetime.date(value_int[0], value_int[1], value_int[2])
        my_month = str(parsed_date)
        my_month = my_month.split("-",2)[1]
        print my_month
        return int(my_month)


    def getDay(this_date):
        value_int = xlrd.xldate_as_tuple(int(this_date), 0)
        parsed_date = datetime.date(value_int[0], value_int[1], value_int[2])
        my_day = str(parsed_date)
        my_day = my_day.split("-",3)[2]
        print my_day
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
                #print row_index+1
                return row_index+1
            

    def getValueToTheRightOfString(string_to_search_for):
        for row_index in xrange(1, getStringCoordinates(string_to_search_for)):
            if sheet_one.cell(row_index, 0).value.strip() == string_to_search_for:
                if sheet_one.cell(row_index, 1).value:
                    #print string_to_search_for+": "+str(parseDates(sheet_one.cell(row_index, 1).value))
                    #print string_to_search_for+": "+str(dateInEnglish(sheet_one.cell(row_index, 1).value))
                    #return parseDates(sheet_one.cell(row_index, 1).value)
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
                    #print string_to_search_for+": "+sheet_one.cell(row_index, 1).value
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


   
    def hawaiiFares(airline_type, total_exceptions_list):
        hawaii_codes = ["OGG","LIH","KOA","HNL"]
        hawaii_list = []
        for col in range(5,7):
            for row in range(2, sheet.nrows):
                if sheet.cell_value(row, 5) == airline_type:
                    if sheet.cell_value(row, 7) in hawaii_codes or sheet.cell_value(row, 9) in hawaii_codes:
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
        
        
    


    #Get all Row of Fares depending on what airline and if Hawaii or International Fares
    def getClub49Fares(airline_type, upper_or_lower):
        alaska_codes = ["ADK","ANC","BRW","BET","CDV","DLG","DUT","FAI","GST","JNU","KTN","AKN","ADQ","OTZ","OME","PSG","SCC","SIT","WRG","YAK"]
        all_other_fares = []
        upper_list = []
        lower_list = []
        for col in range(5,7):
            for row in range(1, sheet.nrows):
                if sheet.cell_value(row, 5) == airline_type:
                    if upper_or_lower == 'upper':
                        if sheet.cell_value(row, 9) in alaska_codes:
                            upper_list.append(row)
                        all_other_fares = upper_list
                    else:
                        if sheet.cell_value(row, 9) not in alaska_codes:
                            lower_list.append(row)         
                        all_other_fares = lower_list
            return all_other_fares



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



    def lastMinuteRows(airline_type, total_exceptions_list):
        last_minute_list = []
        for col in range(5,7):
            for row in range(1, sheet.nrows):
                if sheet.cell_value(row, 5) == airline_type:
                    last_minute_list.append(row)
                else:
                    last_minute_fares = last_minute_list
                    continue
            last_minute_fares = removeDuplicates(last_minute_fares, total_exceptions_list)
            return last_minute_fares


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
        #print ast.literal_eval(json.dumps(my_fares))
        #returns list
        return my_fares

    

    #AWARD SALE SPECIFIC
    def pullAwardSaleFaresAndSaveInList(list_being_passed):
        #This sets the name of all keys for the list of dictionary  
        keys = ["oCode","oCity","dCode","dCity","fare","fees"]
        my_dictionary_list = []
        # this selects how many rows to read
        for row in range(1, sheet.nrows):
            if row in list_being_passed:
                my_dictionary_list.append({keys[0]: sheet.cell(row, 7).value,keys[1]: sheet.cell(row, 8).value,keys[2]: sheet.cell(row, 9).value,keys[3]: sheet.cell(row, 10).value,keys[4]: int(sheet.cell(row, 11).value),keys[5]: int(sheet.cell(row, 12).value)})
        # saves the list into a variable
        #my_fares = sorted(my_dictionary_list, key=itemgetter('fare'), key=itemgetter('oCity'), key=itemgetter('dCity'))
        my_fares = sorted(my_dictionary_list, key=sortkeypicker(['fare', 'oCity', 'dCity']))
        #print ast.literal_eval(json.dumps(my_fares))
        #returns list
        return my_fares





    tree = ET.parse('automated-flights.xml')
    root = tree.getroot()  # now get the root
    root.attrib['xmlns:ss']="urn:schemas-microsoft-com:office:spreadsheet"


    #CREATE GENERIC CLUB 49 DEALSET
    def genericClub49DealSet(which_rows, advance_purchase, start_date, end_date, upper_or_lower, calendar_start, calendar_end):
        
        dealset = ET.SubElement(root, "DealSet")
        dealset.attrib['from']= str(parseDates(getValueToTheRightOfString("Sale Start Date:")))+'T00:00:01'
        dealset.attrib['to']= str(parseDates(getValueToTheRightOfString("Purchase By:")))+'T23:59:59'


        dealinfo = ET.SubElement(dealset, "DealInfo")
        dealinfo.attrib['url']=''
        dealinfo.attrib['code']='CLUB_49_SALE'


        traveldates = ET.SubElement(dealinfo, "TravelDates")
        traveldates.attrib['startdate']= calendar_start+'T00:00:01'  
        traveldates.attrib['enddate']= calendar_end+'T23:59:59'
        #traveldates.attrib['startdate']=str(getProposedDateStart("Calendar Dates - Others"))+'T00:00:01'  
        #traveldates.attrib['enddate']=str(getProposedDateEnd("Calendar Dates - Others"))+'T23:59:59'

        dealtitle = ET.SubElement(dealinfo, "DealTitle")
        
        dealdescription = ET.SubElement(dealinfo, "DealDescrip").text = "<![CDATA[Club 49 Weekly Sale<br>Purchase by "+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+".]]>"
        if upper_or_lower == 'upper':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel within Alaska is valid '+changeDaysFont(getAvailability("Within Alaska"))+' from '+str(dateInEnglish(getTravelStart("Within Alaska")))+' - '+str(dateInEnglish(getTravelEnd("Within Alaska")))+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked.aspx">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'
        else:
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel to the US is valid '+changeDaysFont(getAvailability("To U.S."))+' from '+str(dateInEnglish(getTravelStart("To U.S.")))+' - '+str(dateInEnglish(getTravelEnd("To U.S.")))+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked.aspx">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'

        fares = ET.SubElement(dealset, "Fares")

        #This for loop will create each Row and Cell of XML for each item/dictionary in the list
        #pullFaresAndSaveInList(1, discoverSeparator()) RETURNS list of dictionaries
        for a in pullFaresAndSaveInList(which_rows):
            # print a['oCode'], a['oCity'], a['dCode'], a['dCity'],a['fare']
            row = ET.SubElement(fares, "Row") #showAsDefault="true"
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
        


    #CREATE ALASKA DEALSETS
    def alaskaDealSet(which_rows, advance_purchase, upper_or_lower):
        
        dealset = ET.SubElement(root, "DealSet")
        dealset.attrib['from']= str(parseDates(getValueToTheRightOfString("Sale Start Date:")))+'T00:00:01'
        dealset.attrib['to']= str(parseDates(getValueToTheRightOfString("Purchase By:")))+'T23:59:59'


        dealinfo = ET.SubElement(dealset, "DealInfo")
        dealinfo.attrib['url']=''
        makecode = str(parseDates(getValueToTheRightOfString("Sale Start Date:"))).replace('-', '')

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
        if upper_or_lower == 'hawaii':
            traveldates.attrib['startdate']= str(parseDates(getValueToTheRightOfString("Proposed AS.com")))+'T00:00:01'  
            traveldates.attrib['enddate']= str(parseDates(getValueToTheRightOfString("Calendar Dates - Hawaii")))+'T23:59:59'
        else:
            traveldates.attrib['startdate']= str(parseDates(getValueToTheRightOfString("Proposed AS.com")))+'T00:00:01'  
            traveldates.attrib['enddate']= str(parseDates(getValueToTheRightOfString("Calendar Dates - Others")))+'T23:59:59'


        dealtitle = ET.SubElement(dealinfo, "DealTitle")
        
        dealdescription = ET.SubElement(dealinfo, "DealDescrip").text = "<![CDATA[Purchase by "+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+".]]>"
        if upper_or_lower == 'hawaii':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel to and from Hawaii is valid '+getAvailability("To/From Hawaii")+' from '+str(dateInEnglish(getValueToTheRightOfString("Travel Start:")))+' - '+str(dateInEnglish(getValueToTheRightOfString("Complete Travel By:")))+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked.aspx">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'
       
        if upper_or_lower == 'mexico':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel to Mexico is valid '+getAvailability("AS - to Mexico")+' from '+str(dateInEnglish(getValueToTheRightOfString("Travel Start:")))+' - '+str(dateInEnglish(getValueToTheRightOfString("Complete Travel By:")))+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked.aspx">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'

        if upper_or_lower == 'costarica':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel to Costa Rica is valid '+getAvailability("AS - to Costa Rica")+' from '+str(dateInEnglish(getValueToTheRightOfString("Travel Start:")))+' - '+str(dateInEnglish(getValueToTheRightOfString("Complete Travel By:")))+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked.aspx">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'
        
        if upper_or_lower == 'florida':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel from Florida is valid '+getAvailability("From Florida")+' from '+str(dateInEnglish(getValueToTheRightOfString("Travel Start:")))+' - '+str(dateInEnglish(getValueToTheRightOfString("Complete Travel By:")))+'. Travel to Florida is valid '+getAvailability("To Florida")+' from '+str(dateInEnglish(getValueToTheRightOfString("Travel Start:")))+' - '+str(dateInEnglish(getValueToTheRightOfString("Complete Travel By:")))+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked.aspx">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'


        if upper_or_lower == 'others':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel is valid '+getAvailability("All Others")+' from '+str(dateInEnglish(getValueToTheRightOfString("Travel Start:")))+' - '+str(dateInEnglish(getValueToTheRightOfString("Complete Travel By:")))+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked.aspx">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'

        fares = ET.SubElement(dealset, "Fares")

        #This for loop will create each Row and Cell of XML for each item/dictionary in the list
        #pullFaresAndSaveInList(1, discoverSeparator()) RETURNS list of dictionaries
        for a in pullFaresAndSaveInList(which_rows):
            # print a['oCode'], a['oCity'], a['dCode'], a['dCity'],a['fare']
            row = ET.SubElement(fares, "Row") #showAsDefault="true"
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



    #AWARD SALEDEALSETS
    def awardSaleDealSet(which_rows, advance_purchase, upper_or_lower):
        
        dealset = ET.SubElement(root, "DealSet")
        dealset.attrib['from']= str(parseDates(getValueToTheRightOfString("Sale Start Date:")))+'T00:00:01'
        dealset.attrib['to']= str(parseDates(getValueToTheRightOfString("Purchase By:")))+'T23:59:59'


        dealinfo = ET.SubElement(dealset, "DealInfo")
        dealinfo.attrib['url']=''
        makecode = str(parseDates(getValueToTheRightOfString("Sale Start Date:"))).replace('-', '')

        
        if upper_or_lower == 'hawaii':
            dealinfo.attrib['code']=makecode+'_AWARD_SALE-HI'
        
        if upper_or_lower == 'mexico':
            dealinfo.attrib['code']=makecode+'_AWARD_SALE-MX'

        if upper_or_lower == 'costarica':
            dealinfo.attrib['code']=makecode+'_AWARD_SALE-CR'
        
        if upper_or_lower == 'florida':
            dealinfo.attrib['code']=makecode+'_AWARD_SALE-FL'

        if upper_or_lower == 'others':
            dealinfo.attrib['code']=makecode+'_AWARD_SALE'


        traveldates = ET.SubElement(dealinfo, "TravelDates")
        if upper_or_lower == 'hawaii':
            traveldates.attrib['startdate']= str(parseDates(getValueToTheRightOfString("Proposed AS.com")))+'T00:00:01'  
            traveldates.attrib['enddate']= str(parseDates(getValueToTheRightOfString("Calendar Dates - Hawaii")))+'T23:59:59'
        else:
            traveldates.attrib['startdate']= str(parseDates(getValueToTheRightOfString("Proposed AS.com")))+'T00:00:01'  
            traveldates.attrib['enddate']= str(parseDates(getValueToTheRightOfString("Calendar Dates - Others")))+'T23:59:59'

        dealtitle = ET.SubElement(dealinfo, "DealTitle")
        
        dealdescription = ET.SubElement(dealinfo, "DealDescrip").text = "<![CDATA[Purchase by "+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+".]]>"
        if upper_or_lower == 'hawaii':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel to and from Hawaii is valid '+getAvailability("To/From Hawaii")+' from '+str(dateInEnglish(getValueToTheRightOfString("Travel Start:")))+' - '+str(dateInEnglish(getValueToTheRightOfString("Complete Travel By:")))+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked.aspx">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'
       
        if upper_or_lower == 'mexico':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel to Mexico is valid '+getAvailability("AS - to Mexico")+' from '+str(dateInEnglish(getValueToTheRightOfString("Travel Start:")))+' - '+str(dateInEnglish(getValueToTheRightOfString("Complete Travel By:")))+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked.aspx">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'

        if upper_or_lower == 'costarica':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel to Costa Rica is valid '+getAvailability("AS - to Costa Rica")+' from '+str(dateInEnglish(getValueToTheRightOfString("Travel Start:")))+' - '+str(dateInEnglish(getValueToTheRightOfString("Complete Travel By:")))+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked.aspx">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'
        
        if upper_or_lower == 'florida':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel from Florida is valid '+getAvailability("From Florida")+' from '+str(dateInEnglish(getValueToTheRightOfString("Travel Start:")))+' - '+str(dateInEnglish(getValueToTheRightOfString("Complete Travel By:")))+'. Travel to Florida is valid '+getAvailability("To Florida")+' from '+str(dateInEnglish(getValueToTheRightOfString("Travel Start:")))+' - '+str(dateInEnglish(getValueToTheRightOfString("Complete Travel By:")))+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked.aspx">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'


        if upper_or_lower == 'others':
            terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel is valid '+getAvailability("All Others")+' from '+str(dateInEnglish(getValueToTheRightOfString("Travel Start:")))+' - '+str(dateInEnglish(getValueToTheRightOfString("Complete Travel By:")))+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked.aspx">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'

        fares = ET.SubElement(dealset, "Fares")

        #This for loop will create each Row and Cell of XML for each item/dictionary in the list
        #pullFaresAndSaveInList(1, discoverSeparator()) RETURNS list of dictionaries
        for a in pullAwardSaleFaresAndSaveInList(which_rows):
            # print a['oCode'], a['oCity'], a['dCode'], a['dCity'],a['fare']
            row = ET.SubElement(fares, "Row") #showAsDefault="true"
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
            cell = ET.SubElement(row, "Cell")
            ET.SubElement(cell, "Data").text = str(a['fees'])

        return dealset



    def exceptionDealSet(which_rows, advance_purchase, origin_code, destination_code, travel_valid, travel_start, travel_end):
        dealset = ET.SubElement(root, "DealSet")
        dealset.attrib['from']= str(parseDates(getValueToTheRightOfString("Sale Start Date:")))+'T00:00:01'
        dealset.attrib['to']= str(parseDates(getValueToTheRightOfString("Purchase By:")))+'T23:59:59'

        dealinfo = ET.SubElement(dealset, "DealInfo")
        dealinfo.attrib['url']=''
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

        terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel from '+str(m(origin_code))+'('+str(origin_code)+') to '+str(m(destination_code))+'('+str(destination_code)+')'+' is valid '+str(travel_valid)+' from '+str(dateInEnglish(getValueToTheRightOfString("Travel Start:")))+' - '+str(dateInEnglish(getValueToTheRightOfString("Complete Travel By:")))+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked.aspx">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'

        fares = ET.SubElement(dealset, "Fares")

        #This for loop will create each Row and Cell of XML for each item/dictionary in the list
        #pullFaresAndSaveInList(1, discoverSeparator()) RETURNS list of dictionaries
        for a in pullFaresAndSaveInList(which_rows):
            # print a['oCode'], a['oCity'], a['dCode'], a['dCity'],a['fare']
            row = ET.SubElement(fares, "Row") #showAsDefault="true"
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



    #CREATE ALASKA DEALSETS
    def lastminuteDealSet(which_rows, advance_purchase, upper_or_lower,):
        
        dealset = ET.SubElement(root, "DealSet")
        dealset.attrib['from']= str(parseDates(getValueToTheRightOfString("Sale Start Date:")))+'T00:00:01'
        dealset.attrib['to']= str(parseDates(getValueToTheRightOfString("Purchase By:")))+'T23:59:59'


        dealinfo = ET.SubElement(dealset, "DealInfo")
        dealinfo.attrib['url']=''
        makecode = str(parseDates(getValueToTheRightOfString("Sale Start Date:"))).replace('-', '')

        dealinfo.attrib['code']=makecode+'_LAST-MINUTE'

        traveldates = ET.SubElement(dealinfo, "TravelDates")
        traveldates.attrib['startdate']= str(parseDates(getValueToTheRightOfString("Proposed AS.com")))+'T00:00:01'  
        traveldates.attrib['enddate']= str(parseDates(getValueToTheRightOfString("Calendar Dates - Others")))+'T23:59:59'

        dealtitle = ET.SubElement(dealinfo, "DealTitle")
        
        dealdescription = ET.SubElement(dealinfo, "DealDescrip").text = "<![CDATA[Purchase by "+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+".]]>"

        blackout_dates = str(getAvailability("Blackouts:"))
        day_time = str(getValueToTheRightOfString("Day/Time Availability:").lower())

        if upper_or_lower == 'lastminute':
            if len(blackout_dates) > 0:
                terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel is valid '+day_time+' from '+str(dateInEnglish(getValueToTheRightOfString("Travel Start:")))+' - '+str(dateInEnglish(getValueToTheRightOfString("Complete Travel By:")))+'. Blackout dates are from '+blackout_dates+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked.aspx">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'
            else:
                terms = ET.SubElement(dealinfo, "terms").text = '<![CDATA[<strong>Fare Rules:</strong> Purchase by 11:59 pm (PT) on '+str(dateInEnglish(getValueToTheRightOfString("Purchase By:")))+', and at least '+str(advance_purchase)+' prior to departure. Travel is valid '+day_time+' from '+str(dateInEnglish(getValueToTheRightOfString("Travel Start:")))+' - '+str(dateInEnglish(getValueToTheRightOfString("Complete Travel By:")))+'. Bag fees <a href="#terms">may apply</a> for <a href="/content/travel-info/policies/baggage-checked.aspx">checked baggage</a>. See <a href="#terms">bottom of page</a> for full terms and conditions.]]>'

        fares = ET.SubElement(dealset, "Fares")

        #This for loop will create each Row and Cell of XML for each item/dictionary in the list
        #pullFaresAndSaveInList(1, discoverSeparator()) RETURNS list of dictionaries
        for a in pullFaresAndSaveInList(which_rows):
            # print a['oCode'], a['oCity'], a['dCode'], a['dCity'],a['fare']
            row = ET.SubElement(fares, "Row") #showAsDefault="true"
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





    if(marnel.get() == 'C49'):
        def returnMyActualDateOne(whatday):
            m = datetime.date(getYear(getTravelStart("Within Alaska")),getMonth(getTravelStart("Within Alaska")),getDay(getTravelStart("Within Alaska")))
            next_tuesday = coming_tuesday(m, whatday)
            return next_tuesday

        def returnMyActualDateTwo(whatday, howmanyweeks):
            n = datetime.date(getYear(getTravelStart("Within Alaska")),getMonth(getTravelStart("Within Alaska")),getDay(getTravelStart("Within Alaska")))
            tuesday_after = find_two_tuesday(n, whatday, howmanyweeks)
            return tuesday_after


        def getMyFirstDay(thisday):
            next_tuesday = returnMyActualDateOne(thisday)
            next_tuesday = str(next_tuesday)
            next_tuesday = next_tuesday.split(" ",1)[0]
            a1, b1, c1 = next_tuesday.split("-")
            print "Month of tuesday coming up:",b1
            #getMonth(getTravelStart("Within Alaska"))
            #next_tuesday = next_tuesday.replace("-","")
            return b1
        
        
        def getMySecondDay(thisday, howmanyweeks):
            tuesday_after = returnMyActualDateTwo(thisday, howmanyweeks)
            tuesday_after = str(tuesday_after)
            tuesday_after = tuesday_after.split(" ",1)[0]
            a2, b2, c2 = tuesday_after.split("-")
            print "Month of 2 weeks in future:",b2
            #getMonth(getTravelStart("Within Alaska"))
            #tuesday_after = tuesday_after.replace("-","")
            return b2



        if getMyFirstDay(1) == getMySecondDay(1, 21):
            print "Coming Tuesday From GIVEN DATE: ",returnMyActualDateOne(1)
            print "Two Weeks After GIVEN DATE: ",returnMyActualDateTwo(1, 21) # 21 = 2 weeks span
            calendar_start = str(returnMyActualDateOne(1))
            calendar_end = str(returnMyActualDateTwo(1, 21))
        else:
            print "Coming TUESDAY From GIVEN DATE: ",returnMyActualDateOne(1)
            print "One Week After GIVEN DATE: ",returnMyActualDateTwo(1, 14) # 14 = 1 week span
            calendar_start = str(returnMyActualDateOne(1))
            calendar_end = str(returnMyActualDateTwo(1, 14))


        pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
        pass_UpperStartDate = parseDates(getTravelStart("Within Alaska"))
        pass_UpperEndDate = parseDates(getTravelEnd("Within Alaska"))
        pass_LowerStartDate = parseDates(getTravelStart("To U.S."))
        pass_LowerEndDate = parseDates(getTravelEnd("To U.S."))


        if len(getClub49Fares("C9", 'upper')) > 0:
            print "UPPER: "+str(len(getClub49Fares("C9", 'upper')))
            print getClub49Fares("C9", 'upper')
            genericClub49DealSet(getClub49Fares("C9", 'upper'),pass_AdvancePurchase,pass_UpperStartDate,pass_UpperEndDate,"upper",calendar_start,calendar_end)


        if len(getClub49Fares("C9", 'lower')) > 0:
            print "LOWER: "+str(len(getClub49Fares("C9", 'lower')))
            print getClub49Fares("C9", 'lower')
            genericClub49DealSet(getClub49Fares("C9", 'lower'),pass_AdvancePurchase,pass_LowerStartDate,pass_LowerEndDate,"lower",calendar_start,calendar_end)
            #tree.write("\\\\seavvfile1\\Market_SAIntMktg\\_Offers\\5. In Work\\AK_Weekly Sales\\temp\\temp-xml.xml")




    #CLOSE IN FARES HANDLER
    if(marnel.get() == 'LastMinute'):
        #SERVICE EXCEPTION HANDLER
        if len(exception_codes1.get()) > 0 or len(exception_codes2.get()) > 0 or len(exception_codes3.get()) > 0 or len(exception_codes4.get()) > 0:
            print exception_codes1.get() + exception_days1.get()
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
                total_exceptions_list.append(total_exceptions_1[0])

                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print "CLOSE IN EXCEPTION ONE: "+str(len(serviceExceptionFares("AS", ex_code_1_origin, ex_code_1_destination)))
                print serviceExceptionFares("AS", ex_code_1_origin, ex_code_1_destination)
                serviceExceptionFares("AS", ex_code_1_origin, ex_code_1_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_1_origin, ex_code_1_destination),pass_AdvancePurchase,ex_code_1_origin,ex_code_1_destination,ex_code_1_days,travel_start,travel_end)

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
                total_exceptions_list.append(total_exceptions_2[0])

                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print "CLOSE IN EXCEPTION TWO: "+str(len(serviceExceptionFares("AS", ex_code_2_origin, ex_code_2_destination)))
                print serviceExceptionFares("AS", ex_code_2_origin, ex_code_2_destination)
                serviceExceptionFares("AS", ex_code_2_origin, ex_code_2_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_2_origin, ex_code_2_destination),pass_AdvancePurchase,ex_code_2_origin,ex_code_2_destination,ex_code_2_days,travel_start,travel_end)

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
                total_exceptions_list.append(total_exceptions_3[0])

                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print "CLOSE IN EXCEPTION TWO: "+str(len(serviceExceptionFares("AS", ex_code_3_origin, ex_code_3_destination)))
                print serviceExceptionFares("AS", ex_code_3_origin, ex_code_3_destination)
                serviceExceptionFares("AS", ex_code_3_origin, ex_code_3_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_3_origin, ex_code_3_destination),pass_AdvancePurchase,ex_code_3_origin,ex_code_3_destination,ex_code_3_days,travel_start,travel_end)


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
                total_exceptions_list.append(total_exceptions_4[0])

                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print "CLOSE IN EXCEPTION TWO: "+str(len(serviceExceptionFares("AS", ex_code_4_origin, ex_code_4_destination)))
                print serviceExceptionFares("AS", ex_code_4_origin, ex_code_4_destination)
                serviceExceptionFares("AS", ex_code_4_origin, ex_code_4_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_4_origin, ex_code_4_destination),pass_AdvancePurchase,ex_code_4_origin,ex_code_4_destination,ex_code_4_days,travel_start,travel_end)
        else:
            total_exceptions_list = []

        print total_exceptions_list

        if len(lastMinuteRows("AS", total_exceptions_list)) > 0:
            pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
            print "LAST MINUTE ALL OTHERS: "+str(len(lastMinuteRows("AS", total_exceptions_list)))
            print lastMinuteRows("AS", total_exceptions_list)
            lastminuteDealSet(lastMinuteRows("AS", total_exceptions_list),pass_AdvancePurchase,"lastminute")



    #NORMAL WEEKLY DEALS HANDLER
    if(marnel.get() == 'AS'):

        #SERVICE EXCEPTION HANDLER
        if len(exception_codes1.get()) > 0 or len(exception_codes2.get()) > 0 or len(exception_codes3.get()) > 0 or len(exception_codes4.get()) > 0:
            print exception_codes1.get() + exception_days1.get()
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
                total_exceptions_list.append(total_exceptions_1[0])

                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print "EXCEPTION ONE: "+str(len(serviceExceptionFares("AS", ex_code_1_origin, ex_code_1_destination)))
                print serviceExceptionFares("AS", ex_code_1_origin, ex_code_1_destination)
                serviceExceptionFares("AS", ex_code_1_origin, ex_code_1_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_1_origin, ex_code_1_destination),pass_AdvancePurchase,ex_code_1_origin,ex_code_1_destination,ex_code_1_days,travel_start,travel_end)

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
                total_exceptions_list.append(total_exceptions_2[0])

                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print "EXCEPTION TWO: "+str(len(serviceExceptionFares("AS", ex_code_2_origin, ex_code_2_destination)))
                print serviceExceptionFares("AS", ex_code_2_origin, ex_code_2_destination)
                serviceExceptionFares("AS", ex_code_2_origin, ex_code_2_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_2_origin, ex_code_2_destination),pass_AdvancePurchase,ex_code_2_origin,ex_code_2_destination,ex_code_2_days,travel_start,travel_end)

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
                total_exceptions_list.append(total_exceptions_3[0])

                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print "EXCEPTION TWO: "+str(len(serviceExceptionFares("AS", ex_code_3_origin, ex_code_3_destination)))
                print serviceExceptionFares("AS", ex_code_3_origin, ex_code_3_destination)
                serviceExceptionFares("AS", ex_code_3_origin, ex_code_3_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_3_origin, ex_code_3_destination),pass_AdvancePurchase,ex_code_3_origin,ex_code_3_destination,ex_code_3_days,travel_start,travel_end)


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
                total_exceptions_list.append(total_exceptions_4[0])

                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print "EXCEPTION TWO: "+str(len(serviceExceptionFares("AS", ex_code_4_origin, ex_code_4_destination)))
                print serviceExceptionFares("AS", ex_code_4_origin, ex_code_4_destination)
                serviceExceptionFares("AS", ex_code_4_origin, ex_code_4_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_4_origin, ex_code_4_destination),pass_AdvancePurchase,ex_code_4_origin,ex_code_4_destination,ex_code_4_days,travel_start,travel_end)
        else:
            total_exceptions_list = []

        print total_exceptions_list


        if len(hawaiiFares("AS", total_exceptions_list)) > 0:
            pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
            print "HAWAII: "+str(len(hawaiiFares("AS", total_exceptions_list)))
            print hawaiiFares("AS", total_exceptions_list)
            alaskaDealSet(hawaiiFares("AS", total_exceptions_list),pass_AdvancePurchase,"hawaii")

        if len(mexicoFares("AS", total_exceptions_list)) > 0:
            pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
            print "MEXICO: "+str(len(mexicoFares("AS", total_exceptions_list)))
            print mexicoFares("AS", total_exceptions_list)
            alaskaDealSet(mexicoFares("AS", total_exceptions_list),pass_AdvancePurchase,"mexico")

        if len(costaricaFares("AS", total_exceptions_list)) > 0:
            pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
            print "COSTA RICA: "+str(len(costaricaFares("AS", total_exceptions_list)))
            print costaricaFares("AS", total_exceptions_list)
            alaskaDealSet(costaricaFares("AS", total_exceptions_list),pass_AdvancePurchase,"costarica")

        if len(floridaFares("AS", total_exceptions_list)) > 0:
            pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
            print "FLORIDA: "+str(len(floridaFares("AS", total_exceptions_list)))
            print floridaFares("AS", total_exceptions_list)
            alaskaDealSet(floridaFares("AS", total_exceptions_list),pass_AdvancePurchase,"florida")

        if len(allOtherRows("AS", total_exceptions_list)) > 0:
            pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
            print "ALL OTHERS: "+str(len(allOtherRows("AS", total_exceptions_list)))
            print allOtherRows("AS", total_exceptions_list)
            alaskaDealSet(allOtherRows("AS", total_exceptions_list),pass_AdvancePurchase,"others")




    #AWARD SALE DEALS HANDLER
    if(marnel.get() == 'AwardSale'):

        #SERVICE EXCEPTION HANDLER
        if len(exception_codes1.get()) > 0 or len(exception_codes2.get()) > 0 or len(exception_codes3.get()) > 0 or len(exception_codes4.get()) > 0:
            print exception_codes1.get() + exception_days1.get()
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
                total_exceptions_list.append(total_exceptions_1[0])

                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print "EXCEPTION ONE: "+str(len(serviceExceptionFares("AS", ex_code_1_origin, ex_code_1_destination)))
                print serviceExceptionFares("AS", ex_code_1_origin, ex_code_1_destination)
                serviceExceptionFares("AS", ex_code_1_origin, ex_code_1_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_1_origin, ex_code_1_destination),pass_AdvancePurchase,ex_code_1_origin,ex_code_1_destination,ex_code_1_days,travel_start,travel_end)

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
                total_exceptions_list.append(total_exceptions_2[0])

                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print "EXCEPTION TWO: "+str(len(serviceExceptionFares("AS", ex_code_2_origin, ex_code_2_destination)))
                print serviceExceptionFares("AS", ex_code_2_origin, ex_code_2_destination)
                serviceExceptionFares("AS", ex_code_2_origin, ex_code_2_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_2_origin, ex_code_2_destination),pass_AdvancePurchase,ex_code_2_origin,ex_code_2_destination,ex_code_2_days,travel_start,travel_end)

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
                total_exceptions_list.append(total_exceptions_3[0])

                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print "EXCEPTION TWO: "+str(len(serviceExceptionFares("AS", ex_code_3_origin, ex_code_3_destination)))
                print serviceExceptionFares("AS", ex_code_3_origin, ex_code_3_destination)
                serviceExceptionFares("AS", ex_code_3_origin, ex_code_3_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_3_origin, ex_code_3_destination),pass_AdvancePurchase,ex_code_3_origin,ex_code_3_destination,ex_code_3_days,travel_start,travel_end)


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
                total_exceptions_list.append(total_exceptions_4[0])

                pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
                print "EXCEPTION TWO: "+str(len(serviceExceptionFares("AS", ex_code_4_origin, ex_code_4_destination)))
                print serviceExceptionFares("AS", ex_code_4_origin, ex_code_4_destination)
                serviceExceptionFares("AS", ex_code_4_origin, ex_code_4_destination)
                exceptionDealSet(serviceExceptionFares("AS", ex_code_4_origin, ex_code_4_destination),pass_AdvancePurchase,ex_code_4_origin,ex_code_4_destination,ex_code_4_days,travel_start,travel_end)
        else:
            total_exceptions_list = []

        print total_exceptions_list


        if len(hawaiiFares("AS", total_exceptions_list)) > 0:
            pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
            print "HAWAII: "+str(len(hawaiiFares("AS", total_exceptions_list)))
            print hawaiiFares("AS", total_exceptions_list)
            awardSaleDealSet(hawaiiFares("AS", total_exceptions_list),pass_AdvancePurchase,"hawaii")

        if len(mexicoFares("AS", total_exceptions_list)) > 0:
            pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
            print "MEXICO: "+str(len(mexicoFares("AS", total_exceptions_list)))
            print mexicoFares("AS", total_exceptions_list)
            awardSaleDealSet(mexicoFares("AS", total_exceptions_list),pass_AdvancePurchase,"mexico")

        if len(costaricaFares("AS", total_exceptions_list)) > 0:
            pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
            print "COSTA RICA: "+str(len(costaricaFares("AS", total_exceptions_list)))
            print costaricaFares("AS", total_exceptions_list)
            awardSaleDealSet(costaricaFares("AS", total_exceptions_list),pass_AdvancePurchase,"costarica")

        # if len(floridaFares("AS", total_exceptions_list)) > 0:
        #     pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
        #     print "FLORIDA: "+str(len(floridaFares("AS", total_exceptions_list)))
        #     print floridaFares("AS", total_exceptions_list)
        #     alaskaDealSet(floridaFares("AS", total_exceptions_list),pass_AdvancePurchase,"florida")

        if len(allOtherRows("AS", total_exceptions_list)) > 0:
            pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
            print "ALL OTHERS: "+str(len(allOtherRows("AS", total_exceptions_list)))
            print allOtherRows("AS", total_exceptions_list)
            awardSaleDealSet(allOtherRows("AS", total_exceptions_list),pass_AdvancePurchase,"others")
        

    tree.write("automated-flights.xml")


fileBtn.configure(command=getfile)
runBtn.configure(command=automate)
resBtn.configure(command=reset)

window.mainloop()
