from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from operator import itemgetter
import ScrolledText
import sys
import tkFileDialog
import Tkinter
import xlrd
import json, ast, os, string, random, urllib
import xml.etree.cElementTree as ET
import datetime
import dateparser
import tkMessageBox
 
 
########################################################################
class RedirectText(object):
    """"""
 
    #----------------------------------------------------------------------
    def __init__(self, text_ctrl):
        """Constructor"""
        self.output = text_ctrl
 
    #----------------------------------------------------------------------
    def write(self, string):
        """"""
        self.output.insert(Tkinter.END, string)
 
 
########################################################################
class MySale(object):
    """"""
    
    #----------------------------------------------------------------------
    def __init__(self, parent):
        """Constructor"""
        self.root = parent
        self.root.title("Fare Sales")
        self.frame = Tkinter.Frame(parent)
        self.frame.pack()

        

        var = StringVar(parent)
        var.set("Select Type") # initial value

        top = parent
        l1 = Label(top, text="Exceptions")
        l1.pack(side=LEFT)
        e1 = Entry(top, bd=5)
        e1.pack(side=LEFT)
        
        
        option = OptionMenu(parent, var, "Select Type", "C49-Upper", "C49-Lower", "AS-Hawaii", "AS-Mexico", "AS-Others", "VX-Hawaii", "VX-Others")
        option.pack(side=TOP)


        #def ok():
            #print "You selected: ", var.get()
        #button = Button(parent, text="RUN", command=ok)
        #button.pack()


        #Text Area for OUTPUT
        self.text = ScrolledText.ScrolledText(self.frame)
        self.text.pack()

        
        
        # redirect stdout
        redir = RedirectText(self.text)
        sys.stdout = redir


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


       # d = datetime.datetime.now() # date(2017, 3, 13) year,month,day
        #next_tuesday = coming_tuesday(d, 1) # 0 = Monday, 1=Tuesday, 2=Wednesday...
       # next_tuesday = str(next_tuesday)
       # next_tuesday = next_tuesday.split(" ",1)[0]
       # next_tuesday = next_tuesday.replace("-","")

       # print "Upcoming Tuesday will be: ",next_tuesday





        def select_file():
            #filename =  filedialog.askopenfilename(initialdir = "//seavvfile1/Market_SAIntMktg/_Offers/5. In Work/AK_Weekly Sales/temp/",title = "Select file",filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))
            filename =  filedialog.askopenfilename(initialdir = "C:/Users/v-mmangrub/Desktop/PYTHON PROGRAMS/data-to-read",title = "Select file",filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))


            print "FILEPATH: ",filename
            print "You selected GROUP: ",var.get()

            #Checks if a group type is selected, if not open message box with error message
            if(var.get() == 'Select Type'):
                tkMessageBox.showinfo("Error", "You forgot to select a group type!")


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
                

            #Check to see if there are any Fares with International DEPARTURES
            def hawaiiFares(airline_type):
                hawaii_codes = ["OGG","LIH","KOA","HNL"]
                hawaii_list = []
                for col in range(5,7):
                    for row in range(2, sheet.nrows):
                        if sheet.cell_value(row, 5) == airline_type:
                            if sheet.cell_value(row, 7) in hawaii_codes or sheet.cell_value(row, 9) in hawaii_codes:
                                hawaii_list.append(row+1)
                            else:
                                continue
                        else:
                            my_hawaii_fares = hawaii_list
                            continue
                    return my_hawaii_fares

                
            
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
            def allOtherRows(airline_type):
                combined_hawaii_and_international = ["OGG","HNL","LIH","KOA","MEX","CUN","GDL","LTO","SJD","ZLO","MZT","PVR","ZIH","LIR","SJO","HAV"]
                others_list = []
                for col in range(5,7):
                    for row in range(1, sheet.nrows):
                        if sheet.cell_value(row, 5) == airline_type:
                            if sheet.cell_value(row, 7) not in combined_hawaii_and_international and sheet.cell_value(row, 9) not in combined_hawaii_and_international:
                                others_list.append(row+1)
                            else:
                                continue
                        else:
                            all_other_fares = others_list
                            continue
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
                #print ast.literal_eval(json.dumps(my_fares))
                #returns list
                return my_fares





            tree = ET.parse('C:\\Users\\v-mmangrub\\Desktop\\PYTHON PROGRAMS\\data-to-read\\output.xml')
            root = tree.getroot()  # now get the root
            root.attrib['xmlns:ss']="urn:schemas-microsoft-com:office:spreadsheet"


            #CREATE GENERIC DEALSET
            def genericDealSet(which_rows, advance_purchase, start_date, end_date, upper_or_lower, calendar_start, calendar_end):
                
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
                
    

    
            

            #Books and Sheets
            book = xlrd.open_workbook(filename)
            sheet_one = book.sheet_by_index(0)
            sheet = book.sheet_by_index(3)

            #CLUB 49 DETAILS
            print "Sale Start Date = ",parseDates(getValueToTheRightOfString("Sale Start Date:"))
            print "Purchase By = ",parseDates(getValueToTheRightOfString("Purchase By:"))
            print "Advance Purchase = ",getValueToTheRightOfString("Advance Purchase:")
            print "START DATE TO US = ",parseDates(getTravelStart("To U.S."))
            print "START DATE WITHIN ALASKA = ",parseDates(getTravelStart("Within Alaska"))
            print "END DATE TO US = ",parseDates(getTravelEnd("To U.S."))
            print "END DATE WITHIN ALASKA = ",parseDates(getTravelEnd("Within Alaska"))
            print "AVAILABILITY TO US = ",getAvailability("To U.S.")
            print "AVAILABILITY WITHIN ALASKA = ",getAvailability("Within Alaska")
            pass_AdvancePurchase = getValueToTheRightOfString("Advance Purchase:")
            pass_UpperStartDate = parseDates(getTravelStart("Within Alaska"))
            pass_UpperEndDate = parseDates(getTravelEnd("Within Alaska"))
            pass_LowerStartDate = parseDates(getTravelStart("To U.S."))
            pass_LowerEndDate = parseDates(getTravelEnd("To U.S."))
            #CLUB 49 DETAILS
            


            
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




            if(var.get() == 'C49-Lower'):
                print getClub49Fares("C9", 'lower')
                genericDealSet(getClub49Fares("C9", 'lower'),pass_AdvancePurchase,pass_UpperStartDate,pass_UpperEndDate,"lower",calendar_start,calendar_end)
                tree.write("C:\\Users\\v-mmangrub\\Desktop\\PYTHON PROGRAMS\\data-to-read\\output.xml")
                #tree.write("\\\\seavvfile1\\Market_SAIntMktg\\_Offers\\5. In Work\\AK_Weekly Sales\\temp\\temp-xml.xml")

            if(var.get() == 'C49-Upper'):
                print getClub49Fares("C9", 'upper')
                genericDealSet(getClub49Fares("C9", 'upper'),pass_AdvancePurchase,pass_UpperStartDate,pass_UpperEndDate,"upper",calendar_start,calendar_end)
                tree.write("C:\\Users\\v-mmangrub\\Desktop\\PYTHON PROGRAMS\\data-to-read\\output.xml")
                #tree.write("\\\\seavvfile1\\Market_SAIntMktg\\_Offers\\5. In Work\\AK_Weekly Sales\\temp\\temp-xml.xml")


            
            if(var.get() == 'AS-Hawaii'):
                print hawaiiFares("AS")
                
            if(var.get() == 'AS-Mexico'):
                print internationalDepartureCheck("AS")

            #if(var.get() == 'AS-Others'):
                


            if(var.get() == 'VX-Hawaii'):
                print hawaiiFares("VX")

                
            if(var.get() == 'VX-Others'):
                print allOtherRows("VX")


            
            






                
            
            



        #Menu Bar Options
        menu = Menu(root)
        root.config(menu=menu)
        file = Menu(menu)
        #file.add_command(label = 'Open', command = select_file)
        file.add_command(label = 'Exit', command = self.close_window)
        menu.add_cascade(label = 'File', menu = file)
        
        btn = Button(parent, text="Select File", command=select_file)
        #cls = Tkinter.Button(self.frame, text="Close", command=self.close_window)
        #cls.pack()
        btn.pack(side=TOP)



    def close_window(self):
        global root
        root.destroy()
        
    #----------------------------------------------------------------------
    def open_file(self):
        """
        Open a file, read it line-by-line and print out each line to
        the text control widget
        """
        options = {}
        options['defaultextension'] = '.txt'
        options['filetypes'] = [('all files', '.*'), ('text files', '.txt')]
        options['initialdir'] = '//seavvfile1/Market_SAIntMktg/_Offers/5. In Work'
        options['parent'] = self.root
        options['title'] = "Open a file"
 
        with tkFileDialog.askopenfile(mode='r', **options) as f_handle:
            print options
            #for line in f_handle:
                #print line
   
        
#----------------------------------------------------------------------
if __name__ == "__main__":
    root = Tkinter.Tk()
    root.geometry("800x600")

    
    app = MySale(root)
    
    root.mainloop()
