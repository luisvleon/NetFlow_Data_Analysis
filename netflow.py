#/*------------------------------------------------------------------------
#Name:      Luis Leon
#Email:     lleon082@fiu.edu
#Date:      05/03/2021
#Project:   NetFlow Data Analysis
#----------------------------------------------------------------------------*/
import csv
import ctypes
import datetime
import os
from os import path
from tkinter import *
from tkinter import filedialog
import xlsxwriter

#function to open the log files and send to the getRecords method
def getClientsData(clientsList, file):

    with open(file) as csv_file:                                #open file to upload data
        csv_reader = csv.reader(csv_file, delimiter=',')        #use the reader method with a comma as separator
        line_count = 0                  #counter to finally record the number of lines
        for row in csv_reader:          #Iterate the reader array
            if line_count == 0:         #discriminate the first line with the header
                line_count += 1         #raise the counter by 1
            else:
                #If the record is not the header append the array to store the client
                clientsList.append([row[0],[row[1].strip(), row[2].strip(), row[3].strip(), row[4].strip()]])                
                line_count += 1          #raise the counter by 1

        #return the clients array
        return sorted(clientsList)
    
def getRecords(records, file, rec):
    with open(file) as csv_file:                                #open file to upload data
        csv_reader = csv.reader(csv_file, delimiter=',')        #use the reader method with a comma as separator
        line_count = 0                  #counter to finally record the number of lines
        for row in csv_reader:          #Iterate the reader array
            if line_count == 0:         #discriminate the first line with the header
                line_count += 1         #raise the counter by 1
            else:
                rec += 1
                # If the record is not the header append the array to store the record
                records.append([row[0], [row[1].strip(), row[2].strip(), row[3].strip(), row[4].strip(), row[5].strip(),
                                row[6].strip(), row[7].strip(), row[8].strip(), row[9].strip(), row[10].strip(),
                                row[11].strip(), row[12].strip(), row[13].strip()]])
                line_count +=1          #raise the counter by 1
    #return the records array
    return records, rec

#Function to calculate the total data transferred this week
#also gets the total records
def getTotalWeekTransfered(clients):
    total = 0       #store the total transfer data
    totalrec = 0    #store the number of records
    for indx, client in enumerate(clients):         #iterate the array
        total = total + int(clients[indx][1][0])    #add the record data transfer to the total
        totalrec += 1          #increase the counter

    return total, totalrec      #return a tuple with both values

#Funtion to get the client's host list
def getUniqueHosts(clientsanalysis, hosts):
    totalhosts = []
    counter = 0
    for client in clientsanalysis:
        for host in hosts:
            if not host[1][0] in [j for i in totalhosts for j in i]:
                if client[0] == host[0]:
                    totalhosts.append([client[0], host[1][0], host[1][1], host[1][10], 0, 0, 0, 0])
    return totalhosts

#Funtion to get the row data per host
def getTotalTransfers(records, hosts, ttotal):
    #local variables to help in string comparisons
    hos = ""
    rec = ""
    userhost = ""
    recordhost = ""

    for inxh, host in enumerate(hosts):             #Iterate the hosts array with enumeration
        hos = host[0]                               #Evaluate if the client is the same
        for inxr, record in enumerate(records):     #Iterate the host records
            rec = record[0]                         #Assign the host value to the string helper
            if hos == rec:                          #Evaluate if the host is the same
                userhost = host[1]                  #Assign the host value to the string helper
                recordhost = record[1][0]           #Assign the host value to the string helper
                if recordhost == userhost:          #Evaluate if the host is the same
                    result = int(host[4]) + int(record[1][4])       #Add the host transfer data to the total
                    counter = int(host[6]) + 1                      #Raise te counter by 1
                    rtttotal = int(host[5]) + int(record[1][9])     #Add the RTT value to the total for averaging
                    hosts[inxh][4] = str(result)    #return the final value to the record - total transfer
                    hosts[inxh][5] = str(rtttotal)  #return the final value to the record - total RTT value
                    hosts[inxh][6] = str(counter)   #return the final value to the record - total connections
        hosts[inxh][5] = round(int(hosts[inxh][5])/int(hosts[inxh][6]), 2)      #sets the average value for the RTT
        hosts[inxh][7] = round((int(hosts[inxh][4])/ttotal)*100, 2)             #sets the final percentage value of
                                                                                # transfers from the global total
    #return the array
    return hosts

#function to create an array of dictionaries from the array data
def createDict(cliRow):
    cliDict = {}
    for row in cliRow:
        if row[0] in cliDict:
            #if the client does exist append the dictionary to add the host
            cliDict[row[0]].append([row[1], row[2], row[3], row[4], row[5], row[6], row[7]])
        else:
            #if the client does not existe, create the dictionary
            cliDict[row[0]] = []
            cliDict[row[0]].append([row[1], row[2], row[3], row[4], row[5], row[6], row[7]])
            
    #returns the dictionary
    return cliDict

#Function header
#Function to display the program greeting
def showHeader():
    print("*********************************************************************\n")
    print("******                NetFlow Analysis Program            ***********\n")
    print("*********************************************************************\n")
    print("                                                                     \n")
    print("      Please follow the Graphical User Interface instructions        \n")

#Function to write the report to a text file
def writeReport(clientPerHost, maxcli):

    x = datetime.datetime.now()     #Create a datetime variable and get the date and time
    datesheet = str(x.year) + "_" + str(x.month) + "_" + str(x.day)
    savepath = filedialog.askdirectory()

    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook(savepath + '/NetFlow_Analysis.xlsx')
    worksheet = workbook.add_worksheet(datesheet)

    #start at line No. 3
    line = 4
    #Store clients data dictionary
    PieClients = {}

    #Store the partial total percentage in case it is not a full report
    partialpercentage = 0

    # Sets the columns width 
    worksheet.set_column('B:B', 22)
    worksheet.set_column('C:C', 20)
    worksheet.set_column('D:D', 18)
    worksheet.set_column('E:E', 15)
    worksheet.set_column('F:F', 9)
    worksheet.set_column('G:G', 11)
    worksheet.set_column('H:H', 12)
    worksheet.set_column('I:I', 13)

    #create a format for individual cells
    cell_format = workbook.add_format()
    cell_format.set_align('center')

    # create a format for small windows size
    cell_format_red = workbook.add_format()
    cell_format_red.set_align('center')
    cell_format_red.set_bg_color('orange')

    # create a format for medium windows size
    cell_format_yellow = workbook.add_format()
    cell_format_yellow.set_align('center')
    cell_format_yellow.set_bg_color('yellow')


    #create a defined format for the document title
    cell_format_title = workbook.add_format()
    cell_format_title.set_font_size(30)
    cell_format_title.set_bold()

    #format for the subtitles
    cell_format_subtitle = workbook.add_format()
    cell_format_subtitle.set_font_size(22)
    cell_format_subtitle.set_bold()

    #Create a dictionary to handle the data
    reportDict = createDict(clientPerHost)

    #Starts the maxclients counter
    maxclicounter = 0
    #Iterate the dictionary to get the client names
    for name in reportDict.keys():

        if maxclicounter < maxcli:

            # Add a bold format to use to highlight cells.
            bold = workbook.add_format({'bold': True})
            line += 1
            worksheet.write('B' + str(line), 'Client: ' + name, bold)
            PieClients[name] = [line,0]
            line += 1

            # Write the header row.
            worksheet.write('B2', 'NetFlow Data Analysis', cell_format_title)
            worksheet.write('B3', 'Report for: ___________________', cell_format_subtitle)
            worksheet.write('B' + str(line), 'Hostname', bold)
            worksheet.write('C' + str(line), 'IP', bold)
            worksheet.write('D' + str(line), 'Window Size (bits)', bold)
            worksheet.write('E' + str(line), 'Transfered (TB)', bold)
            worksheet.write('F' + str(line), 'RTT AVG', bold)
            worksheet.write('G' + str(line), 'Connections', bold)
            worksheet.write('H' + str(line), '% from total', bold)
            line += 1

            k = 0                   #Variable to store the client group number of iterations
            totalGroup = 0          #valriable to store the total percentages from the total transfers

            maxclicounter += 1

            for hostname in reportDict[name]:

                k += 1              #raise the number of connections per client per host
                totalGroup = totalGroup + float(hostname[6])        #add the total transferred data value

                #write the row data per cell
                worksheet.write('B' + str(line), hostname[0])
                worksheet.write('C' + str(line), hostname[1])

                if float(hostname[2]) <= float(entry_size_low.get()):
                    worksheet.write('D' + str(line), float(hostname[2]), cell_format_red)
                elif float(hostname[2]) <= float(entry_size_medium.get()):
                    worksheet.write('D' + str(line), float(hostname[2]), cell_format_yellow)
                else:
                    worksheet.write('D' + str(line), float(hostname[2]), cell_format)

                worksheet.write('E' + str(line), round(int(hostname[3]) / 1000000000), cell_format)
                worksheet.write('F' + str(line), hostname[4], cell_format)
                worksheet.write('G' + str(line), int(hostname[5]), cell_format)
                worksheet.write('H' + str(line), hostname[6], cell_format)
                partialpercentage += hostname[6]
                positiondata = PieClients.get(name)

                positiondata[1] = line
                #add one line
                line += 1

                if k > 1:       #if the host count is greater than one write also the sum of the total host transfers percentages
                    worksheet.write('H' + str(line), totalGroup, cell_format)
                    positiondata = PieClients.get(name)
                    positiondata[1] = line

    #If it is a partial report from the top clients then create an others client's value
    if partialpercentage < 100:
        # add one line
        line += 1
        #Write a others client name
        worksheet.write('B' + str(line), 'Client: ' + 'Others', bold)
        #Write the header for the others %
        worksheet.write('H' + str(line), '% from total', bold)
        line += 1
        # Write the value from 100 less the all reported clients together
        worksheet.write('H' + str(line), (100 - partialpercentage), cell_format)
        #add manually the others value for the pie chart
        PieClients['Others'] = [(line - 1), line]

    # Create a new chart object.
    chart1 = workbook.add_chart({'type': 'pie'})

    #create the string for categories
    categoriesstr = '=('
    valuesstr = '=('

    #iterate the dictionary to get individual values
    for val in PieClients.values():
        categoriesstr = categoriesstr + datesheet + '!$B$' + str(val[0]) + ','
        valuesstr = valuesstr + datesheet + '!$H$' + str(val[1]) + ','

    categoriesstr = categoriesstr[:-1] + ')'
    valuesstr = valuesstr[:-1] + ')'
    
    # Configure the series. Note the use of the dictionary syntax to define ranges:
    chart1.add_series({
        'name':       'Clients',
        'categories': categoriesstr,
        'values':     valuesstr,
        'data_labels': {'value': True},

        'data_labels': {'percentage': True},
    })

    # Add a title.
    if partialpercentage == 100:
        chart1.set_title({'name': 'All Network Talkers'})
    else:
        chart1.set_title({'name': 'Top ' + str(maxcli) + ' Network Talkers'})
    chart1.set_size({'width': 720, 'height': 500})

    # Set an Excel chart style. Colors with white outline and shadow.
    chart1.set_style(10)
    chart1.set_rotation(90)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('I5', chart1, {'x_offset': 25, 'y_offset': 10})

    #Close the workbook
    workbook.close()

#Main flow of the program
def main(usersfile, recordsfile, maxclireport):
    # Variable declaration
    clients = []  # clients ArrayList
    totalRecords = []  # Store the individual records
    uniqueHosts = []  # Nested array to store records by host
    totalWeekTransfered = 0  # Get the total week data transfer value
    totalclients = 0
    totallines = 0

    showHeader()  # show program gretting
    getClientsData(clients, usersfile)  # collectFileNames()
    (totalRecords, totallines) = getRecords(totalRecords, recordsfile, totallines)  # get transfer records

    # get the global amount of data transfer
    (totalWeekTransfered, totalclients) = getTotalWeekTransfered(clients)
    # Get the records by host
    uniqueHosts = getUniqueHosts(clients, totalRecords)
    # Create a data array per client and per host
    clientsPerHost = getTotalTransfers(totalRecords, uniqueHosts, totalWeekTransfered)
    # sort the array by descending order based on the percentage transfered from the total transfers
    clientsPerHost.sort(key=lambda x: x[7], reverse=True)

    #if the total clients analysed is greater than the clients report limit then
    if totalclients > maxclireport:
        maxcli = totalclients

    # write the report on a text file
    writeReport(clientsPerHost, maxclireport)
    print("Total clients processed: " + str(totalclients))
    print("Total records: " + str(totallines))
    labelFilesOk.config(text="Output file created!", fg="green")
    ctypes.windll.user32.MessageBoxW(0, "Your Excel report was successfully created !", "NetFlow Data Analysis", 1)
    buttonExecute.config(state=DISABLED)

#function to load default path values
def loadUserPath():

    #you can modify the file names that should be inthe working directory wan.1.txt and wan.2.txt
    userslogfile = "wan.1.txt"      #<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    datalogfile = "wan.2.txt"       #<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    #Stores the centinel bolean variable
    filesok = False

    #Gets the working directory with the data files
    #You could hard code the working directory path here:       <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    #pathToFiles = 'some/directory/inside/your/home'

    #or use the filedialog to get it
    pathToFiles = filedialog.askdirectory()
    
    #test for the existence of the working files and asign the names by default
    if path.exists(os.path.join(pathToFiles, userslogfile)) and path.exists(os.path.join(pathToFiles, datalogfile)):            
        
        labelUsersPath.config(text=os.path.join(pathToFiles, userslogfile))      
        labelDataPath.config(text=os.path.join(pathToFiles, datalogfile))      
        labelFilesOk.config(text="File names verified!", fg="green")

        #Enables the process button
        buttonExecute.config(state=NORMAL)

        #If the files where correctly processed then exit the bucle
        filesok = True
    else:
        labelFilesOk.config(text="Files does not exist in the selected path!", fg="red")

#funtion to start the main function
#this funcion call is the one actually starts the processes
#sending the 2 file names and path of the 2 working log files
def createxcelfile():


    main(labelDataPath['text'], labelUsersPath['text'], int(entryMax.get()))


##########################################################################################################
#The program starts here.
#Once the GUI programs starts, the button starts the process

#Create GUI objects
root = Tk()
root.title("NetFlow Analysis Script")

#sets the GUI size
root.geometry("850x400")

#Creates the title label
labelTitle = Label(root, text="NetFlow Data Analysis", font=("Arial", 25))
labelTitle.grid(row=0, column=0, columnspan=3, padx=200, pady=20)

##################################################################################
#   you can define in the nex lines the default low and medium window TCP size   #
#   to highlight on orange and yellow colors the low values
##################################################################################

low_tcp_size = 212992
medium_tcp_size = 16777216
default_max_clients = 10

##################################################################################

#Creates the max clients label
labelMax = Label(root, text="Max clients")
labelMax.grid(row=1, column=0, columnspan=1)

#create a input box
entryMax = Entry(root, justify=CENTER, relief=SUNKEN)
entryMax.insert(END, default_max_clients)
entryMax.grid(row=1, column=1)

#Low window size label
label_size_low = Label(root, text="Low TCP size")
label_size_low.grid(row=2, column=0)

#Low size input box
entry_size_low = Entry(root, justify=CENTER, relief=SUNKEN)
entry_size_low.insert(END, low_tcp_size)
entry_size_low.grid(row=2, column=1)

#Medium window size label
label_size_medium = Label(root, text="Medium TCP size")
label_size_medium.grid(row=3, column=0)

#Medium size input box
entry_size_medium = Entry(root, justify=CENTER, relief=SUNKEN)
entry_size_medium.insert(END, medium_tcp_size)
entry_size_medium.grid(row=3, column=1)

#create two buttons
buttonUsers = Button(root, text="Select users file", command=loadUserPath)
buttonUsers.grid(row=4, column=0, padx=10, pady=10)
buttonExecute = Button(root, text="Select destination folder", command=createxcelfile, state=DISABLED)
buttonExecute.grid(row=8, column=0, padx=10, pady=10)

#create instances of several labels
labelFileUsers = Label(root, text="Users file: ")
labelFileUsers.grid(row=5, column=0, padx=10, pady=10)
labelUsersPath = Label(root, text="")
labelUsersPath.grid(row=5, column=1, padx=10, pady=10)
labelFileData = Label(root, text="Data file: ")
labelFileData.grid(row=6, column=0, padx=10, pady=10)
labelDataPath = Label(root, text="")
labelDataPath.grid(row=6, column=1, padx=10, pady=10)
labelFilesOk = Label(root, text="")
labelFilesOk.grid(row=7, column=0, padx=10, pady=10)

# Button for closing
exit_button = Button(root, text="<<<<  Exit  >>>>", command=root.destroy)
exit_button.grid(row=8, column=2, padx=10, pady=10)

#sets the main loop for tkinter
root.mainloop()
