#required in order to use 'string.ascii_lowercase'
import string
import time

#logging
Settings.UserLogs = True
Settings.UserLogPrefix = 'user'
Settings.UserLogTime = True

#global variables
userDetails = ['number', 'surname', 'forname', 'gender', 'dob', 'feestat', 'MRC_course', 'SRS_course', 'block', 'occurrence', 'year', 'mod', 'SCE_UDF6', 'NAT_GEGC']
sitsVersion = ['Dev', 'Test']
col = 0
row = 3
maxRow = 0
refDoc = None
user = None
QAS = None
CAPS = None
ATR = None
ATRfail = None
timeStamp = time.strftime("%H.%M")

#multiple TABs
def doTab(n):
    i = 0
    while i < n:
        type (Key.TAB)
        i += 1

#multiple Down Keys
def doKeyDown(n):
    i = 0
    while i < n:
        type (Key.DOWN)
        i += 1

#open SITS Dev or Test and prompt to log in
def openSITS():
    SITS = select('Dev or Test?', 'SITS Version', options = sitsVersion)
    if SITS == 'Dev':
        App.open('\\\\atwin-sitst2\\sits_clients\\vision\\dev\\uniface\\bin\\uniface.exe /asn=\\\\atwin-sitst2\\sits_clients\\vision\\dev\\adm\\uclisipr.asn winlog logo=\\\\atwin-sitst2\\sits_clients\\vision\\dev\\adm\\SITS-Login-LIVE.png')
        wait("1524838031037.png", FOREVER)
        wait(0.5)
        popup('Please enter your SITS Dev password')
        wait("1524653569015.png", FOREVER)
        
    if SITS == 'Test':
        App.open('\\\\atwin-sitsd2\\sits_clients\\vision\\live\\uniface\\bin\\uniface.exe /asn=\\\\atwin-sitsd2\\sits_clients\\vision\\live\\adm\\\uclisipr.asn winlog')
        wait("1524838031037.png", FOREVER)
        wait(0.5)
        popup('Please enter your SITS Test password')
        wait("1524653569015.png", FOREVER)

#function to open the reference excel doc and prompt to re-enter doc name if not found
def openRefDoc(refDoc):
    i = 0
    while i < 1:
        type('r', KeyModifier.WIN)
        paste(refDoc)
        type(Key.ENTER)
        wait(2)
        if exists("1524643294341.png" or "1524730349042.png"):
            i = 1
            Debug.user('%s opened',refDoc)
        
        elif exists("1524730602718.png"):
            Debug.user('%s not found',refDoc)
            type(Key.ENTER)
            popup("file not found")
            refDoc = input('Reference Excel Document name (without file extension)')
            refDoc = '%userprofile%\\documents\\' + refDoc + '.xlsx'


#function to store the contents of an entire specific row in the 'userDetails' list (columns A - J)
def storeDetails(x, y):
    App.focus('Excel')
    while (x <= 12):
        copyCell(string.ascii_lowercase[x], y)
        userDetails[x] = Env.getClipboard()
        x += 1
    #Debug.user(userDetails[0] + ' - data stored in memory')
    global row
    row += 1

#function to copy the contents of a specific cell to the clipboard
def copyCell(x, y):
    type('g', Key.CTRL)
    type(str(x) + str(y) + Key.ENTER)
    wait(0.05)
    type('c', KeyModifier.CTRL)
    wait(0.15)

#function to enter QAS details
def enterQASDetails():
    if exists("1524653530693.png"):
        exit
    else:
        if exists("1524653569015.png"):
            type('qas' + Key.TAB)
            wait("1524653530693.png")
            exit
    type(Key.F12)
    wait(0.3)
    type(Key.ENTER)
    paste(userDetails[6])#MRC course
    doTab(1)
    wait(1)
    if exists("1524748542437.png"):
        type(Key.TAB + Key.ENTER)
    type(Key.F7)
    paste(userDetails[8])#BLOK   
    doTab(1)
    type(Key.F7)
    paste(userDetails[9])#Occurrence
    doTab(1)
    type(Key.F7)
    paste(userDetails[11])#Mode of Attendance
    doTab(3)
    type(Key.F7)
    paste(userDetails[10])#Year
    doTab(2)
    if exists("1524748542437.png"):
        type(Key.TAB + Key.ENTER)
    doTab(2)
    if exists("1524748542437.png"):
        type(Key.TAB + Key.ENTER)
    doTab(4)
    paste(userDetails[0])#Number
    doTab(1)
    type(Key.F7)
    paste(userDetails[1])#Surname
    doTab(1)
    type(Key.F7)
    paste(userDetails[2])#Forename
    doTab(2)
    type(Key.F7)
    paste(userDetails[3])#Title
    doTab(1)
    wait(0.2)
    doTab(1)
    type(Key.F7)
    paste(userDetails[4])#DOB
    doTab(1)
    type(Key.F7)
    paste(userDetails[5])#Fee Status
    click("1524745563338.png")
    wait(2)

#QAS Success Scenarios
    if exists("1525960954443.png",0.2):
        doTab(2)
        type(Key.ENTER)
        wait("1525951858086.png",FOREVER)
        global QAS
        QAS = 'SUCCESS'
        Debug.user(userDetails[0] + ' - QAS COMPLETED successfully')
        wait(1)
        exit

    elif exists("1524746160161.png",0.2):
        click("1524746181629.png")
        wait("1525951858086.png",FOREVER)
        global QAS
        QAS = 'SUCCESS'
        Debug.user(userDetails[0] + ' - QAS COMPLETED successfully')
        wait(1)
        exit

    elif exists("1524747688423.png",0.2):
        type(Key.ENTER)
        global QAS
        QAS = 'SUCCESS'
        Debug.user(userDetails[0] + ' - QAS SUCCESS (Student Number Already Exists, 2nd CAP Record Created)')
        exit

    elif exists("1524747688423.png",0.2):
        type(Key.ENTER)
        global QAS
        QAS = 'SUCCESS'
        Debug.user(userDetails[0] + ' - QAS SUCCESS (Student Number Already Exists, 2nd CAP Record Created)')
        exit

    elif exists("1527158069042.png",0.2):
        type(Key.ENTER)
        global QAS
        QAS = 'SUCCESS'
        Debug.user(userDetails[0] + ' - QAS SUCCESS (Student Number Already Exists, 2nd CAP Record Created)')
        exit

#QAS Failure Scenarios
    elif exists("1524745907937.png",0.2):
        type(Key.ENTER)
        global QAS
        QAS = 'FAIL'
        Debug.user(userDetails[0] + ' - QAS FAILED (UCAS course)')
        exit

    elif exists("1525951561125.png",0.2):
        type(Key.TAB)
        type(Key.ENTER)
        global QAS
        QAS = 'FAIL'
        Debug.user(userDetails[0] + ' - QAS FAILED (Mode of Attencance missing)')
        exit

    elif exists("1525951615389.png",0.2):
        type(Key.TAB)
        type(Key.ENTER)
        global QAS
        QAS = 'FAIL'
        Debug.user(userDetails[0] + ' - QAS FAILED (Academic Year missing)')
        exit

    elif exists("1525951389270.png",0.2):
        type(Key.TAB)
        type(Key.ENTER)
        global QAS
        QAS = 'FAIL'
        Debug.user(userDetails[0] + ' - QAS FAILED (Course ID not Valid)')
        exit

    elif exists("1525951465804.png",0.2):
        type(Key.TAB)
        type(Key.ENTER)
        global QAS
        QAS = 'FAIL'
        Debug.user(userDetails[0] + ' - QAS FAILED (Surname or DoB missing)')
        exit

    elif exists("1525951515498.png",0.2):
        type(Key.TAB)
        type(Key.ENTER)
        global QAS
        QAS = 'FAIL'
        Debug.user(userDetails[0] + ' - QAS FAILED (DoB Invalid)')
        exit

    else:
        global QAS
        QAS = 'FAIL'
        Debug.user(userDetails[0] + ' - QAS FAILED (unknown error)')
        exit
    
    type(Key.F4) #close QAS

#function to update CAP Details
def enterCAPSDetails():
    if exists("1525255114352.png"):
        exit
    else:
        if exists("1524653569015.png"):
            type('CAPS' + Key.TAB)
            wait("1525255114352.png")
            exit
    type(Key.F12)
    wait(0.3)
    paste(userDetails[0])#ID
    doTab(15)
    paste(userDetails[6])#MRC course
    type(Key.F5)
    wait(0.5)
    doTab(7)
    type('QV')#Qualification Status (Qualified)
    doTab(13)
    type('U')#Decision (Unconditional Offer)
    doTab(1)
    type('F')#Response (Firm)
    doTab(1)
    type('s', KeyModifier.CTRL)
    wait(1)

#CAPS Success Scenario
    if exists("1525952379949.png",0.2):
        global CAPS
        CAPS = 'SUCCESS'
        Debug.user(userDetails[0] + ' - CAPS UPDATED')
        wait(1)
        exit

#CAPS Failure Scenario
    elif exists("1525952141546.png",0.2):
        global CAPS
        CAPS = 'FAIL'
        Debug.user(userDetails[0] + ' - CAPS FAILED (no changes made)')
        exit

    type(Key.F4) #close CAPS

#function to update ATR Details
def enterATRDetails():
    if exists("1525952654814.png"):
        exit
    else:
        if exists("1524653569015.png"):
            type('ATR' + Key.TAB)
            wait("1525952654814.png")
            exit
    paste(userDetails[10])#Year
    doTab(4)
    paste(userDetails[0])#Number
    doTab(5)
    paste(userDetails[6])#MRC course
    type(Key.F5)
    wait(0.3)
    click("1525953126635.png")
    wait(1)
    
#ATR Success Scenario
    if exists("1525955213623.png",4):
        global ATR
        ATR = 'SUCCESS'
        Debug.user(userDetails[0] + ' - ATR COMPLETED Successfully')
        wait(1)
        exit

#ATR Failure Scenario
    elif exists("1525953285531.png",0.2):
        type('m', KeyModifier.CTRL) #open Message buffer
        #start triple click#
        click("1525959401097.png")
        mouseDown(Button.LEFT)
        mouseUp(Button.LEFT)
        wait(0.01)
        mouseDown(Button.LEFT)
        mouseUp(Button.LEFT)
        #end triple click#
        type('c', KeyModifier.CTRL)
        ATRfail = Env.getClipboard()
        type(Key.F4) #close Message buffer
        global ATR
        ATR = 'FAIL'
        Debug.user(userDetails[0] + ' - ATR FAILED (' + ATRfail + ')')
        exit

    type(Key.F4) #close CAPS
    

#function to update SCE Details
def enterSCEDetails():
    if exists("1525949870304.png"):
        exit
    else:
        if exists("1524653569015.png"):
            type('SCE' + Key.TAB)
            wait("1525949870304.png")
            exit
    type(Key.F12)
    wait(0.3)
    paste(userDetails[0])#ID
    click("1525963718523.png")
    doKeyDown(12)
    type(Key.ENTER)
    type(Key.ENTER)
    doTab(4)
    paste(userDetails[7])#SRS course
    doTab(1)
    paste(userDetails[8])#BLOK
    doTab(1)
    paste(userDetails[10])#Year
    doTab(1)
    paste(userDetails[9])#Occurrence
    type(Key.F5)
    wait(0.1)
    click("1525962023187.png")
    wait(1.5)
    if SITS == 'Dev':
        doTab(5)
        exit
    elif SITS == 'Test':
        doTab(2)
        exit
    wait(0.5)
    paste(userDetails[12])#SCE_UDF6 New/Returning
    wait(0.5)
    click("1525962494090.png")
    wait(0.5)
    type('s', KeyModifier.CTRL) #save SCE
    wait(1)

#SCE Success Scenario
    if exists("1525952379949.png",0.5):
        Debug.user(userDetails[0] + ' - SCE UPDATED')
        exit

#SCE Failure Scenario
    elif exists("1525952141546.png",0.5):
        Debug.user(userDetails[0] + ' - SCE FAILED (no changes made)')
        exit

#SCE Failure Scenario
    elif exists("1526569766560.png",0.5):
        Debug.user(userDetails[0] + ' - SCE FAILED (no changes made)')
        exit
    
    type(Key.F4) #close SCE

#User input to gather username (only used to create relitive path for loggin)
user = input('Please ener your username')

Debug.setUserLogFile('C:\\users\\' + user + '\\documents\\autoSITS-' + timeStamp + '.log')

Debug.user('BEGIN')

#User input to specify the reference document name and modify to
refDoc = input('Reference Excel Document name (without file extension)')
refDoc = '%userprofile%\\documents\\' + refDoc + '.xlsx'

#Open the reference document
openRefDoc(refDoc)

maxRow = int(input('number of rows to process:'))
maxRow += 2

#User input to select SITS Version
SITS = select('Dev or Test?', 'SITS Version', options = sitsVersion)

#Open SITS Dev
if SITS == 'Dev':
    openApp('\\\\atwin-sitst2\\sits_clients\\vision\\dev\\uniface\\bin\\uniface.exe /asn=\\\\atwin-sitst2\\sits_clients\\vision\\dev\\adm\\uclisipr.asn winlog logo=\\\\atwin-sitst2\\sits_clients\\vision\\dev\\adm\\SITS-Login-LIVE.png')
    wait("1524838031037.png", FOREVER)
    Debug.user('SITS Dev opened')
    popup("Please log in")
    wait("1524653569015.png", FOREVER)
    Debug.user(user + ' logged in to SITS Dev')

#Open SITS Test 
if SITS == 'Test':
    openApp('\\\\atwin-sitsd2\\sits_clients\\vision\\live\\uniface\\bin\\uniface.exe /asn=\\\\atwin-sitsd2\\sits_clients\\vision\\live\\adm\\\uclisipr.asn winlog')
    wait("1524838031037.png", FOREVER)
    Debug.user('SITS Test opened')
    popup('Please log in')
    wait("1524653569015.png", FOREVER)
    Debug.user(user + ' logged in to SITS Test')

#Store row of data from the reference document, enter data into QAS then loop for the number of rows specified
while (row <= maxRow):
    storeDetails(col, row)
    App.focus('SITS:Vision')
    enterQASDetails()
    if QAS == 'SUCCESS':
        enterCAPSDetails()
        if CAPS == 'SUCCESS':
            enterATRDetails()
            if ATR == 'SUCCESS':
                enterSCEDetails()

Debug.user('END')
popup('Program completed successfully. Please check the log file for any errors: C:\\users\\' + user + '\\documents\\autoSITS-' + timeStamp + '.log')
exit(0)