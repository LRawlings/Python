This script is disigned to create multiple test student accounts in SITS:Vision. The main use is for creating applicant accounts in SITS Test for testing the Online Registration portal.

The program I have used to create and run these skripts is SikuiX.

IMPORTANT
Only use this automation tool to make change to TEST or DEVELOPMENT environments. Using automation such as this on the live system is too dangerous as human error in the code could make irreversible unintended changes.

SikuliX lets you script Windows desktop functions such tabbing through menus and typing into search boxes. It can also locate areas on your screen based on a saved image and click. For example, you can set your script to find the Generate button on the SITS QAS screen and click it: 

SikuliX is a Java based application and the code can be written in various scripting languages but we have chosen to use Python level 2.7 (supported by Jython).

SikuliX includes it's own IDE which provides basic support for for editing and running scripts. This is the option we are using however it is possible to to integrate SikuliX with other more mature IDE's.

To run the pre-packaged script (.jar file), you will need to install the latest version of Java.

To edit and see more advanced logs for debugging, you will need to install the SikuliX IDE. Details of how to install can be found here: http://sikulix.com/quickstart/.

What the Script Does:

    Create a log file in the currently logged on users Documents folder named autoSITS-time.log where time is the current time the log is started
    Open an excel document stored in the current logged on users documents folder
    Open SITS Dev or Test based on user input and prompt to log on
    Wait for the SITS login to complete and the SITS Menu System to load
    Change focus back to the reference excel document
    Store the 3rd row of data (the first 2 rows are header rows) into a list. This script is configured to store 13 strings of data into this list but this can be changed if more details are needed.
    Change focus to SITS
    Open the QAS screen if not already open
    Enter the stored data into the relevant fields and click Generate
    A log entry will be made based on the outcome of the QAS generation (success/failure)
    Close QAS
    If QAS was successful open CAPS
    Enter data into CAPS and store
    A log entry will be made based on the outcome of the CAPS update
    Close CAPS
    If CAPS update was successful open ATR
    Enter ATR details and click Run
    A log entry will be made based on the outcome of the ATR run
    Close ATR
    If ATR run was successful open SCE
    Enter details in SCE and store
    A log entry will be made based on the outcome of the SCE update
    Close SCE
    Based on the number of rows to be processed (set by the user at the being of the script) steps 5 to 10 will be repeated.
    A pop-up will confirm the script has completed successfully and confirms the location and name of the log file.
    User can then check the log to see which accounts created successfully and which failed (and why).