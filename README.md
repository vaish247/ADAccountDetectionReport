# ADAccountDetectionReport
Finds AD accounts, and generates an excel report

This article will outline how to use the PowerShell script that was written to help detect AD users or contacts. This script is called ADScan.ps1 and it's stored on the J Drive. The specific file location is here: J:\Scripts\ADDetection. You do not need an be an administrator to utilise the script.

Purpose of the script
The purpose of this script is to identify which people have AD contacts or accounts. So what it does is process an excel file, which has a list of names (the user has to provide the excel file. Then it edits the excel file and produces a report. The report identifies which people have AD contacts, and AD accounts, which users are active or disabled and it also provides a list of people who most likely don't have an AD account (They're called undetected users).
 
Warning
This script is not completely fool proof. If the program is unable to detect if an employee has an AD account, it does not mean that the employee does not have an AD account. This program works on the assumption that the user and AD have the correct employee's name. There are many cases where AD has the employee's name spelt wrong, or did not update the employee's name when they changed their name. So it is in the best interest of the user to investigate the undetected users.
  
Outline of the script
The most important piece of code in the script is the function scanData and the class ExcelObject.
The ExcelObject class is responsible for opening and writing into an excel file. The function scanData is responsible for processing every name provided by the user. When it processes each name, it will try and identify any AD Users or Contacts. After it has processed the name, it will call a function from the ExcelObject class to write the appropriate data into an Excel File.
 
Prerequisites
Before you can use the script, you need to prepare an excel file that has a list of users you want to try and find on AD. The excel file should look something like the image below.
![image](https://user-images.githubusercontent.com/87791446/221014977-716de2cf-14dc-4c80-8be5-45519c6d036b.png)



So there are 3 big things to note. Make sure the surname of the user is in the first column, and the first name is in the second column. The second thing to make sure is that the worksheets name is "Sheet1".  The other thing to note is that  make sure the first 2 cells are labels like "Surname" or "First Name". If you do not follow this format, the script will not work.

There is an example of how the excel file should be formatted in J:\Scripts\ADDetection.

Running the Script

So there are limitations with this script, this script can only run on the Windows PowerShell ISE application. So open up 
Windows PowerShell ISE, and in the PowerShell area change your current file location by typing and entering this line: cd J:\Scripts\ADDetection. Then run the program by typing and entering this: .\ADScan.ps1.

After running the script, a file explorer should pop up. This will allow you to select the excel file that you have prepared early on (The excel file that has a list of users that you want to find on AD). After you have chosen the appropriate excel file, the program will automatically start the scan.

When the scan is finished, the excel file that you have prepared should pop up. But now it should be filled with additional information. The excel file will now have 5 extra worksheets, that will display specific information. The name of the worksheet, should be enough to explain the information on the sheet.

Should look like this:
![image](https://user-images.githubusercontent.com/87791446/221015049-03ad4def-b88a-4208-a79b-a7a048458a2a.png)



Scripts Weakness:
The results of the script is a little erratic. So be careful.

Sometimes the script doesn't run properly, and most likely the problem is with the excel file, so your best bet is to make a copy of the excel file, and run the program using the copy. That generally solves most bugs. You could also sign out of your computer and sign back in and redo the scan process, which is another solution.

The bugs usually come in the form of random empty spaces which commonly occur in the "Details" worksheet.
![image](https://user-images.githubusercontent.com/87791446/221015332-264e628f-e2de-431a-a7da-31455838713d.png)

Another way to tell if the program screwed up is if the number of user errors is more than 0, it means that some names haven't been properly processed by the program,. You can find the Summary in the "Details" worksheet.
Note: Even though the program might not process a few names, it will not affect the accuracy of some of the other names.
![image](https://user-images.githubusercontent.com/87791446/221015122-7947af20-022c-49d6-9121-083709c838b5.png)

You can also check the log files to see if there's any bugs occurring . You can find the log files in the J Drive, or more specifically in this location:  J:\Scripts\ADDetection\Logs.  In the image below, it shows that when the program was scanning  Purna Thapa, an exception was thrown. If this happens, you can run the scan again, but what would most likely happen is that another user will experience the bug. The best solution is to sign out of your computer and sign back in and restart the scan.
![image](https://user-images.githubusercontent.com/87791446/221015502-b886e9f2-b81f-468f-96c5-4a7a2694b98f.png)

Naming standard of logs:    ADScanLogFile dd-MM-yyy hh-mm-ss.txt

dd -Days
MM- Months
yyy -Years
hh -Hours
mm -Minutes
ss -Seconds

Additional Info

There is a scanner that has a graphical user interface which is called ADScanGUI.ps1, however it will only work when you run the program in an Windows PowerShell ISE environment. Also, the interface doesn't provide any notable benefits, except running the scan multiple times in one session. It unfortunately does not  run on PowerShell
![image](https://user-images.githubusercontent.com/87791446/221014833-4f4b57b2-5102-497e-8dd2-36471a41b467.png)
