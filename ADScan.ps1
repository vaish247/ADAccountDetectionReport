#Written by Vaishnav

###########################################################################################################
#Description
#############################################################################################################

#So this program is supposed to detect if employees have AD accounts. So what the program does is take
#in an excel spreadsheet that have a list of employee names, then the program process that excel file and
#determine if the employees have any AD accounts. Then the programedit the excel file to displays it's results.


#If the program finds that an employee does have an AD account, it could be either listed as a disabled or an active account.
#If the program does not find any AD account for the user, then it will be labeled as "undetected".

#NOTE: If the program could not detect an AD account for an employee, it might be because their
#name on their AD accounts are wrong (this is quite common), so be carefull.


#######################################################################################
#Initializing
#########################################################################################

Add-Type -AssemblyName PresentationFramework #Used to support rendering the WPF Form 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") #used to help load the file explorer
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") #used to help load the file explorer


#######################################################################################
#Global Variables
#########################################################################################
$global:logFilePath= "J:\Scripts\ADDetection\Logs\ADScanLogFile "+ $(Get-date -format 'dd-MM-yyy hh-mm-ss')+".txt" #File path to the log file

######################################################################################
#Functions
######################################################################################


###Used to Write log messages inside the log file
Function logMessages(){

    #$logFilePath: file path of the log messages
    param($message,$logFilePath)
    try{
        Add-Content -Path $logFilePath -Value $message
        Write-Host "Message: '$message' Has been logged to file'"

    }catch{
        Write-Host "Ran into an issue: $PSItem"
    }
    

}

#Used to write the overarching result of the scan
Function summaryLog{

    param($excelObj)
    
    #Messages about how many types/list of users are there
    $activeMessageStr = "Number of active Users: " + $excelObj.userStats.activeCounter
    $activeMessageStr1 = "List of active Users: " + $excelObj.userStats.activeUsersList
    $disabeledMessageStr = "Number of disabled Users: " + $excelObj.userStats.disabledCounter
    $disabledMessageStr1 = "List of disabled Users: " + $excelObj.userStats.disabledUsersList
    $undetectedMessageStr = "Number of undetected Users: " + $excelObj.userStats.undetectedCounter
    $undetectedMessageStr1 = "List of undetected Users: " + $excelObj.userStats.undetectedUsersList
    $contactMessageStr = "Number of Users with Contacts: " + $excelObj.userStats.contactCounter
    $contactMessageStr1 = "List of Users with Contacts: " + $excelObj.userStats.contactCounterList 

    #Adding message to the log file, simplify later
    
    logMessages -message "-------------------------------------------------------------------------------------" -logFilePath $global:logFilePath
    logMessages -message "Summary Results" -logFilePath $global:logFilePath
    
    logMessages -message "" -logFilePath $global:logFilePath
    logMessages -message $activeMessageStr -logFilePath $global:logFilePath
    logMessages -message $activeMessageStr1 -logFilePath $global:logFilePath
    logMessages -message "" -logFilePath $global:logFilePath

    logMessages -message $disabeledMessageStr -logFilePath $global:logFilePath
    logMessages -message $disabledMessageStr1 -logFilePath $global:logFilePath
    logMessages -message "" -logFilePath $global:logFilePath

    logMessages -message $undetectedMessageStr -logFilePath $global:logFilePath
    logMessages -message $undetectedMessageStr1 -logFilePath $global:logFilePath
    logMessages -message "" -logFilePath $global:logFilePath

    logMessages -message $contactMessageStr -logFilePath $global:logFilePath
    logMessages -message $contactMessageStr1 -logFilePath $global:logFilePath
    logMessages -message "-------------------------------------------------------------------------------------" -logFilePath $global:logFilePath
}

### Used to open up the File explorer, so users can choose the file through the GUI
Function openFile
{
    #$directory: The 
    param([string] $directory)

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $directory
    $OpenFileDialog.filter = "All files (*.*)| *.xlsx*"
    $OpenFileDialog.ShowDialog() |  Out-Null  
    return $OpenFileDialog.filename

}


### Scanning Users on the excel sheet
Function scanData
{
    param($fileLocation, $excelObj)
    
    # Used to change the excel cell position
    $counter = 1

    #Checking if cell is empty
    $testingCell = $excelObj.ExcelWorkSheet.Range("B" + $counter).Text
    $cellCheck = [string]::IsNullOrEmpty($testingCell)
    
    # This loops untill the programs find an empty cell
    while($cellCheck -eq $false)
    {
        try{
        $counter += 1;
        ###Building the users full name
        $FName = $excelObj.ExcelWorkSheet.Range("B" + $counter).Text
        $lName = $excelObj.ExcelWorkSheet.Range("A" + $counter).Text
        $fullName = "$FName $lName"

        ###Changing how apostrophes are written in the name ( Apostrophe in a name causes string related issues when using the Get-ADUser function ,())
        if($fullName.IndexOf("'") -ne -1){
            $fullName =  $fullName.Replace("'","''")
        }

        $getUser = Get-ADUser -Filter "Name -like '$fullName'" -Properties displayname

        ###Testing through console
        Write-Host "$fullName"
        Write-Host "$getUser"

        $getContact = Get-ADObject -Filter "objectClass -eq 'contact' -and name -like '$fullName'" -Properties *

        #f
        if([string]::IsNullOrEmpty($getContact) -eq $false ){
            $excelObj.writeADData("ActiveC",$counter,$FName,$lName,$getUser.SamAccountName)
        }
        

        if($getUser -ne $null)
        {    
            
            #gets active users
            if($getUser.Enabled -eq $true )
            {
                $excelObj.writeADData("Active",$counter,$FName,$lName,$getUser.SamAccountName)
            #gets disabled users
            }else{ 
                $excelObj.writeADData("Disabled",$counter,$FName,$lName,$getUser.SamAccountName)
                
            }
        }
        #gets undetected users
        else{
            $excelObj.writeADData("Undetected",$counter,$FName,$lName,"N/A")
            
        }

        Write-Host "_______________________________________________________________________________________________"
       
        #Checking if the next cell down is empty
        $nextCounter = $counter+1
        $testingCell = $excelObj.ExcelWorkSheet.Range("B" + $nextCounter ).Text
        $cellCheck = [string]::IsNullOrEmpty($testingCell)
        }catch{
            $msgString = $PSItem.Exception.Message + " occuring with " + $fullName
            $excelObj.userStats.updateEUser()
            $excelObj.userStats.updateEUsersList()
            logMessages -message "" -logFilePath $global:logFilePath
            logMessages -message "ERRORS ARE HAPPENING!!!" -logFilePath $global:logFilePath
            logMessages -message $msgString -logFilePath $global:logFilePath
            logMessages -message "" -logFilePath $global:logFilePath
            $nextCounter = $counter+1
            $testingCell = $excelObj.ExcelWorkSheet.Range("B" + $nextCounter ).Text
            $cellCheck = [string]::IsNullOrEmpty($testingCell)

        }
    }
    $excelObj.summaryStats();
    $excelObj.autoFit() 
}


###########################################################################################################################
#Classes
###########################################################################################################################

# Used to collect data about the scan
Class InfoStatus{
    #Counter variables, used to count how many contact, disabled... users, should delete the counters later, kind of redundant 
    $contactCounter
    $activeCounter
    $disabledCounter 
    $undetectedCounter 
    $errorCounter

    #List variables, used to list the contact, disabled... users
    $activeUsersList
    $contactCounterList
    $disabledUsersList 
    $undetectedUsersList
    $errorList

    InfoStatus(){
        $this.contactCounter =0;
        $this.activeCounter =0;
        $this.disabledCounter =0;
        $this.undetectedCounter =0
        $this.errorCounter =0
        $this.contactCounterList = New-Object -TypeName 'System.Collections.ArrayList';
        $this.activeUsersList = New-Object -TypeName 'System.Collections.ArrayList';
        $this.disabledUsersList = New-Object -TypeName 'System.Collections.ArrayList';
        $this.undetectedUsersList = New-Object -TypeName 'System.Collections.ArrayList'
        $this.errorList = New-Object -TypeName 'System.Collections.ArrayList'
    }

    #retrieving stats
    [string] getErrorCounter(){
        return $this.errorCounter
    }

    [string] getUndetectedCounter(){
        return $this.undetectedCounter
    }

    [string] getDisabledCounter(){
        return $this.disabledCounter
    }

    [string] getContactCounter(){
        return $this.contactCounter
    }

    [string] getActiveCounter(){
        return $this.activeCounter
    }

    #updating stats
    [void] updateEUser(){
        $this.errorCounter += 1;
    }

    [void] updateDUser(){
        $this.disabledCounter += 1;
    }
    
    [void] updateUUser(){
        $this.undetectedCounter += 1;
    }
    [void] updateAUser(){
        $this.activeCounter += 1;
    }

    [void] updateCUser(){
        $this.contactCounter += 1;
    }
    [void] updateCUsersList($usersName){
        $this.contactCounterList.Add($usersName +",")
    }
    [void] updateAUsersList($usersName){
        $this.activeUsersList.Add($usersName +",")
    }

    [void] updateDUsersList($usersName){
        $this.disabledUsersList.Add($usersName +",")
    }

    [void] updateUUsersList($usersName){
        $this.undetectedUsersList.Add($usersName +",")
    }
    
    [void] updateEUsersList($usersName){
        $this.errorList.Add($usersName +",")
    }


}





###Used to deal with Excel related stuff such as opening an excel file, or writing into an excel file
Class ExcelObject{

    $ExcelObj #Used to open an excel object
    $ExcelWorkBook #Used to open an Excel workbook
    $ExcelWorkSheet #used to open the desired sheet

    #A bunch of new sheets that will be created 
    $worksheetDetails 
    $worksheetActive
    $worksheetDisable
    $worksheetUndetect
    $worksheetContact

    #An object thats going to collect information, while info is being written inside the excel sheet.
    $userStats;

    #initialize the excel objects (basically opens up the excel files and create multiple new sheets to work on)
    ExcelObject($chosenFile){
        $this.ExcelObj = New-Object -comobject Excel.Application
        $this.ExcelObj.visible = $true
        $this.ExcelWorkBook = $this.ExcelObj.Workbooks.Open($chosenFile)
        $this.ExcelWorkSheet = $this.ExcelWorkBook.Sheets.Item("Sheet1")

        #initializes sheets
        $this.createNewSheets()

        #setting up the new Sheets
        $this.initializingSheets($this.worksheetActive,"Active Users",$false)
        $this.initializingSheets($this.worksheetDisable,"Disabled Users",$false)
        $this.initializingSheets($this.worksheetUndetect,"Undetected Users",$false)
        $this.initializingSheets($this.worksheetContact,"Users with AD Contacts",$true)
        $this.initializingSheets($this.worksheetDetails,"Details",$true)

        $this.userStats = [InfoStatus]::new()

    }


    #initializing sheets
    [void] createNewSheets(){

        $this.worksheetDetails = $this.ExcelWorkBook.Worksheets.Add()
        $this.worksheetActive = $this.ExcelWorkBook.Worksheets.Add()
        $this.worksheetDisable = $this.ExcelWorkBook.Worksheets.Add()
        $this.worksheetUndetect = $this.ExcelWorkBook.Worksheets.Add()
        $this.worksheetContact = $this.ExcelWorkBook.Worksheets.Add()

    }

    #Writing data into all the excel worksheets
    [void]writeADData($status,$counter,$FName, $lName,$username){

        $this.worksheetDetails.cells.item($counter,1) = $FName
        $this.worksheetDetails.cells.item($counter,2) = $lName
        $this.worksheetDetails.cells.item($counter,3) = $username
        $fullName = $FName +" " +$lName

        if($username -ne "N/A"){
            $this.worksheetDetails.cells.item($counter,3).Interior.ColorIndex = 43
        }

        #writing data into specified sheets
        switch ($status)
        {
            "Active" {
                "Active Account"
                $this.worksheetDetails.cells.item($counter,4) = "Active"
                $this.worksheetDetails.cells.item($counter,4).Interior.ColorIndex  = 4
                $this.userStats.updateAUser()
                $this.userStats.updateAUsersList($fullName)

                $rowNum = [int]$this.userStats.getActiveCounter()+1
                $this.writingSortedInfo($this.worksheetActive,$rowNum,$FName,$lName,$username,"Active",4,$false)
                
            }
            "Disabled" {
                "Disabled Account"
                $this.worksheetDetails.cells.item($counter,4) = "Disabled"
                $this.worksheetDetails.cells.item($counter,4).Interior.Color = 255
                $this.userStats.updateDUser()
                $this.userStats.updateDUsersList($fullName)

                $rowNum = [int]$this.userStats.getDisabledCounter()+1
                $this.writingSortedInfo($this.worksheetDisable,$rowNum,$FName,$lName,$username,"Disabled",3,$false)

            }
            "Undetected" {
                "Undetected Account"
                $this.worksheetDetails.cells.item($counter,4) = "Undetected"
                $this.worksheetDetails.cells.item($counter,4).Interior.ColorIndex  = 27
                $this.userStats.updateUUser()
                $this.userStats.updateUUsersList($fullName)

                $rowNum = [int]$this.userStats.getUndetectedCounter()+1
                $this.writingSortedInfo($this.worksheetUndetect,$rowNum,$FName,$lName,"N/A","Undetected",27,$false)

            }
            "ActiveC" {
                "Active Contact"
                $this.worksheetDetails.cells.item($counter,5) = "Active Contact"
                $this.worksheetDetails.cells.item($counter,5).Interior.ColorIndex  = 10
                $this.userStats.updateCUser()
                $this.userStats.updateCUsersList($fullName)

                $rowNum = [int]$this.userStats.getContactCounter()+1
                $this.writingSortedInfo($this.worksheetContact,$rowNum,$FName,$lName,$username,"Info not here",10,$true)
            }
        }  
    }


    
    #writing the AD information into the excel for a single worksheets
    [void]writingSortedInfo($sheet,$rowNum,$FName,$lName,$username,$info,$colour,$contactStatus ){
        $sheet.cells.item($rowNum,1) = $FName
        $sheet.cells.item($rowNum,2) = $lName
        $sheet.cells.item($rowNum,3) = $username
        $sheet.cells.item($rowNum,4) = $info
        $sheet.cells.item($rowNum,4).Interior.ColorIndex  = $colour
        if($contactStatus){
            $sheet.cells.item($rowNum,5)= "Active Contact"
            $sheet.cells.item($rowNum,5).Interior.ColorIndex  = $colour
            $sheet.cells.item($rowNum,4).Interior.ColorIndex  = 2
        }


    }

    #setting up new sheets
    [void] initializingSheets($sheet,$name,$contactStatus){
        $sheet.name =$name
        $sheet.cells.item(1,1) = 'First Name'
        $sheet.cells.item(1,1).Font.Bold = $True

        $sheet.cells.item(1,2) = 'Last Name'
        $sheet.cells.item(1,2).Font.Bold = $True

        $sheet.cells.item(1,3) = 'Username'
        $sheet.cells.item(1,3).Font.Bold = $True

        $sheet.cells.item(1,4) = 'AD Account Status'
        $sheet.cells.item(1,4).Font.Bold = $True

        if($contactStatus){
            $sheet.cells.item(1,5) = 'AD Contact Status'
            $sheet.cells.item(1,5).Font.Bold = $True
        }

    }

    #Used to auto fit all the text
    [void]autoFit(){
        $this.worksheetDetails.columns.item("A:G").EntireColumn.AutoFit() | out-null
        $this.worksheetActive.columns.item("A:E").EntireColumn.AutoFit() | out-null
        $this.worksheetDisable.columns.item("A:E").EntireColumn.AutoFit() | out-null
        $this.worksheetUndetect.columns.item("A:E").EntireColumn.AutoFit() | out-null
        $this.worksheetContact.columns.item("A:E").EntireColumn.AutoFit() | out-null
    }


    #used to write a summary result inside the excel sheet
    [void]summaryStats(){
        $this.worksheetDetails.cells.item(4,6) = "Summary"
        $this.worksheetDetails.cells.item(4,6).Font.Bold = $True

        $this.worksheetDetails.cells.item(5,7) = $this.userStats.activeCounter
        $this.worksheetDetails.cells.item(5,6) = "Number of Active Users"

        $this.worksheetDetails.cells.item(6,7) = $this.userStats.disabledCounter
        $this.worksheetDetails.cells.item(6,6) = "Number of Disabled Users"

        $this.worksheetDetails.cells.item(7,7) = $this.userStats.undetectedCounter
        $this.worksheetDetails.cells.item(7,6) = "Number of undetected Users"

        $this.worksheetDetails.cells.item(8,7) = $this.userStats.contactCounter
        $this.worksheetDetails.cells.item(8,6) = "Number of Users with contacts"

        $this.worksheetDetails.cells.item(9,7) = $this.userStats.errorCounter
        $this.worksheetDetails.cells.item(9,6) = "Number of User Errors"
    }
}


###########################################################################################################################
#Execution
###########################################################################################################################




#Creating log file and the required folder(if needed)
New-Item $global:logFilePath #creating the log file
logMessages -message "Starting the script" -logFilePath $global:logFilePath

#Starting the program
logMessages -message "Opening file explorer" -logFilePath $global:logFilePath
$file = openFile -directory $env:USERPROFILE

if ($file -ne "") 
{
    Write-Host "You choose FileName: $file" 
    $mString = "Chosen File: " + $file
    logMessages -message $mString -logFilePath $global:logFilePath
    $excelObj = [ExcelObject]::new($file)
    logMessages -message "Running the scan" -logFilePath $global:logFilePath
    scanData -fileLocation $file -excelObj $excelObj
    summaryLog -excelObj $excelObj
    Write-Host "Program is closing now"
    logMessages -message "Program is closing now" -logFilePath $global:logFilePath
    
} 
else 
{
    Write-Host "No File was chosen, program is closing now!"
    logMessages -message "No File was chosen, exiting program" -logFilePath $global:logFilePath
}

