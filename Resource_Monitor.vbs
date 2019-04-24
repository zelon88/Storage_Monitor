'File Name: Storage_Monitor.vbs
'Version: v1.9, 4/23/2019, Fix bugs with string comparison.
'Author: Justin Grimes, 5/31/2018

Option Explicit
Dim inputCache, outputCache, objShell, Result, DiskSet, Disk, oFSO, mailFile, oCacheHandle, iCacheHandle, mFileHandle, Device, strComputerName, outCacheData, inCacheData, inCacheString, _
outCacheString, strLogFilePath, strSafeDate, strSafeTime, strDateTime, strLogFileName, homeFolder, objLogFile, Alert, pre, fireEmail, mailHandle, outCacheNew, _
toEmail, fromEmail, companyAbbreviation, companyName, strDiff, re, installPath

'Define variables & basic objects for the session.
fireEmail = False
Alert = pre = Device = ""
Set objShell = Wscript.CreateObject("WScript.Shell")
Set re = New RegExp
re.Pattern = "\s+"
re.Global  = True
homeFolder = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
'The following variables are required to create a logfile in the network Logs directory.
strComputerName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
strSafeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
strDateTime = strSafeDate & "-" & strSafeTime
strLogFileName = strLogFilePath & "\" & strComputerName & "-" & strDateTime & "-storage_monitor.txt"
'----------
'The variables within this comment block should be adjusted to your environment.
installPath = "\\Server\AutomationScripts\Storage_Monitor"
mailFile = homeFolder & "\Storage_Monitor_Warning.mail"
inputCache = homeFolder & "\diskCache0.dat"
outputCache = homeFolder & "\diskCache1.dat"
strComputerName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
strLogFilePath = "\\Server\Logs"
toEmail = "IT@company.com"
fromEmail = "Server@company.com"
companyAbbreviation = "Company"
companyName = "Company Inc."
'----------

'Set some handles for disk objects (from WMI) and file system objects.
Set DiskSet = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery ("select * from Win32_LogicalDisk")
Set oFSO = CreateObject("Scripting.FileSystemObject")

'Verify that an output cache exists and create one if it does not.
Set oCacheHandle = oFSO.CreateTextFile(outputCache, True, False)
oCacheHandle.Close

'Verify that an input cache exists and create one if it does not.
'Also sets a handle for writing to the input cache.
If Not (oFSO.FileExists(inputCache)) Then
  Set iCacheHandle = oFSO.CreateTextFile(inputCache, True, False)
End If

'A function for running SendMail.
Function SendEmail() 
 objShell.run installPath & "\sendmail.exe " & mailFile 
End Function

'A function to create a log file.
Function CreateLog(strEventInfo)
  strSafeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
  strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
  strDateTime = strSafeDate & "-" & strSafeTime
  strLogFileName = strLogFilePath & "\" & strComputerName & "-" & strDateTime & "-storage_monitor.txt"
  If Not (strEventInfo = "") Then
    Set objLogFile = oFSO.CreateTextFile(strLogFileName, False, False)
    objLogFile.WriteLine(strEventInfo)
    objLogFile.Close
  End If
End Function

'Check each disk for available space.
For Each Disk In DiskSet

  'Retrieve the drive letter of each device.
  If (Device <> "") Then
    Device = Device & "," & Disk.Name
  Else
    Device = Disk.Name
  End If

  'Retrieve the amount of free space on the disk.
  Disk.FreeSpace = Disk.FreeSpace/1024
  Disk.FreeSpace = Disk.FreeSpace/1024
  Result = Disk.FreeSpace/1024

  'Prepare delimiters for the list of devices that are low on storage.
  If (Alert = "") Then
    pre = ""
  End If
  If (Alert <> "") Then
    pre = ","
  End If

  'Set the threshold for amount of disk space remaining before a warning email is sent.
  If (Result <= 15) Then
    Alert = Alert & pre & Disk.Name
  End If
Next

'Rewrite the output cache.
Set oCacheHandle = oFSO.CreateTextFile(outputCache, True, False)
oCacheHandle.WriteLine(Device)
oCacheHandle.Close

'Retrieve the contents of the input cache file.
Set inCacheData = oFSO.OpenTextFile(inputCache, 1)
If Not inCacheData.AtEndOfStream Then
  inCacheString = inCacheData.ReadAll
Else
  inCacheString = ""
End If
inCacheData.Close

'Compare the contents of the two cache files.
Device = Trim(re.Replace(Device, ""))
inCacheString = Trim(re.Replace(inCacheString, ""))
strDiff = StrComp(Device, inCacheString, vbTextCompare)
If (strDiff <> 0) Then
  fireEmail = False
End If

'Regenerate the input cache file with data from the output cache file.
'Retrieve the contents of the input cache file.
Set outCacheData = oFSO.OpenTextFile(outputCache, 1)
outCacheNew = outCacheData.ReadAll
outCacheData.Close

Set inCacheData = oFSO.CreateTextFile(inputCache, True, False)
inCacheData.Write outCacheNew
inCacheData.Close

'Send one email if a storage device is low on space (after all loops have completed).
If (len(Alert) >= 1) Then
  Set mFileHandle = oFSO.CreateTextFile(mailFile, True, False)
  mFileHandle.Write "To: "&toEmail&vbNewLine&"From: "&fromEmail&vbNewLine&"Subject: "&companyAbbreviation&" Low Storage Space Warning!!!"&vbNewLine& _
   "This is an automatic email from the "&companyName&" Network to notify you that a storage device is almost full and requires attention."&vbNewLine&vbNewLine& _
   "Please log-in and verify that the equipment listed below has adequate storage space."&vbNewLine&vbNewLine&"IMPACTED DEVICE: "&strComputerName&vbNewLine&"DRIVES: "&Alert& _
   vbNewLine&vbNewLine&"This check was generated by "&strComputerName&" and is performed every 30 minutes."&vbNewLine&vbNewLine&"Script: ""Storage_Monitor.vbs""" 
  mFileHandle.Close
  SendEmail
  CreateLog("The storage devices of " & strComputerName & " are almost full on " & strDateTime & "!" & vbNewLine & vbNewLine & "DRIVES: " & Alert)
  WScript.Sleep 1000
End If

'Send one email if storage configuration has changed (after all loops have completed).
If (fireEmail = True) Then
  Set mFileHandle = oFSO.CreateTextFile(mailFile, True, False)
  mFileHandle.Write "To: "&toEmail&vbNewLine&"From: "&fromEmail&vbNewLine&"Subject: "&companyAbbreviation&" Storage Device Change Warning!!!"&vbNewLine& _
   "This is an automatic email from the "&companyName&" Network to notify you that a storage device configuration has changed and requires attention."&vbNewLine&vbNewLine& _
   "Please log-in and verify that the equipment listed below has it's storage devices configured correctly."&vbNewLine&vbNewLine&"IMPACTED DEVICE: "&strComputerName&vbNewLine&"DRIVES: "&Device& _
   vbNewLine&vbNewLine&"This check was generated by "&strComputerName&" and is performed every 30 minutes."&vbNewLine&vbNewLine&"Script: ""Storage_Monitor.vbs""" 
  mFileHandle.Close
  SendEmail
  CreateLog("The storage configuration on " & strComputerName & " has changed on " & strDateTime & "!" & vbNewLine & vbNewLine & "DRIVES: " & Device)
End If