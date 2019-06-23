strUserName = "Sagar"



result = MsgBox ("Good Morning " & strUserName & ", Have you sent Morning Report", vbYesNo, "Morning Report Confirmation")

Select Case result
Case vbYes
    MsgBox("Good job "  & strUserName)
Case vbNo
    CreateMorningReportFile()
End Select

Sub CreateMorningReportFile()


    'WScript.Echo Day(Date)
    'WScript.Echo MonthName(Month(Date))
    'WScript.Echo Year(Date)

    strDay = Day(Date)
    strMonth = MonthName(Month(Date))
    strYear = Year(Date)

    strfullDate = strDay & " " & strMonth & " " & strYear
    'WScript.Echo strfullDate
    strFilePath = "H:\Small Demos\VBScript\"
    strFileName = "Sagar Dudhaiya Morning Report " & strfullDate & ".txt"

    'WScript.Echo "notepad.exe" & strFilePath & strFileName

    strFileNameWithPath = strFilePath & strFileName

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FileExists(strFileNameWithPath) Then
    Set objFile = objFSO.OpenTextFile(strFilePath & strFileName,2,true)
    objFile.WriteLine("KSEYE" & vbNewLine)
    objFile.WriteLine("test")
    objFile.Close
        OpenFile(strFileNameWithPath)
    Else

    Set objFile = objFSO.OpenTextFile(strFileNameWithPath,2,true)
    objFile.WriteLine("KSEYE" & vbNewLine)
    objFile.WriteLine("test")
    objFile.Close

    OpenFile(strFileNameWithPath)
    End If

End Sub

Sub OpenFile(FileName)

    With CreateObject("WScript.Shell")
        .Run "notepad.exe " & FileName
    End With

End Sub
