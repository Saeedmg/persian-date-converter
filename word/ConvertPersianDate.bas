' -------------------------------------------------------------------
'  DEFAULT - uses whatever format is set in the GUI (settings.ini)
' -------------------------------------------------------------------
Sub ConvertPersianDate()
    ConvertDateWithFormat ""
End Sub


' -------------------------------------------------------------------
'  PRIORITY FORMATS
' -------------------------------------------------------------------

' 1) American full: September 17, 2024
Sub ConvertPersianDate_Full()
    ConvertDateWithFormat "5"
End Sub

' 2) American abbreviated: Sep. 17, 2024
Sub ConvertPersianDate_Abbrev()
    ConvertDateWithFormat "7"
End Sub

' 3) British full: 17 September 2024
Sub ConvertPersianDate_British()
    ConvertDateWithFormat "1"
End Sub

' 4) British abbreviated: 17 Sep. 2024
Sub ConvertPersianDate_BritishAbbrev()
    ConvertDateWithFormat "6"
End Sub

' 5) ISO dashes: 2024-09-17
Sub ConvertPersianDate_ISO()
    ConvertDateWithFormat "4"
End Sub

' 6) ISO dots: 2024.09.17
Sub ConvertPersianDate_DotISO()
    ConvertDateWithFormat "21"
End Sub

' 7) American slashes: 09/17/2024
Sub ConvertPersianDate_SlashUS()
    ConvertDateWithFormat "2"
End Sub


' ===================================================================
'  SHARED ENGINE
' ===================================================================
Private Sub ConvertDateWithFormat(fmtKey As String)
    Dim dateInput As String
    Dim exePath As String
    Dim vbsPath As String
    Dim tmpFile As String
    Dim Q As String
    Dim cmdLine As String
    Dim convertedDate As String
    Dim wsh As Object
    Dim fso As Object
    Dim stream As Object

    Q = Chr(34)

    ' exePath points to the FAST build (inside the subfolder)
    exePath = "C:\PersianDateTool\dist\PersianDateConverterCLI\PersianDateConverterCLI.exe"
    ' vbsPath stays in dist\ (top level)
    vbsPath = "C:\PersianDateTool\dist\RunHidden.vbs"
    tmpFile = Environ("TEMP") & "\pdc_word_out.txt"

    dateInput = InputBox("Enter the Persian date (YYYY/MM/DD):", _
                         "Persian Date Conversion")
    If dateInput = "" Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(tmpFile) Then fso.DeleteFile tmpFile, True

    If fmtKey = "" Then
        ' No format digit ? CLI reads settings.ini (the GUI's choice)
        cmdLine = "cmd.exe /c " & Q & _
                  "wscript.exe " & Q & vbsPath & Q & " " & _
                  Q & exePath & Q & " " & _
                  Q & dateInput & Q & _
                  " > " & Q & tmpFile & Q & " 2>&1" & Q
    Else
        ' Explicit format digit ? CLI uses this format
        cmdLine = "cmd.exe /c " & Q & _
                  "wscript.exe " & Q & vbsPath & Q & " " & _
                  Q & exePath & Q & " " & _
                  Q & dateInput & Q & " " & Q & fmtKey & Q & _
                  " > " & Q & tmpFile & Q & " 2>&1" & Q
    End If

    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmdLine, 0, True

    ' Read the temp file as UTF-8 (Arabic, French, Turkish, ...)
    If fso.FileExists(tmpFile) Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2
        stream.Charset = "utf-8"
        stream.Open
        stream.LoadFromFile tmpFile
        convertedDate = Trim(stream.ReadText(-1))
        stream.Close
        Set stream = Nothing
        fso.DeleteFile tmpFile, True
    Else
        convertedDate = ""
    End If

    convertedDate = Replace(convertedDate, vbCrLf, "")
    convertedDate = Replace(convertedDate, vbLf, "")
    convertedDate = Replace(convertedDate, vbCr, "")

    If convertedDate = "" Or InStr(convertedDate, "Error:") > 0 Then
        MsgBox "Error: Please enter a valid Persian date in the format YYYY/MM/DD."
    Else
        Selection.TypeText convertedDate
    End If
End Sub


