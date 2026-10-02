Option Explicit

Dim sh, exePath, i, command, process, output

If WScript.Arguments.Count < 1 Then
    WScript.Quit 1
End If

Set sh = CreateObject("WScript.Shell")

exePath = WScript.Arguments(0)
command = """" & exePath & """"
For i = 1 To WScript.Arguments.Count - 1
    command = command & " """ & WScript.Arguments(i) & """"
Next

Set process = sh.Exec(command)
output = ""
Do While Not process.StdOut.AtEndOfStream
    output = output & process.StdOut.ReadAll()
Loop

WScript.StdOut.Write Trim(output)