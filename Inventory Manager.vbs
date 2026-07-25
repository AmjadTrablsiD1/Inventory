' Inventory Manager - Windows launcher
'
' Double-click this file to start the app WITHOUT a console window.
' Keep it in the same folder as app.py.
'
' It finds Python, installs Flask on first run if needed, starts the server
' hidden, and app.py then opens your browser. If the app is already running it
' just reopens the browser instead of starting a second copy.

Option Explicit

Const PORT = "8765"

Dim fso, sh, projectDir, appPath, running, http
Dim consolePy, windowPy

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")

projectDir = fso.GetParentFolderName(WScript.ScriptFullName)
appPath = fso.BuildPath(projectDir, "app.py")
sh.CurrentDirectory = projectDir

If Not fso.FileExists(appPath) Then
    MsgBox "app.py was not found next to this launcher." & vbCrLf & vbCrLf & _
           "Keep ""Inventory Manager.vbs"" in the same folder as app.py.", _
           vbCritical, "Inventory Manager"
    WScript.Quit 1
End If

' ---- already running? just reopen the browser ----
running = False
On Error Resume Next
Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
If Err.Number = 0 Then
    http.setTimeouts 1000, 1000, 1000, 2000
    http.Open "GET", "http://127.0.0.1:" & PORT & "/api/state", False
    http.Send
    If Err.Number = 0 Then
        If http.Status = 200 Then running = True
    End If
End If
Err.Clear
On Error GoTo 0

If running Then
    sh.Run "http://127.0.0.1:" & PORT, 1, False
    WScript.Quit 0
End If

' ---- locate Python 3 ----
consolePy = ""
If TryRun(sh, "python -c ""pass""") = 0 Then
    consolePy = "python"
    windowPy = "pythonw"
ElseIf TryRun(sh, "py -3 -c ""pass""") = 0 Then
    consolePy = "py -3"
    windowPy = "pyw -3"
End If

If consolePy = "" Then
    MsgBox "Python 3 was not found." & vbCrLf & vbCrLf & _
           "Install it from https://www.python.org/downloads/ and tick" & vbCrLf & _
           """Add Python to PATH"" during setup, then try again.", _
           vbCritical, "Inventory Manager"
    WScript.Quit 1
End If

' ---- make sure Flask is installed ----
If TryRun(sh, consolePy & " -c ""import flask""") <> 0 Then
    MsgBox "Flask is missing and will now be installed." & vbCrLf & _
           "This happens only once and may take a minute." & vbCrLf & vbCrLf & _
           "Click OK and wait - the app opens by itself when it is ready.", _
           vbInformation, "Inventory Manager"
    If TryRun(sh, consolePy & " -m pip install --user flask") <> 0 Then
        If TryRun(sh, consolePy & " -m pip install flask") <> 0 Then
            MsgBox "Could not install Flask automatically." & vbCrLf & vbCrLf & _
                   "Open Command Prompt and run:" & vbCrLf & vbCrLf & _
                   "    python -m pip install flask", _
                   vbCritical, "Inventory Manager"
            WScript.Quit 1
        End If
    End If
End If

' ---- start it hidden (app.py opens the browser itself) ----
If TryRun(sh, windowPy & " -c ""pass""") <> 0 Then
    windowPy = consolePy          ' no pythonw available - run console python hidden
End If

sh.Run windowPy & " """ & appPath & """", 0, False
WScript.Quit 0


' Runs a command hidden, waits, and returns its exit code (-1 if it could not run).
Function TryRun(shell, command)
    Dim code
    code = -1
    On Error Resume Next
    code = shell.Run(command, 0, True)
    If Err.Number <> 0 Then code = -1
    Err.Clear
    On Error GoTo 0
    TryRun = code
End Function
