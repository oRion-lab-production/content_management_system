' ============================================================
'  RunPocTests.vbs
'  End-to-end smoke test harness for the WCMS Migration API.
'  Tests the JSON API and the Excel/CSV API (plus Swagger + storage).
'
'  Run from a console:   cscript //nologo RunPocTests.vbs
'  Or just double-click it: it re-launches itself under cscript so
'  output goes to a console window instead of Windows Script Host popups.
' ============================================================
Option Explicit

' --- Force a real console even when launched by wscript (double-click) ---
' Re-runs itself inside cmd.exe /k so the window stays open and the user
' can watch live progress and read the results afterwards.
If InStr(1, LCase(WScript.FullName), "wscript.exe") > 0 Then
    Dim relauncher
    Set relauncher = CreateObject("WScript.Shell")
    relauncher.Run "cmd.exe /k cscript.exe //nologo """ & WScript.ScriptFullName & """", 1, False
    WScript.Quit 0
End If

Const BASE_URL   = "http://localhost:5233"

Dim fso, shell, repo, proj, dataDir, apiLog
Dim passCount, failCount, runId, startedHere, READY_URL
Dim XLSX_B64, XLS_B64

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
repo = fso.GetParentFolderName(WScript.ScriptFullName)
proj = repo & "\WebAPIs\WebAPIs.csproj"
dataDir = repo & "\testdata"
' Transient location for the API console output (kept out of the repo, deleted at the end).
apiLog = fso.GetSpecialFolder(2) & "\wcms_poc_api.log"
READY_URL = BASE_URL & "/swagger/v1/swagger.json"
runId = "R" & CStr(CLng(Timer * 1000))

' ---- Embedded binary test fixtures (base64) ----
XLSX_B64 = "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"
XLS_B64  = "0M8R4KGxGuEAAAAAAAAAAAAAAAAAAAAAOwADAP7/CQAGAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAEAAAAgAAAAEAAAD+////AAAAAAEAAAD///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9SAG8AbwB0ACAARQBuAHQAcgB5AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFgAFAf//////////AQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMAAACACQAAAAAAAFcAbwByAGsAYgBvAG8AawAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAASAAIB////////////////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHUJAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP7////9/////v///wQAAAAFAAAABgAAAAcAAAD+////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////AQAAAAIAAAADAAAABAAAAAUAAAAGAAAABwAAAAgAAAAJAAAACgAAAAsAAAAMAAAADQAAAA4AAAAPAAAAEAAAABEAAAASAAAAEwAAABQAAAAVAAAAFgAAABcAAAAYAAAAGQAAABoAAAAbAAAAHAAAAB0AAAAeAAAAHwAAACAAAAAhAAAAIgAAACMAAAAkAAAAJQAAAP7///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8JCBAAAAYFANMQzAdBAAAABgAAAOEAAgCwBMEAAgAAAOIAAABcAHAACAAAQU41MTVfNTIgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIEIAAgCwBGEBAgAAAD0BAgAAAJwAAgAOABkAAgAAABIAAgAAABMAAgAAAK8BAgAAALwBAgAAAD0AEgBoAQ4BXDq+IzgAAAAAAAEAWAJAAAIAAACNAAIAAAAiAAIAAAAOAAIAAQC3AQIAAADaAAIAAAAxABUAyAAAAP9/kAEAAAAAAAAFAEFyaWFsMQAVAMgAAAD/f5ABAAAAAAAABQBBcmlhbDEAFQDIAAAA/3+QAQAAAAAAAAUAQXJpYWwxABUAyAAAAP9/kAEAAAAAAAAFAEFyaWFsHgQaAAUAFQAAIiQiIywjIzBfKTsoIiQiIywjIzApHgQfAAYAGgAAIiQiIywjIzBfKTtbUmVkXSgiJCIjLCMjMCkeBCAABwAbAAAiJCIjLCMjMC4wMF8pOygiJCIjLCMjMC4wMCkeBCUACAAgAAAiJCIjLCMjMC4wMF8pO1tSZWRdKCIkIiMsIyMwLjAwKR4ENQAqADAAAF8oIiQiKiAjLCMjMF8pO18oIiQiKiAoIywjIzApO18oIiQiKiAiLSJfKTtfKEBfKR4ELAApACcAAF8oKiAjLCMjMF8pO18oKiAoIywjIzApO18oKiAiLSJfKTtfKEBfKR4EPQAsADgAAF8oIiQiKiAjLCMjMC4wMF8pO18oIiQiKiAoIywjIzAu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"

passCount = 0
failCount = 0

WScript.Echo "=================================================="
WScript.Echo " WCMS Migration API - POC test harness"
WScript.Echo " Repo : " & repo
WScript.Echo " Run  : " & runId
WScript.Echo "=================================================="

' ---- 1. Show where everything lives -------------------------
ShowLocations

' ---- 2. Ensure toolchain + build the project ----------------
If Not EnsureDotNet() Then WScript.Quit 1
If Not BuildApi() Then
    WScript.Echo "[FATAL] Build failed; aborting."
    WScript.Quit 1
End If

' ---- 3. Prepare test data -----------------------------------
PrepareTestData

' ---- 4. Ensure the API is running ---------------------------
startedHere = False
If Not ServerReady(2) Then
    WScript.Echo "[..] Starting API (dotnet run)..."
    startedHere = True
    If Not StartApi() Then
        WScript.Echo "[FATAL] Could not start the API."
        WScript.Quit 1
    End If
    If Not WaitForServer(120) Then
        WScript.Echo "[FATAL] API did not become ready in time."
        StopApi
        WScript.Quit 1
    End If
    WScript.Echo "[OK] API started."
Else
    WScript.Echo "[OK] API already running."
End If

WScript.Echo ""

' ---- 3. Test scenarios --------------------------------------
WScript.Echo "-- Swagger / hosting --"
Check "GET  /swagger/v1/swagger.json", GetStatus(READY_URL), 200
Check "GET  /swagger/index.html", GetStatus(BASE_URL & "/swagger/index.html"), 200
Check "GET  / (landing page)", GetStatus(BASE_URL & "/"), 200

WScript.Echo ""
WScript.Echo "-- JSON API  POST /api/v1/Migration/PushContent --"
Dim jsonKey, validJson
jsonKey   = runId & "-json"
validJson = "{""sourceSystem"":""POC-" & runId & """,""items"":[" & _
            "{""sku"":""A1"",""name"":""Alpha"",""qty"":""5"",""colour"":""red""}," & _
            "{""sku"":""A2"",""name"":""Beta"",""qty"":""9"",""colour"":""blue""}]}"
Check "valid payload (2 items) -> accepted",     JsonStatus(validJson, jsonKey), 204
Check "same payload again    -> duplicate",      JsonStatus(validJson, jsonKey), 200
Check "empty items           -> bad request",    JsonStatus("{""sourceSystem"":""POC"",""items"":[]}", runId & "-empty"), 400
Check "501 items             -> threshold",      JsonStatus(BuildJsonItems(501), runId & "-bulk"), 400

WScript.Echo ""
WScript.Echo "-- Excel/CSV API  POST /api/v1/Migration/UploadExcel --"
Check "upload sample.xlsx    -> accepted",       UploadStatus(dataDir & "\sample.xlsx", runId & "-xlsx"), 202
Check "same xlsx again       -> duplicate",      UploadStatus(dataDir & "\sample.xlsx", runId & "-xlsx"), 200
Check "upload sample.xls     -> accepted",       UploadStatus(dataDir & "\sample.xls",  runId & "-xls"), 202
Check "upload sample.csv     -> accepted",       UploadStatus(dataDir & "\sample.csv",  runId & "-csv"), 202
Check "upload empty file     -> bad request",    UploadStatus(dataDir & "\empty.bin",   runId & "-empty"), 400

WScript.Echo ""
WScript.Echo "-- Storage (WebAPIs\storage) --"
Dim nosqlCount, fileCount
nosqlCount = CountFiles(repo & "\WebAPIs\storage\nosql")
fileCount  = CountFiles(repo & "\WebAPIs\storage\file")
Check "storage\nosql has documents", IIF(nosqlCount > 0, "yes", "no"), "yes"
Check "storage\file  has uploads",   IIF(fileCount  > 0, "yes", "no"), "yes"
WScript.Echo "     nosql files: " & nosqlCount & " | file uploads: " & fileCount

' ---- 4. Summary ---------------------------------------------
WScript.Echo ""
WScript.Echo "=================================================="
WScript.Echo " RESULT : " & passCount & " passed, " & failCount & " failed"
WScript.Echo "=================================================="

If startedHere Then
    WScript.Echo "[..] Stopping API..."
    StopApi
End If

CleanupRun

If failCount > 0 Then WScript.Quit 1 Else WScript.Quit 0

' ============================================================
'  Helpers
' ============================================================

Sub PrepareTestData
    If Not fso.FolderExists(dataDir) Then fso.CreateFolder(dataDir)

    WriteBytes dataDir & "\sample.xlsx", B64ToBytes(XLSX_B64)
    WriteBytes dataDir & "\sample.xls",  B64ToBytes(XLS_B64)

    Dim csv
    csv = "ProductName,Brand,Price,InStock,Notes" & vbCrLf & _
          "Widget A,Acme,9.99,yes,""A fine, sturdy widget""" & vbCrLf & _
          "Gadget X,Initech,0,yes,sale" & vbCrLf
    WriteText dataDir & "\sample.csv", csv

    Dim ts
    Set ts = fso.CreateTextFile(dataDir & "\empty.bin", True)
    ts.Close
End Sub

Sub ShowLocations
    WScript.Echo "Paths (resolved relative to this script):"
    WScript.Echo "  Script    : " & WScript.ScriptFullName
    WScript.Echo "  Repo root : " & repo
    WScript.Echo "  Project   : " & proj
    WScript.Echo "  Test data : " & dataDir
    WScript.Echo "  Storage   : " & repo & "\WebAPIs\storage"
    WScript.Echo "              - nosql : " & repo & "\WebAPIs\storage\nosql"
    WScript.Echo "              - file  : " & repo & "\WebAPIs\storage\file"
    WScript.Echo ""
End Sub

Function EnsureDotNet
    Dim rc
    rc = shell.Run("cmd /c where dotnet >nul 2>&1", 0, True)
    If rc <> 0 Then
        WScript.Echo "[FATAL] 'dotnet' SDK was not found on PATH. Install the .NET SDK first."
        EnsureDotNet = False
    Else
        WScript.Echo "[OK] .NET SDK found."
        EnsureDotNet = True
    End If
End Function

Function BuildApi
    Dim rc, buildLog
    buildLog = repo & "\poc-build.log"
    WScript.Echo "[..] Building the project (dotnet build)..."
    rc = shell.Run("cmd /c ""cd /d """ & repo & """ && dotnet build """ & proj & """ -v q -nologo > """ & buildLog & """ 2>&1""", 0, True)
    If rc = 0 Then
        WScript.Echo "[OK] Build succeeded."
        BuildApi = True
    Else
        WScript.Echo "[FAIL] Build failed (exit code " & rc & ")."
        WScript.Echo "       See build log: " & buildLog
        BuildApi = False
    End If
End Function

Function StartApi
    Dim cmd
    cmd = "cmd /c ""cd /d """ & repo & """ && set ASPNETCORE_ENVIRONMENT=Development&& dotnet run --no-build --project """ & proj & """ > """ & apiLog & """ 2>&1"""
    On Error Resume Next
    shell.Run cmd, 0, False
    StartApi = (Err.Number = 0)
    On Error GoTo 0
End Function

Sub StopApi
    On Error Resume Next
    shell.Run "taskkill /F /IM WebAPIs.exe /T", 0, True
    ' Give the redirected cmd/dotnet host a moment to release the log file handle.
    WScript.Sleep 800
    On Error GoTo 0
End Sub

Function ServerReady(seconds)
    ServerReady = WaitForServer(seconds)
End Function

Function WaitForServer(timeoutSec)
    Dim started
    started = Now
    Do
        If GetStatusSafe(READY_URL) = 200 Then
            WaitForServer = True
            Exit Function
        End If
        WScript.Sleep 500
    Loop While DateDiff("s", started, Now) < timeoutSec
    WaitForServer = False
End Function

Function GetStatusSafe(url)
    On Error Resume Next
    Dim http
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.Send
    If Err.Number = 0 Then GetStatusSafe = http.Status Else GetStatusSafe = 0
    On Error GoTo 0
End Function

Function GetStatus(url)
    Dim http
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.Send
    GetStatus = http.Status
End Function

Function JsonStatus(jsonBody, idemKey)
    Dim http
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", BASE_URL & "/api/v1/Migration/PushContent", False
    http.SetRequestHeader "Content-Type", "application/json"
    If Len(idemKey) > 0 Then http.SetRequestHeader "Idempotency-Key", idemKey
    http.Send Utf8Bytes(jsonBody)
    JsonStatus = http.Status
End Function

Function UploadStatus(filePath, idemKey)
    Dim boundary, pre, post, st, body, http, fileBytes
    boundary = "----WCMSBoundary" & runId & Int(Timer * 1000)
    pre = "--" & boundary & vbCrLf & _
          "Content-Disposition: form-data; name=""file""; filename=""" & fso.GetFileName(filePath) & """" & vbCrLf & _
          "Content-Type: application/octet-stream" & vbCrLf & vbCrLf
    post = vbCrLf & "--" & boundary & "--" & vbCrLf

    fileBytes = ReadFileBytes(filePath)

    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write Utf8Bytes(pre)
    If LenB(fileBytes) > 0 Then st.Write fileBytes
    st.Write Utf8Bytes(post)
    st.Position = 0
    body = st.Read
    st.Close

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", BASE_URL & "/api/v1/Migration/UploadExcel?notificationEmail=qa@example.com", False
    http.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    If Len(idemKey) > 0 Then http.SetRequestHeader "Idempotency-Key", idemKey
    http.Send body
    UploadStatus = http.Status
End Function

Function BuildJsonItems(count)
    Dim sb, i
    sb = "{""sourceSystem"":""POC-BULK"",""items"":["
    For i = 1 To count
        If i > 1 Then sb = sb & ","
        sb = sb & "{""externalId"":""BULK-" & i & """,""title"":""T" & i & """,""slug"":""s" & i & """,""body"":""b"",""category"":""N""}"
    Next
    sb = sb & "]}"
    BuildJsonItems = sb
End Function

Function Utf8Bytes(s)
    Dim st
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText s
    st.Position = 0
    st.Type = 1
    ' ADODB.Stream prepends a UTF-8 BOM (3 bytes); skip it so payloads are exact.
    If st.Size > 3 Then
        st.Position = 3
        Utf8Bytes = st.Read
    Else
        Utf8Bytes = Empty
    End If
    st.Close
End Function

Function ReadFileBytes(path)
    Dim st
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.LoadFromFile path
    ReadFileBytes = st.Read
    st.Close
End Function

Function B64ToBytes(b64)
    Dim doc, node
    On Error Resume Next
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    If Err.Number <> 0 Then
        Err.Clear
        Set doc = CreateObject("MSXML2.DOMDocument")
    End If
    On Error GoTo 0
    Set node = doc.createElement("bin")
    node.dataType = "bin.base64"
    node.text = b64
    B64ToBytes = node.nodeTypedValue
End Function

Sub WriteBytes(path, bytes)
    Dim st
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write bytes
    st.SaveToFile path, 2
    st.Close
End Sub

Sub WriteText(path, text)
    Dim st
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "us-ascii"
    st.Open
    st.WriteText text
    st.SaveToFile path, 2
    st.Close
End Sub

Function CountFiles(folder)
    Dim total, f, sub_
    total = 0
    If Not fso.FolderExists(folder) Then
        CountFiles = 0
        Exit Function
    End If
    For Each f In fso.GetFolder(folder).Files
        total = total + 1
    Next
    For Each sub_ In fso.GetFolder(folder).SubFolders
        total = total + CountFiles(sub_.Path)
    Next
    CountFiles = total
End Function

Sub Check(name, actual, expected)
    Dim ok
    ok = (CStr(actual) = CStr(expected))
    If ok Then
        passCount = passCount + 1
        WScript.Echo "  [PASS] " & name
    Else
        failCount = failCount + 1
        WScript.Echo "  [FAIL] " & name & "  (expected " & expected & ", got " & actual & ")"
    End If
End Sub

Function IIF(cond, a, b)
    If cond Then IIF = a Else IIF = b
End Function

Sub CleanupRun
    Dim i
    On Error Resume Next
    ' Drop the generated test fixtures; the real captured output lives in WebAPIs\storage.
    If fso.FolderExists(dataDir) Then fso.DeleteFolder dataDir, True
    ' Drop the transient API console log (retry while the handle is still being released).
    For i = 1 To 10
        If Not fso.FileExists(apiLog) Then Exit For
        fso.DeleteFile apiLog, True
        If fso.FileExists(apiLog) Then WScript.Sleep 400
    Next
    On Error GoTo 0
    WScript.Echo "[..] Cleaned up test data (kept: " & repo & "\WebAPIs\storage)."
End Sub

