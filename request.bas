' --- Configuration ---
Const API_URL As String = "http://localhost:10900"
Const DELAY_MS As Long = 100        ' ms to wait between every request (0 = no delay)
' ----------------------

Sub SendApiRequests()
    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim url As String
    Dim requestBody As String
    Dim respText As String
    Dim respStatus As String

    Set wsInput = ActiveSheet

    ' -- Create or reuse Results sheet ---------------------------------
    If wsInput.Index < ThisWorkbook.Worksheets.Count Then
        Set wsOutput = ThisWorkbook.Worksheets(wsInput.Index + 1)
    Else
        Set wsOutput = ThisWorkbook.Worksheets.Add(After:=wsInput)
        wsOutput.Name = "Results"
    End If

    wsOutput.Cells(1, 1).Value = "Response"
    wsOutput.Cells(1, 2).Value = "Status"

    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        requestBody = wsInput.Cells(i, 1).Value

        If Len(Trim(requestBody)) > 0 Then
            #If Mac Then
                Call PostMac(API_URL, requestBody, "text/plain", respText, respStatus)
            #Else
                Call PostWindows(API_URL, requestBody, "text/plain", respText, respStatus)
            #End If
            wsOutput.Cells(i, 1).Value = respText
            wsOutput.Cells(i, 2).Value = respStatus
        End If

        Application.StatusBar = "Processing row " & i & " of " & lastRow

        ' Rate limiting: per-request delay
        If DELAY_MS > 0 Then
            Application.Wait Now() + DELAY_MS / 86400000#
        End If

    Next i

    Application.StatusBar = False
    wsOutput.Columns("A:B").AutoFit
    MsgBox "Done! Processed " & (lastRow - 1) & " rows. Results are on the '" & wsOutput.Name & "' sheet.", vbInformation
End Sub

' Windows: uses WinHttp COM object
Sub PostWindows(url As String, body As String, contentType As String, ByRef respText As String, ByRef respStatus As String)
    Dim httpReq As Object
    On Error Resume Next
    Set httpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    If httpReq Is Nothing Then Set httpReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If httpReq Is Nothing Then Set httpReq = CreateObject("MSXML2.ServerXMLHTTP")

    If Not httpReq Is Nothing Then
        Err.Clear
        httpReq.SetTimeouts 5000, 5000, 30000, 30000  ' resolve, connect, send, receive (ms)
        httpReq.Open "POST", url, True  ' True = async; keeps Excel UI responsive
        httpReq.setRequestHeader "Content-Type", contentType
        httpReq.send body
        ' Poll until complete, yielding control on each iteration
        Dim startT As Double
        startT = Timer
        Do While httpReq.ReadyState <> 4  ' 4 = READYSTATE_COMPLETE
            DoEvents
            If Timer - startT > 35 Then
                respText = "ERROR: Request timed out"
                respStatus = "Timeout"
                Set httpReq = Nothing
                On Error GoTo 0
                Exit Sub
            End If
        Loop
        If Err.Number = 0 Then
            respText = httpReq.responseText
            respStatus = httpReq.Status & " " & httpReq.statusText
        Else
            respText = "ERROR: " & Err.Description
            respStatus = "Failed"
            Err.Clear
        End If
    Else
        respText = "ERROR: Could not create HTTP object"
        respStatus = "Failed"
    End If
    On Error GoTo 0
    Set httpReq = Nothing
End Sub


#If Mac Then
' =============================================================================
' Mac: writes body to a temp file then POSTs it via curl + polling
' =============================================================================
Sub PostMac(url As String, body As String, contentType As String, ByRef respText As String, ByRef respStatus As String)
    Dim tmpDir As String
    Dim bodyFile As String
    Dim respFile As String
    Dim codeFile As String
    Dim doneFile As String
    Dim scriptFile As String
    Dim fn As Integer
    Dim j As Long
    Dim b As Byte
    Dim lineStr As String

    tmpDir = Environ("TMPDIR")
    If Len(tmpDir) = 0 Then tmpDir = "/tmp/"
    If Right(tmpDir, 1) <> "/" Then tmpDir = tmpDir & "/"
    bodyFile = tmpDir & "vba_req_body.bin"
    respFile = tmpDir & "vba_req_resp.txt"
    codeFile = tmpDir & "vba_req_code.txt"
    doneFile = tmpDir & "vba_req_done.txt"
    scriptFile = tmpDir & "vba_req.sh"

    On Error Resume Next
    Kill doneFile
    Kill codeFile
    On Error GoTo 0

    ' Write body to temp file byte-by-byte
    On Error Resume Next
    fn = FreeFile
    Open bodyFile For Binary As #fn
    For j = 1 To Len(body)
        b = Asc(Mid(body, j, 1))
        Put #fn, , b
    Next j
    Close #fn
    If Err.Number <> 0 Then
        respText = "ERROR writing temp file: " & Err.Description
        respStatus = "Failed"
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' Write shell script: curl with JSON content type
    On Error Resume Next
    fn = FreeFile
    Open scriptFile For Output As #fn
    Print #fn, "#!/bin/sh"
    Print #fn, "curl -s -X POST -H 'Content-Type: " & contentType & "' --data-binary '@" & bodyFile & "' -o '" & respFile & "' -w '%{http_code}' '" & url & "' > '" & codeFile & "'"
    Print #fn, "echo 1 > '" & doneFile & "'"
    Close #fn
    If Err.Number <> 0 Then
        respText = "ERROR writing script: " & Err.Description
        respStatus = "Failed"
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Shell "/bin/sh " & scriptFile

    ' Poll for the done sentinel file (max 60 seconds)
    Dim startT As Double
    startT = Timer
    Do
        Application.Wait Now() + (DELAY_MS / 86400000#)
        DoEvents
        If Dir(doneFile) <> "" Then Exit Do
        If Timer - startT > 60 Then
            respText = "ERROR: Request timed out after 60 seconds"
            respStatus = "Timeout"
            GoTo Cleanup
        End If
    Loop

    ' Read HTTP status code
    Dim httpCode As String
    httpCode = ""
    On Error Resume Next
    fn = FreeFile
    Open codeFile For Input As #fn
    Line Input #fn, httpCode
    httpCode = Trim(httpCode)
    Close #fn
    On Error GoTo 0
    If Len(httpCode) = 0 Then httpCode = "0"

    ' Read response body
    respText = ""
    On Error Resume Next
    fn = FreeFile
    Open respFile For Input As #fn
    Do While Not EOF(fn)
        Line Input #fn, lineStr
        respText = respText & lineStr & vbLf
    Loop
    Close #fn
    If Len(respText) > 0 Then respText = Left(respText, Len(respText) - 1)
    On Error GoTo 0

    respStatus = httpCode

Cleanup:
    On Error Resume Next
    Kill bodyFile
    Kill respFile
    Kill codeFile
    Kill doneFile
    Kill scriptFile
    On Error GoTo 0
End Sub
#End If
