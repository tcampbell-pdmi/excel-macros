Option Explicit

' =============================================================================
' Excel-to-JSON Request Builder
' =============================================================================
' 1. Reads each row from the active sheet (Row 1 = headers).
' 2. Builds a flat JSON object  { "Header1": "Value1", "Header2": "Value2", ... }
' 3. POSTs the JSON to MIRTH_API_ENDPOINT_URL (application/json).
' 4. Takes the response and POSTs it again to MIRTH_API_ENDPOINT_URL (application/octet-stream).
' 5. Writes both responses and status codes to a new "Results" sheet.
'
' Works on both Windows and Mac (Office 365).
' =============================================================================

' --- Configuration ---
Const MIRTH_API_ENDPOINT_URL As String = "http://localhost:10900"
Const JSON_TO_EDI_API_ENDPOINT_URL  As String = "https://lng2kp9ykj-vpce-0464c135eebb87a6f.execute-api.us-east-1.amazonaws.com:4443/dev/ncpdpd0b1"
Const DELAY_MS As Long = 100        ' ms to wait between every request (0 = no delay)
' ----------------------


Public Sub SendApiRequests()

    Dim wsInput     As Worksheet
    Dim wsOutput    As Worksheet
    Dim lastRow     As Long
    Dim lastCol     As Long
    Dim headers()   As String
    Dim i           As Long
    Dim c           As Long
    Dim jsonBody    As String
    Dim resp1Text   As String
    Dim resp1Status As String
    Dim resp2Text   As String
    Dim resp2Status As String

    Set wsInput = ActiveSheet

    ' -- Determine data boundaries -----------------------------------
    lastRow = wsInput.UsedRange.Rows(wsInput.UsedRange.Rows.Count).Row
    lastCol = wsInput.UsedRange.Columns(wsInput.UsedRange.Columns.Count).Column

    If lastRow < 2 Then
        MsgBox "No data rows found (Row 1 must be headers, data starts at Row 2).", vbExclamation
        Exit Sub
    End If

    ' -- Read headers from Row 1 -------------------------------------
    ReDim headers(1 To lastCol)
    For c = 1 To lastCol
        headers(c) = Trim$(wsInput.Cells(1, c).Value)
    Next c

    ' -- Create or reuse Results sheet ---------------------------------
    If wsInput.Index < ThisWorkbook.Worksheets.Count Then
        Set wsOutput = ThisWorkbook.Worksheets(wsInput.Index + 1)
    Else
        Set wsOutput = ThisWorkbook.Worksheets.Add(After:=wsInput)
        wsOutput.Name = "Results"
    End If

    wsOutput.Cells(1, 1).Value = "JSON Response"
    wsOutput.Cells(1, 2).Value = "JSON Status"
    wsOutput.Cells(1, 3).Value = "Raw Response"
    wsOutput.Cells(1, 4).Value = "Raw Status"

    ' -- Process each data row ------------------------------------------
    For i = 2 To lastRow

        jsonBody = RowToJson(wsInput, headers, i, lastCol)

        If Len(Trim(jsonBody)) > 2 Then   ' more than just "{}"
            ' 1st request: JSON
            #If Mac Then
                Call PostMac(JSON_TO_EDI_API_ENDPOINT_URL, jsonBody, "application/json", resp1Text, resp1Status)
            #Else
                Call PostWindows(JSON_TO_EDI_API_ENDPOINT_URL, jsonBody, "application/json", resp1Text, resp1Status)
            #End If
            wsOutput.Cells(i, 1).Value = resp1Text
            wsOutput.Cells(i, 2).Value = resp1Status

            ' 2nd request: raw/binary - send the 1st response as-is
            If Len(Trim(resp1Text)) > 0 Then
                #If Mac Then
                    Call PostMac(MIRTH_API_ENDPOINT_URL, resp1Text, "application/octet-stream", resp2Text, resp2Status)
                #Else
                    Call PostWindows(MIRTH_API_ENDPOINT_URL, resp1Text, "application/octet-stream", resp2Text, resp2Status)
                #End If
                wsOutput.Cells(i, 3).Value = resp2Text
                wsOutput.Cells(i, 4).Value = resp2Status
            End If
        End If

        Application.StatusBar = "Processing row " & i & " of " & lastRow
        DoEvents

        If DELAY_MS > 0 Then
            Application.Wait Now() + DELAY_MS / 86400000#
        End If

    Next i

    Application.StatusBar = False
    wsOutput.Columns("A:D").AutoFit
    MsgBox "Done! Processed " & (lastRow - 1) & " rows. Results are on the '" & wsOutput.Name & "' sheet.", vbInformation

End Sub


' =============================================================================
' RowToJson - converts a single CSV row into the nested NCPDP D0 JSON structure
' expected by the claim parser API.
' =============================================================================
Private Function RowToJson(ws As Worksheet, _
                           headers() As String, _
                           row As Long, _
                           lastCol As Long) As String

    ' ---- Read all CSV cells into local variables by column index ----
    ' Col 1:  ClaimSetName            (skip)
    ' Col 2:  ClaimName               (skip)
    Dim vSoftwareVendorCertID As String:   vSoftwareVendorCertID = Trim(CStr(ws.Cells(row, 3).Value))
    Dim vCardholderID As String:           vCardholderID = Trim(CStr(ws.Cells(row, 4).Value))
    Dim vPersonCode As String:             vPersonCode = Trim(CStr(ws.Cells(row, 5).Value))
    Dim vGroupID As String:                vGroupID = Trim(CStr(ws.Cells(row, 6).Value))
    Dim vPatRelCode As String:             vPatRelCode = Trim(CStr(ws.Cells(row, 7).Value))
    Dim vPatFirstName As String:           vPatFirstName = Trim(CStr(ws.Cells(row, 8).Value))
    Dim vPatLastName As String:            vPatLastName = Trim(CStr(ws.Cells(row, 9).Value))
    Dim vPatGender As String:              vPatGender = Trim(CStr(ws.Cells(row, 10).Value))
    Dim vDOB As String:                    vDOB = Trim(CStr(ws.Cells(row, 11).Value))
    Dim vBIN As String:                    vBIN = Trim(CStr(ws.Cells(row, 12).Value))
    Dim vGrossAmtDue As String:            vGrossAmtDue = Trim(CStr(ws.Cells(row, 13).Value))
    Dim vPatPaidAmt As String:             vPatPaidAmt = Trim(CStr(ws.Cells(row, 14).Value))
    Dim vIncentiveAmt As String:           vIncentiveAmt = Trim(CStr(ws.Cells(row, 15).Value))
    Dim vOtherAmtQual As String:           vOtherAmtQual = Trim(CStr(ws.Cells(row, 16).Value))
    Dim vOtherAmt As String:               vOtherAmt = Trim(CStr(ws.Cells(row, 17).Value))
    Dim vProductSvcID As String:           vProductSvcID = Trim(CStr(ws.Cells(row, 18).Value))
    Dim vDateRxWritten As String:          vDateRxWritten = Trim(CStr(ws.Cells(row, 19).Value))
    Dim vDateOfService As String:          vDateOfService = Trim(CStr(ws.Cells(row, 20).Value))
    Dim vDAW As String:                    vDAW = Trim(CStr(ws.Cells(row, 21).Value))
    Dim vUCCharge As String:               vUCCharge = Trim(CStr(ws.Cells(row, 22).Value))
    Dim vIngredCost As String:             vIngredCost = Trim(CStr(ws.Cells(row, 23).Value))
    Dim vQtyDispensed As String:           vQtyDispensed = Trim(CStr(ws.Cells(row, 24).Value))
    Dim vDaysSupply As String:             vDaysSupply = Trim(CStr(ws.Cells(row, 25).Value))
    Dim vPrescriberID As String:           vPrescriberID = Trim(CStr(ws.Cells(row, 26).Value))
    Dim vSvcProviderID As String:          vSvcProviderID = Trim(CStr(ws.Cells(row, 27).Value))
    Dim vRxRefNum As String:               vRxRefNum = Trim(CStr(ws.Cells(row, 28).Value))
    Dim vCOBType As String:                vCOBType = Trim(CStr(ws.Cells(row, 29).Value))
    Dim vOPAmtPaidQual As String:          vOPAmtPaidQual = Trim(CStr(ws.Cells(row, 30).Value))
    Dim vOPAmtPaid As String:              vOPAmtPaid = Trim(CStr(ws.Cells(row, 31).Value))
    Dim vOPDate As String:                 vOPDate = Trim(CStr(ws.Cells(row, 32).Value))
    Dim vOPPatRespQual As String:          vOPPatRespQual = Trim(CStr(ws.Cells(row, 33).Value))
    Dim vOPPatRespAmt As String:           vOPPatRespAmt = Trim(CStr(ws.Cells(row, 34).Value))
    Dim vOPRejectCode As String:           vOPRejectCode = Trim(CStr(ws.Cells(row, 35).Value))
    Dim vOPID As String:                   vOPID = Trim(CStr(ws.Cells(row, 36).Value))
    Dim vPriorAuthNum As String:           vPriorAuthNum = Trim(CStr(ws.Cells(row, 37).Value))
    Dim vPriorAuthType As String:          vPriorAuthType = Trim(CStr(ws.Cells(row, 38).Value))
    Dim vPCN As String:                    vPCN = Trim(CStr(ws.Cells(row, 39).Value))
    Dim vBenefitStageQual As String:       vBenefitStageQual = Trim(CStr(ws.Cells(row, 40).Value))
    Dim vBenefitStageAmt As String:        vBenefitStageAmt = Trim(CStr(ws.Cells(row, 41).Value))
    Dim vRxOriginCode As String:           vRxOriginCode = Trim(CStr(ws.Cells(row, 42).Value))
    Dim vPharmSvcType As String:           vPharmSvcType = Trim(CStr(ws.Cells(row, 43).Value))
    Dim vPatResidence As String:           vPatResidence = Trim(CStr(ws.Cells(row, 44).Value))
    Dim vReasonSvcCode As String:          vReasonSvcCode = Trim(CStr(ws.Cells(row, 45).Value))
    Dim vProfSvcCode As String:            vProfSvcCode = Trim(CStr(ws.Cells(row, 46).Value))
    Dim vResultSvcCode As String:          vResultSvcCode = Trim(CStr(ws.Cells(row, 47).Value))
    Dim vDispFee As String:                vDispFee = Trim(CStr(ws.Cells(row, 48).Value))
    Dim vBasisCostDet As String:           vBasisCostDet = Trim(CStr(ws.Cells(row, 49).Value))
    Dim vTxnCode As String:                vTxnCode = Trim(CStr(ws.Cells(row, 50).Value))

    ' ---- Build TransactionHeader ----
    Dim hdr(0 To 8) As String
    hdr(0) = JsonStr("BINNumber_2", vBIN)
    hdr(1) = JsonStr("VersionReleaseNumber_3", "D0")
    hdr(2) = JsonStr("TransactionCode_4", vTxnCode)
    hdr(3) = JsonStr("ProcessorControlNumber_5", vPCN)
    hdr(4) = JsonStr("TransactionCount_6", "1")
    hdr(5) = JsonStr("ServiceProviderIDQualifier_7", "01")
    hdr(6) = JsonStr("ServiceProviderID_8", vSvcProviderID)
    hdr(7) = JsonStr("DateOfService_9", FormatDateCCYYMMDD(vDateOfService))
    hdr(8) = JsonStr("SoftwareVendorCertificationID_10", vSoftwareVendorCertID)
    Dim sHeader As String
    sHeader = """TransactionHeader"": {" & JoinParts(hdr) & "}"

    ' ---- Build Insurance_AM04 ----
    Dim ins(0 To 3) As String
    ins(0) = JsonStr("CardholderID_C2", vCardholderID)
    ins(1) = JsonStr("GroupID_C1", vGroupID)
    ins(2) = JsonStr("PersonCode_C3", vPersonCode)
    ins(3) = JsonNum("PatientRelationshipCode_C6", vPatRelCode)
    Dim sInsurance As String
    sInsurance = """Insurance_AM04"": {" & JoinParts(ins) & "}"

    ' ---- Build Patient_AM01 ----
    Dim pat(0 To 4) As String
    pat(0) = JsonStr("DateOfBirth_C4", FormatDateISO(vDOB))
    pat(1) = JsonNum("PatientGenderCode_C5", vPatGender)
    pat(2) = JsonStr("PatientFirstName_CA", vPatFirstName)
    pat(3) = JsonStr("PatientLastName_CB", vPatLastName)
    pat(4) = JsonNum("PatientResidence_4X", vPatResidence)
    Dim sPatient As String
    sPatient = """Patient_AM01"": {" & JoinParts(pat) & "}"

    ' ---- Build Prescriber_AM03 (inside Transactions[0]) ----
    Dim psc(0 To 1) As String
    psc(0) = JsonStr("PrescriberIDQualifier_EZ", "01")
    psc(1) = JsonStr("PrescriberID_DB", vPrescriberID)
    Dim sPrescriber As String
    sPrescriber = """Prescriber_AM03"": {" & JoinParts(psc) & "}"

    ' ---- Build Claim_AM07 (inside Transactions[0]) ----
    Dim clm(0 To 12) As String
    clm(0) = JsonStr("PrescriptionServiceReferenceNumberQualifier_EM", "1")
    clm(1) = JsonNum("PrescriptionServiceReferenceNumber_D2", vRxRefNum)
    clm(2) = JsonStr("ProductServiceIDQualifier_E1", "03")
    clm(3) = JsonStr("ProductServiceID_D7", vProductSvcID)
    clm(4) = JsonNum("QuantityDispensed_E7", vQtyDispensed)
    clm(5) = JsonNum("DaysSupply_D5", vDaysSupply)
    clm(6) = JsonStr("DispenseAsWritten_D8", vDAW)
    clm(7) = JsonStr("DatePrescriptionWritten_DE", FormatDateISO(vDateRxWritten))
    clm(8) = JsonNum("PrescriptionOriginCode_DJ", vRxOriginCode)
    clm(9) = JsonNum("OtherCoverageCode_C8", vCOBType)
    clm(10) = JsonNum("PharmacyServiceType_U7", vPharmSvcType)
    clm(11) = JsonStr("PriorAuthorizationNumberSubmitted_EV", vPriorAuthNum)
    clm(12) = JsonStr("PriorAuthorizationTypeCode_EU", vPriorAuthType)
    Dim sClaim As String
    sClaim = """Claim_AM07"": {" & JoinParts(clm) & "}"

    ' ---- Build Pricing_AM11 (inside Transactions[0]) ----
    Dim prc(0 To 7) As String
    prc(0) = JsonNum("IngredientCostSubmitted_D9", vIngredCost)
    prc(1) = JsonNum("DispensingFeeSubmitted_DC", vDispFee)
    prc(2) = JsonNum("UsualAndCustomaryCharge_DQ", vUCCharge)
    prc(3) = JsonNum("GrossAmountDue_DU", vGrossAmtDue)
    prc(4) = JsonStr("BasisOfCostDetermination_DN", vBasisCostDet)
    prc(5) = JsonNum("PatientPaidAmountSubmitted_DX", vPatPaidAmt)
    prc(6) = JsonNum("IncentiveAmountSubmitted_E3", vIncentiveAmt)
    ' OtherAmountClaimed sub-array (if qualifier is present)
    Dim sOtherAmtArr As String: sOtherAmtArr = ""
    If Len(vOtherAmtQual) > 0 Then
        sOtherAmtArr = """OtherAmountClaimedSubmittedCount_H7"": 1, " & _
            """OtherAmountClaimedSubmitteds"": [{" & _
            JsonStr("OtherAmountClaimedSubmittedQualifier_H8", vOtherAmtQual) & ", " & _
            JsonNum("OtherAmountClaimedSubmitted_H9", vOtherAmt) & "}]"
    End If
    prc(7) = sOtherAmtArr
    Dim sPricing As String
    sPricing = """Pricing_AM11"": {" & JoinParts(prc) & "}"

    ' ---- Build CoordinationOfBenefitsOtherPayments_AM05 (conditional) ----
    Dim sCOB As String: sCOB = ""
    Dim hasCOB As Boolean
    hasCOB = (Len(vOPID) > 0 Or Len(vOPDate) > 0 Or Len(vOPAmtPaid) > 0 Or _
              Len(vOPPatRespAmt) > 0 Or Len(vOPRejectCode) > 0)
    If hasCOB Then
        Dim cobInner As String: cobInner = ""
        ' OtherPayerID
        If Len(vOPID) > 0 Then
            cobInner = cobInner & JsonStr("OtherPayerIDQualifier_6C", "03") & ", "
            cobInner = cobInner & JsonStr("OtherPayerID_7C", vOPID) & ", "
        End If
        ' OtherPayerCoverageType
        cobInner = cobInner & JsonStr("OtherPayerCoverageType_5C", "01") & ", "
        ' OtherPayerDate
        If Len(vOPDate) > 0 Then
            cobInner = cobInner & JsonStr("OtherPayerDate_E8", FormatDateISO(vOPDate)) & ", "
        End If
        ' OtherPayerPatientResponsibilityAmounts sub-array
        If Len(vOPPatRespQual) > 0 And Len(vOPPatRespAmt) > 0 Then
            cobInner = cobInner & """OtherPayerPatientResponsibilityAmountCount_NR"": 1, "
            cobInner = cobInner & """OtherPayerPatientResponsibilityAmounts"": [{" & _
                JsonStr("OtherPayerPatientResponsibilityAmountQualifier_NP", vOPPatRespQual) & ", " & _
                JsonNum("OtherPayerPatientResponsibilityAmount_NQ", vOPPatRespAmt) & "}], "
        End If
        ' OtherPayerAmountPaids sub-array
        If Len(vOPAmtPaidQual) > 0 And Len(vOPAmtPaid) > 0 Then
            cobInner = cobInner & """OtherPayerAmountPaidCount_HB"": 1, "
            cobInner = cobInner & """OtherPayerAmountPaids"": [{" & _
                JsonStr("OtherPayerAmountPaidQualifier_HC", vOPAmtPaidQual) & ", " & _
                JsonNum("OtherPayerAmountPaid_DV", vOPAmtPaid) & "}], "
        End If
        ' OtherPayerRejects sub-array
        If Len(vOPRejectCode) > 0 Then
            cobInner = cobInner & """OtherPayerRejectCount_5E"": 1, "
            cobInner = cobInner & """OtherPayerRejects"": [{" & _
                JsonStr("OtherPayerRejectCode_6E", vOPRejectCode) & "}], "
        End If
        ' BenefitStages sub-array
        If Len(vBenefitStageQual) > 0 And Len(vBenefitStageAmt) > 0 Then
            cobInner = cobInner & """BenefitStageCount_MU"": ""1"", "
            cobInner = cobInner & """BenefitStages"": [{" & _
                JsonStr("BenefitStageQualifier_MV", vBenefitStageQual) & ", " & _
                JsonNum("BenefitStageAmount_MW", vBenefitStageAmt) & "}], "
        End If
        ' Trim trailing ", "
        If Right(cobInner, 2) = ", " Then cobInner = Left(cobInner, Len(cobInner) - 2)
        sCOB = """CoordinationOfBenefitsOtherPayments_AM05"": {" & _
            """CoordinationOfBenefitsOtherPaymentsCount_4C"": 1, " & _
            """CoordinationOfBenefitsOtherPayments"": [{" & cobInner & "}]}"
    End If

    ' ---- Build DURPPS_AM08 (conditional) ----
    Dim sDUR As String: sDUR = ""
    If Len(vReasonSvcCode) > 0 Or Len(vProfSvcCode) > 0 Or Len(vResultSvcCode) > 0 Then
        Dim durParts(0 To 2) As String
        durParts(0) = JsonStr("ReasonForServiceCode_E4", vReasonSvcCode)
        durParts(1) = JsonStr("ProfessionalServiceCode_E5", vProfSvcCode)
        durParts(2) = JsonStr("ResultofServiceCode_E6", vResultSvcCode)
        Dim durInner As String
        durInner = JoinParts(durParts)
        If Len(durInner) > 0 Then
            sDUR = """DURPPS_AM08"": {""DURPPSCodeCounter_7E"": 1, ""DURPPSCodes"": [{" & durInner & "}]}"
        End If
    End If

    ' ---- Assemble Transaction object ----
    Dim txnParts(0 To 4) As String
    txnParts(0) = sClaim
    txnParts(1) = sPricing
    txnParts(2) = sPrescriber
    txnParts(3) = sCOB
    txnParts(4) = sDUR

    Dim sTransaction As String
    sTransaction = """Transactions"": [{" & JoinParts(txnParts) & "}]"

    ' ---- Assemble full JSON ----
    Dim topParts(0 To 3) As String
    topParts(0) = sHeader
    topParts(1) = sInsurance
    topParts(2) = sPatient
    topParts(3) = sTransaction

    RowToJson = "{" & JoinParts(topParts) & "}"

End Function


' =============================================================================
' EscapeJsonString - minimal JSON string escaping.
' =============================================================================
Private Function EscapeJsonString(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscapeJsonString = s
End Function


' =============================================================================
' Helper: Format a date from M/D/YY (or M/D/YYYY) to YYYY-MM-DD
' =============================================================================
Private Function FormatDateISO(ByVal s As String) As String
    Dim d As Date
    If Len(Trim(s)) = 0 Then
        FormatDateISO = ""
        Exit Function
    End If
    On Error Resume Next
    d = CDate(s)
    If Err.Number <> 0 Then
        FormatDateISO = ""
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    FormatDateISO = Format(d, "YYYY-MM-DD")
End Function


' =============================================================================
' Helper: Format a date from M/D/YY to CCYYMMDD (no dashes)
' =============================================================================
Private Function FormatDateCCYYMMDD(ByVal s As String) As String
    Dim d As Date
    If Len(Trim(s)) = 0 Then
        FormatDateCCYYMMDD = ""
        Exit Function
    End If
    On Error Resume Next
    d = CDate(s)
    If Err.Number <> 0 Then
        FormatDateCCYYMMDD = ""
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    FormatDateCCYYMMDD = Format(d, "YYYYMMDD")
End Function


' =============================================================================
' Helper: Return a JSON string value: "key": "value"  (empty if val is blank)
' =============================================================================
Private Function JsonStr(ByVal key As String, ByVal val As String) As String
    If Len(Trim(val)) = 0 Then
        JsonStr = ""
    Else
        JsonStr = """" & key & """: """ & EscapeJsonString(val) & """"
    End If
End Function


' =============================================================================
' Helper: Return a JSON numeric value: "key": 123.45  (empty if val is blank)
' =============================================================================
Private Function JsonNum(ByVal key As String, ByVal val As String) As String
    If Len(Trim(val)) = 0 Then
        JsonNum = ""
    ElseIf IsNumeric(val) Then
        JsonNum = """" & key & """: " & CDbl(val)
    Else
        JsonNum = ""
    End If
End Function


' =============================================================================
' Helper: Join non-empty JSON parts with commas
' =============================================================================
Private Function JoinParts(parts() As String) As String
    Dim result As String
    Dim i As Long
    result = ""
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            If Len(result) > 0 Then result = result & ", "
            result = result & parts(i)
        End If
    Next i
    JoinParts = result
End Function


' =============================================================================
' Windows: uses WinHttp COM object
' =============================================================================
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