Option Explicit

' =============================================================================
' Excel-to-JSON Request Builder
' =============================================================================
' 1. Reads each row from the active sheet (Row 1 = headers).
' 2. Builds a nested NCPDP D0 JSON object from column headers.
'    Column order does not matter; fields are looked up by header name.
' 3. POSTs the JSON to JSON_TO_EDI_API_ENDPOINT_URL (application/json).
' 4. Takes the EDI response and POSTs it to MIRTH_API_ENDPOINT_URL (text/plain).
' 5. Converts the EDI response back to JSON via EDI_TO_JSON_API_ENDPOINT_URL.
' 6. Writes all responses and status codes to "Results" and "Response Data" sheets.
'
' Works on both Windows and Mac (Office 365).
' =============================================================================

' --- Configuration ---
Const MIRTH_API_ENDPOINT_URL As String = "http://localhost:10900"
Const JSON_TO_EDI_API_ENDPOINT_URL  As String = "https://lng2kp9ykj-vpce-0464c135eebb87a6f.execute-api.us-east-1.amazonaws.com:4443/dev/ncpdpd0b1"
Const EDI_TO_JSON_API_ENDPOINT_URL  As String = "https://lng2kp9ykj-vpce-0464c135eebb87a6f.execute-api.us-east-1.amazonaws.com:4443/dev/standardd0"
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
    Dim resp3Text   As String
    Dim resp3Status As String
    Dim wsResponseData As Worksheet
    Dim nextRespCol As Long
    Dim claimSetCol As Long
    Dim claimNameCol As Long

    Set wsInput = ActiveSheet

    ' -- Determine data boundaries -----------------------------------
    lastCol = wsInput.Cells(1, wsInput.Columns.Count).End(xlToLeft).Column
    lastRow = 1
    Dim tmpRow As Long
    For c = 1 To lastCol
        tmpRow = wsInput.Cells(wsInput.Rows.Count, c).End(xlUp).Row
        If tmpRow > lastRow Then lastRow = tmpRow
    Next c

    If lastRow < 2 Then
        MsgBox "No data rows found on sheet '" & wsInput.Name & "' (Row 1 must be headers, data starts at Row 2).", vbExclamation
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

    wsOutput.Cells(1, 1).Value = "JSON Sent"
    wsOutput.Cells(1, 2).Value = "JSON Response"
    wsOutput.Cells(1, 3).Value = "JSON Status"
    wsOutput.Cells(1, 4).Value = "Raw Response"
    wsOutput.Cells(1, 5).Value = "Raw Status"
    wsOutput.Cells(1, 6).Value = "EDI-JSON Status"

    ' -- Create or reuse Response Data sheet ----------------------------
    If wsOutput.Index < ThisWorkbook.Worksheets.Count Then
        Set wsResponseData = ThisWorkbook.Worksheets(wsOutput.Index + 1)
    Else
        Set wsResponseData = ThisWorkbook.Worksheets.Add(After:=wsOutput)
        wsResponseData.Name = "Response Data"
    End If
    wsResponseData.Cells(1, 1).Value = "ClaimSetName"
    wsResponseData.Cells(1, 2).Value = "ClaimName"
    nextRespCol = 3
    claimSetCol = GetCol(headers, "ClaimSetName")
    claimNameCol = GetCol(headers, "ClaimName")

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
            wsOutput.Cells(i, 1).Value = jsonBody
            wsOutput.Cells(i, 2).Value = resp1Text
            wsOutput.Cells(i, 3).Value = resp1Status

            ' 2nd request: raw/binary - send the 1st response as-is
            If Len(Trim(resp1Text)) > 0 Then
                #If Mac Then
                    Call PostMac(MIRTH_API_ENDPOINT_URL, resp1Text, "text/plain", resp2Text, resp2Status)
                #Else
                    Call PostWindows(MIRTH_API_ENDPOINT_URL, resp1Text, "text/plain", resp2Text, resp2Status)
                #End If
                wsOutput.Cells(i, 4).Value = resp2Text
                wsOutput.Cells(i, 5).Value = resp2Status

                ' 3rd request: convert EDI response back to JSON
                If Len(Trim(resp2Text)) > 0 Then
                    #If Mac Then
                        Call PostMac(EDI_TO_JSON_API_ENDPOINT_URL, resp2Text, "text/plain", resp3Text, resp3Status)
                    #Else
                        Call PostWindows(EDI_TO_JSON_API_ENDPOINT_URL, resp2Text, "text/plain", resp3Text, resp3Status)
                    #End If
                    wsOutput.Cells(i, 6).Value = resp3Status

                    ' Parse JSON response into Response Data sheet
                    If Len(Trim(resp3Text)) > 0 Then
                        If claimSetCol > 0 Then wsResponseData.Cells(i, 1).Value = wsInput.Cells(i, claimSetCol).Value
                        If claimNameCol > 0 Then wsResponseData.Cells(i, 2).Value = wsInput.Cells(i, claimNameCol).Value
                        ParseJsonToCells wsResponseData, i, resp3Text, nextRespCol
                    End If
                End If
            End If
        End If

        Application.StatusBar = "Processing row " & i & " of " & lastRow
        DoEvents

        If DELAY_MS > 0 Then
            Application.Wait Now() + DELAY_MS / 86400000#
        End If

    Next i

    Application.StatusBar = False
    wsOutput.Columns("A:G").AutoFit
    wsResponseData.Columns.AutoFit
    MsgBox "Done! Processed " & (lastRow - 1) & " rows. Results are on the '" & wsOutput.Name & "' and '" & wsResponseData.Name & "' sheets.", vbInformation

End Sub


' =============================================================================
' RowToJson - converts a single row into the nested NCPDP D0 JSON structure
' expected by the claim parser API.
' Phase 1: Iterates all columns dynamically to collect key/value pairs.
' Phase 2: Builds nested JSON segments by looking up values by field name.
' =============================================================================
Private Function RowToJson(ws As Worksheet, _
                           headers() As String, _
                           row As Long, _
                           lastCol As Long) As String

    ' ---- Phase 1: Collect all field values by header name ----
    Dim fNames() As String
    Dim fVals()  As String
    Dim fCount   As Long
    Dim c        As Long
    Dim hdr      As String
    Dim val      As String

    ReDim fNames(1 To lastCol)
    ReDim fVals(1 To lastCol)
    fCount = 0

    For c = 1 To lastCol
        hdr = Trim$(headers(c))
        If LCase(hdr) = "claimsetname" Or LCase(hdr) = "claimname" Then GoTo NextCol
        If Len(hdr) = 0 Then GoTo NextCol

        val = Trim$(CStr(ws.Cells(row, c).Value))
        fCount = fCount + 1
        fNames(fCount) = hdr
        fVals(fCount) = val
NextCol:
    Next c

    ' ---- Phase 2: Build nested JSON segments ----
    ' Shorthand lookup
    Dim v As String

    ' -- TransactionHeader --
    v = LookupVal(fNames, fVals, fCount, "TransactionCode_4")
    If Len(v) = 0 Then v = "B1"

    Dim txHdr(0 To 8) As String
    txHdr(0) = JsonStr("BINNumber_2", LookupVal(fNames, fVals, fCount, "BINNumber_2"))
    txHdr(1) = JsonStr("VersionReleaseNumber_3", "D0")
    txHdr(2) = JsonStr("TransactionCode_4", v)
    txHdr(3) = JsonStr("ProcessorControlNumber_5", LookupVal(fNames, fVals, fCount, "ProcessorControlNumber_5"))
    txHdr(4) = JsonStr("TransactionCount_6", "1")
    txHdr(5) = JsonStr("ServiceProviderIDQualifier_7", "01")
    txHdr(6) = JsonStr("ServiceProviderID_8", LookupVal(fNames, fVals, fCount, "ServiceProviderID_8"))
    txHdr(7) = JsonStr("DateOfService_9", FormatDateCCYYMMDD(LookupVal(fNames, fVals, fCount, "DateOfService_9")))
    txHdr(8) = JsonStr("SoftwareVendorCertificationID_10", LookupVal(fNames, fVals, fCount, "SoftwareVendorCertificationID_10"))
    Dim sHeader As String
    sHeader = """TransactionHeader"": {" & JoinParts(txHdr) & "}"

    ' -- Insurance_AM04 --
    Dim ins(0 To 3) As String
    ins(0) = JsonStr("CardholderID_C2", LookupVal(fNames, fVals, fCount, "CardholderID_C2"))
    ins(1) = JsonStr("GroupID_C1", LookupVal(fNames, fVals, fCount, "GroupID_C1"))
    ins(2) = JsonStr("PersonCode_C3", LookupVal(fNames, fVals, fCount, "PersonCode_C3"))
    ins(3) = JsonNum("PatientRelationshipCode_C6", LookupVal(fNames, fVals, fCount, "PatientRelationshipCode_C6"))
    Dim sInsurance As String
    sInsurance = """Insurance_AM04"": {" & JoinParts(ins) & "}"

    ' -- Patient_AM01 --
    Dim pat(0 To 4) As String
    pat(0) = JsonStr("DateOfBirth_C4", FormatDateISO(LookupVal(fNames, fVals, fCount, "DateOfBirth_C4")))
    pat(1) = JsonNum("PatientGenderCode_C5", LookupVal(fNames, fVals, fCount, "PatientGenderCode_C5"))
    pat(2) = JsonStr("PatientFirstName_CA", LookupVal(fNames, fVals, fCount, "PatientFirstName_CA"))
    pat(3) = JsonStr("PatientLastName_CB", LookupVal(fNames, fVals, fCount, "PatientLastName_CB"))
    pat(4) = JsonNum("PatientResidence_4X", LookupVal(fNames, fVals, fCount, "PatientResidence_4X"))
    Dim sPatient As String
    sPatient = """Patient_AM01"": {" & JoinParts(pat) & "}"

    ' -- Prescriber_AM03 (inside Transactions[0]) --
    Dim psc(0 To 1) As String
    psc(0) = JsonStr("PrescriberIDQualifier_EZ", "01")
    psc(1) = JsonStr("PrescriberID_DB", LookupVal(fNames, fVals, fCount, "PrescriberID_DB"))
    Dim sPrescriber As String
    sPrescriber = """Prescriber_AM03"": {" & JoinParts(psc) & "}"

    ' -- Claim_AM07 (inside Transactions[0]) --
    Dim clm(0 To 12) As String
    clm(0) = JsonStr("PrescriptionServiceReferenceNumberQualifier_EM", "1")
    clm(1) = JsonNum("PrescriptionServiceReferenceNumber_D2", LookupVal(fNames, fVals, fCount, "PrescriptionServiceReferenceNumber_D2"))
    clm(2) = JsonStr("ProductServiceIDQualifier_E1", "03")
    clm(3) = JsonStr("ProductServiceID_D7", LookupVal(fNames, fVals, fCount, "ProductServiceID_D7"))
    clm(4) = JsonNum("QuantityDispensed_E7", LookupVal(fNames, fVals, fCount, "QuantityDispensed_E7"))
    clm(5) = JsonNum("DaysSupply_D5", LookupVal(fNames, fVals, fCount, "DaysSupply_D5"))
    clm(6) = JsonStr("DispenseAsWritten_D8", LookupVal(fNames, fVals, fCount, "DispenseAsWritten_D8"))
    clm(7) = JsonStr("DatePrescriptionWritten_DE", FormatDateISO(LookupVal(fNames, fVals, fCount, "DatePrescriptionWritten_DE")))
    clm(8) = JsonNum("PrescriptionOriginCode_DJ", LookupVal(fNames, fVals, fCount, "PrescriptionOriginCode_DJ"))
    clm(9) = JsonNum("OtherCoverageCode_C8", LookupVal(fNames, fVals, fCount, "OtherCoverageCode_C8"))
    clm(10) = JsonNum("PharmacyServiceType_U7", LookupVal(fNames, fVals, fCount, "PharmacyServiceType_U7"))
    clm(11) = JsonStr("PriorAuthorizationNumberSubmitted_EV", LookupVal(fNames, fVals, fCount, "PriorAuthorizationNumberSubmitted_EV"))
    clm(12) = JsonStr("PriorAuthorizationTypeCode_EU", LookupVal(fNames, fVals, fCount, "PriorAuthorizationTypeCode_EU"))
    Dim sClaim As String
    sClaim = """Claim_AM07"": {" & JoinParts(clm) & "}"

    ' -- Pricing_AM11 (inside Transactions[0]) --
    Dim vOtherAmtQual As String: vOtherAmtQual = LookupVal(fNames, fVals, fCount, "OtherAmountClaimedSubmittedQualifier_H8")
    Dim vOtherAmt As String:     vOtherAmt = LookupVal(fNames, fVals, fCount, "OtherAmountClaimedSubmitted_H9")

    Dim prc(0 To 7) As String
    prc(0) = JsonNum("IngredientCostSubmitted_D9", LookupVal(fNames, fVals, fCount, "IngredientCostSubmitted_D9"))
    prc(1) = JsonNum("DispensingFeeSubmitted_DC", LookupVal(fNames, fVals, fCount, "DispensingFeeSubmitted_DC"))
    prc(2) = JsonNum("UsualAndCustomaryCharge_DQ", LookupVal(fNames, fVals, fCount, "UsualAndCustomaryCharge_DQ"))
    prc(3) = JsonNum("GrossAmountDue_DU", LookupVal(fNames, fVals, fCount, "GrossAmountDue_DU"))
    prc(4) = JsonStr("BasisOfCostDetermination_DN", LookupVal(fNames, fVals, fCount, "BasisOfCostDetermination_DN"))
    prc(5) = JsonNum("PatientPaidAmountSubmitted_DX", LookupVal(fNames, fVals, fCount, "PatientPaidAmountSubmitted_DX"))
    prc(6) = JsonNum("IncentiveAmountSubmitted_E3", LookupVal(fNames, fVals, fCount, "IncentiveAmountSubmitted_E3"))
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

    ' -- CoordinationOfBenefitsOtherPayments_AM05 (conditional) --
    Dim vOPID As String:          vOPID = LookupVal(fNames, fVals, fCount, "OtherPayerID_7C")
    Dim vOPDate As String:        vOPDate = LookupVal(fNames, fVals, fCount, "OtherPayerDate_E8")
    Dim vOPAmtPaidQual As String: vOPAmtPaidQual = LookupVal(fNames, fVals, fCount, "OtherPayerAmountPaidQualifier_HC")
    Dim vOPAmtPaid As String:     vOPAmtPaid = LookupVal(fNames, fVals, fCount, "OtherPayerAmountPaid_DV")
    Dim vOPPatRespQual As String: vOPPatRespQual = LookupVal(fNames, fVals, fCount, "OtherPayerPatientResponsibilityAmountQualifier_NP")
    Dim vOPPatRespAmt As String:  vOPPatRespAmt = LookupVal(fNames, fVals, fCount, "OtherPayerPatientResponsibilityAmount_NQ")
    Dim vOPRejectCode As String:  vOPRejectCode = LookupVal(fNames, fVals, fCount, "OtherPayerRejectCode_6E")
    Dim vBenefitStageQual As String: vBenefitStageQual = LookupVal(fNames, fVals, fCount, "BenefitStageQualifier_MV")
    Dim vBenefitStageAmt As String:  vBenefitStageAmt = LookupVal(fNames, fVals, fCount, "BenefitStageAmount_MW")

    Dim sCOB As String: sCOB = ""
    Dim hasCOB As Boolean
    hasCOB = (Len(vOPID) > 0 Or Len(vOPDate) > 0 Or Len(vOPAmtPaid) > 0 Or _
              Len(vOPPatRespAmt) > 0 Or Len(vOPRejectCode) > 0)
    If hasCOB Then
        Dim cobInner As String: cobInner = ""
        If Len(vOPID) > 0 Then
            cobInner = cobInner & JsonStr("OtherPayerIDQualifier_6C", "03") & ", "
            cobInner = cobInner & JsonStr("OtherPayerID_7C", vOPID) & ", "
        End If
        cobInner = cobInner & JsonStr("OtherPayerCoverageType_5C", "01") & ", "
        If Len(vOPDate) > 0 Then
            cobInner = cobInner & JsonStr("OtherPayerDate_E8", FormatDateISO(vOPDate)) & ", "
        End If
        If Len(vOPPatRespQual) > 0 And Len(vOPPatRespAmt) > 0 Then
            cobInner = cobInner & """OtherPayerPatientResponsibilityAmountCount_NR"": 1, "
            cobInner = cobInner & """OtherPayerPatientResponsibilityAmounts"": [{" & _
                JsonStr("OtherPayerPatientResponsibilityAmountQualifier_NP", vOPPatRespQual) & ", " & _
                JsonNum("OtherPayerPatientResponsibilityAmount_NQ", vOPPatRespAmt) & "}], "
        End If
        If Len(vOPAmtPaidQual) > 0 And Len(vOPAmtPaid) > 0 Then
            cobInner = cobInner & """OtherPayerAmountPaidCount_HB"": 1, "
            cobInner = cobInner & """OtherPayerAmountPaids"": [{" & _
                JsonStr("OtherPayerAmountPaidQualifier_HC", vOPAmtPaidQual) & ", " & _
                JsonNum("OtherPayerAmountPaid_DV", vOPAmtPaid) & "}], "
        End If
        If Len(vOPRejectCode) > 0 Then
            cobInner = cobInner & """OtherPayerRejectCount_5E"": 1, "
            cobInner = cobInner & """OtherPayerRejects"": [{" & _
                JsonStr("OtherPayerRejectCode_6E", vOPRejectCode) & "}], "
        End If
        If Len(vBenefitStageQual) > 0 And Len(vBenefitStageAmt) > 0 Then
            cobInner = cobInner & """BenefitStageCount_MU"": ""1"", "
            cobInner = cobInner & """BenefitStages"": [{" & _
                JsonStr("BenefitStageQualifier_MV", vBenefitStageQual) & ", " & _
                JsonNum("BenefitStageAmount_MW", vBenefitStageAmt) & "}], "
        End If
        If Right(cobInner, 2) = ", " Then cobInner = Left(cobInner, Len(cobInner) - 2)
        sCOB = """CoordinationOfBenefitsOtherPayments_AM05"": {" & _
            """CoordinationOfBenefitsOtherPaymentsCount_4C"": 1, " & _
            """CoordinationOfBenefitsOtherPayments"": [{" & cobInner & "}]}"
    End If

    ' -- DURPPS_AM08 (conditional) --
    Dim vReasonSvcCode As String:  vReasonSvcCode = LookupVal(fNames, fVals, fCount, "ReasonForServiceCode_E4")
    Dim vProfSvcCode As String:    vProfSvcCode = LookupVal(fNames, fVals, fCount, "ProfessionalServiceCode_E5")
    Dim vResultSvcCode As String:  vResultSvcCode = LookupVal(fNames, fVals, fCount, "ResultofServiceCode_E6")

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
' Helper: Look up a value by field name from collected row data.
' Returns empty string if the field is missing or has no value.
' =============================================================================
Private Function LookupVal(fNames() As String, fVals() As String, _
                           fCount As Long, ByVal key As String) As String
    Dim i As Long
    For i = 1 To fCount
        If fNames(i) = key Then
            LookupVal = fVals(i)
            Exit Function
        End If
    Next i
    LookupVal = ""
End Function


' =============================================================================
' GetCol - returns the 1-based column index for a given header name.
' Returns 0 if not found.
' =============================================================================
Private Function GetCol(headers() As String, ByVal name As String) As Long
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If headers(i) = name Then
            GetCol = i
            Exit Function
        End If
    Next i
    GetCol = 0
End Function


' =============================================================================
' ExtractJsonValue - simple string-based JSON value extraction.
' Searches for "keyName": in the JSON string and returns the value.
' =============================================================================
Private Function ExtractJsonValue(ByVal json As String, ByVal keyName As String) As String
    Dim searchKey As String
    Dim pos As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim ch As String
    Dim depth As Long
    Dim inString As Boolean

    searchKey = """" & keyName & """:"
    pos = InStr(1, json, searchKey)
    If pos = 0 Then
        ExtractJsonValue = ""
        Exit Function
    End If

    ' Move past the key and colon
    startPos = pos + Len(searchKey)

    ' Skip whitespace
    Do While startPos <= Len(json)
        ch = Mid(json, startPos, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        startPos = startPos + 1
    Loop

    If startPos > Len(json) Then
        ExtractJsonValue = ""
        Exit Function
    End If

    ch = Mid(json, startPos, 1)

    ' null
    If Mid(json, startPos, 4) = "null" Then
        ExtractJsonValue = ""
        Exit Function
    End If

    ' String value
    If ch = """" Then
        endPos = startPos + 1
        Do While endPos <= Len(json)
            ch = Mid(json, endPos, 1)
            If ch = "\" Then
                endPos = endPos + 2
            ElseIf ch = """" Then
                ExtractJsonValue = Mid(json, startPos + 1, endPos - startPos - 1)
                ExtractJsonValue = Replace(ExtractJsonValue, "\""", """")
                ExtractJsonValue = Replace(ExtractJsonValue, "\\", "\")
                ExtractJsonValue = Replace(ExtractJsonValue, "\t", vbTab)
                Exit Function
            Else
                endPos = endPos + 1
            End If
        Loop
        ExtractJsonValue = ""
        Exit Function
    End If

    ' Array or Object - bracket match
    If ch = "[" Or ch = "{" Then
        depth = 1
        endPos = startPos + 1
        inString = False
        Do While endPos <= Len(json) And depth > 0
            ch = Mid(json, endPos, 1)
            If inString Then
                If ch = "\" Then
                    endPos = endPos + 1
                ElseIf ch = """" Then
                    inString = False
                End If
            Else
                If ch = """" Then
                    inString = True
                ElseIf ch = "[" Or ch = "{" Then
                    depth = depth + 1
                ElseIf ch = "]" Or ch = "}" Then
                    depth = depth - 1
                End If
            End If
            endPos = endPos + 1
        Loop
        ExtractJsonValue = Mid(json, startPos, endPos - startPos)
        Exit Function
    End If

    ' Numeric value
    endPos = startPos
    Do While endPos <= Len(json)
        ch = Mid(json, endPos, 1)
        If ch = "," Or ch = "}" Or ch = "]" Or ch = " " Or ch = vbCr Or ch = vbLf Then Exit Do
        endPos = endPos + 1
    Loop
    ExtractJsonValue = Trim(Mid(json, startPos, endPos - startPos))
End Function


' =============================================================================
' ParseJsonToCells - extracts all key-value pairs from a JSON response and
' writes them to the sheet. Dynamically creates new column headers as new
' keys are encountered. nextCol tracks the next available column.
' =============================================================================
Private Sub ParseJsonToCells(ws As Worksheet, row As Long, _
                              ByVal json As String, ByRef nextCol As Long)
    Dim keys() As String
    Dim vals() As String
    Dim kvCount As Long
    Dim k As Long
    Dim colIdx As Long

    ExtractAllJsonKeyValues json, keys, vals, kvCount

    For k = 1 To kvCount
        colIdx = FindOrCreateColumn(ws, keys(k), nextCol)
        ws.Cells(row, colIdx).Value = vals(k)
    Next k
End Sub


' =============================================================================
' FindOrCreateColumn - looks for an existing header in row 1 of the sheet.
' If not found, creates it at nextCol and increments nextCol.
' =============================================================================
Private Function FindOrCreateColumn(ws As Worksheet, ByVal key As String, _
                                     ByRef nextCol As Long) As Long
    Dim c As Long
    For c = 1 To nextCol - 1
        If ws.Cells(1, c).Value = key Then
            FindOrCreateColumn = c
            Exit Function
        End If
    Next c
    ' Not found - add new column header
    ws.Cells(1, nextCol).Value = key
    FindOrCreateColumn = nextCol
    nextCol = nextCol + 1
End Function


' =============================================================================
' ExtractAllJsonKeyValues - wrapper that recursively flattens a JSON object
' into parallel key/value arrays. Nested objects use "parent only" dot
' notation (e.g. "ResponseStatus_AM21.TransactionResponseStatus_AN").
' Arrays are transparent - their elements are flattened under the array key.
' Empty arrays/objects produce a header with an empty-string value.
' =============================================================================
Private Sub ExtractAllJsonKeyValues(ByVal json As String, _
                                     ByRef keys() As String, _
                                     ByRef vals() As String, _
                                     ByRef count As Long)
    count = 0
    ReDim keys(1 To 100)
    ReDim vals(1 To 100)
    FlattenJsonObject json, "", keys, vals, count
End Sub


' =============================================================================
' FlattenJsonObject - parses a JSON object string and flattens all nested
' values into the keys/vals arrays. prefix is the current "parent only"
' context key (reset at each nesting level).
' =============================================================================
Private Sub FlattenJsonObject(ByVal json As String, _
                               ByVal prefix As String, _
                               ByRef keys() As String, _
                               ByRef vals() As String, _
                               ByRef count As Long)
    Dim pos As Long
    Dim length As Long
    Dim ch As String

    length = Len(json)
    pos = 1

    ' Skip to first {
    Do While pos <= length
        If Mid(json, pos, 1) = "{" Then
            pos = pos + 1
            Exit Do
        End If
        pos = pos + 1
    Loop

    ' Check for empty object {}
    Dim tmpPos As Long: tmpPos = pos
    Do While tmpPos <= length
        ch = Mid(json, tmpPos, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        tmpPos = tmpPos + 1
    Loop
    If tmpPos <= length And Mid(json, tmpPos, 1) = "}" Then
        ' Empty object - store header with empty value
        If Len(prefix) > 0 Then
            count = count + 1
            If count > UBound(keys) Then
                ReDim Preserve keys(1 To count + 100)
                ReDim Preserve vals(1 To count + 100)
            End If
            keys(count) = prefix
            vals(count) = ""
        End If
        Exit Sub
    End If

    Do While pos <= length
        ' Skip whitespace and commas
        Do While pos <= length
            ch = Mid(json, pos, 1)
            If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf And ch <> "," Then Exit Do
            pos = pos + 1
        Loop

        If pos > length Then Exit Do
        ch = Mid(json, pos, 1)
        If ch = "}" Then Exit Do

        ' Expect a key string starting with "
        If ch <> """" Then
            pos = pos + 1
            GoTo FlatObjContinue
        End If

        ' Extract key
        pos = pos + 1
        Dim keyStart As Long: keyStart = pos
        Do While pos <= length
            ch = Mid(json, pos, 1)
            If ch = "\" Then
                pos = pos + 2
            ElseIf ch = """" Then
                Exit Do
            Else
                pos = pos + 1
            End If
        Loop
        Dim key As String
        key = Mid(json, keyStart, pos - keyStart)
        pos = pos + 1  ' skip closing "

        ' Build the full header for this key
        Dim fullKey As String
        If Len(prefix) > 0 Then
            fullKey = prefix & "." & key
        Else
            fullKey = key
        End If

        ' Skip whitespace and colon
        Do While pos <= length
            ch = Mid(json, pos, 1)
            If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf And ch <> ":" Then Exit Do
            pos = pos + 1
        Loop

        If pos > length Then GoTo FlatObjContinue

        ch = Mid(json, pos, 1)

        If ch = """" Then
            ' String value - store as leaf
            pos = pos + 1
            Dim valStart As Long: valStart = pos
            Do While pos <= length
                ch = Mid(json, pos, 1)
                If ch = "\" Then
                    pos = pos + 2
                ElseIf ch = """" Then
                    Exit Do
                Else
                    pos = pos + 1
                End If
            Loop
            Dim sVal As String
            sVal = Mid(json, valStart, pos - valStart)
            sVal = Replace(sVal, "\""", """")
            sVal = Replace(sVal, "\\", "\")
            pos = pos + 1  ' skip closing "
            count = count + 1
            If count > UBound(keys) Then
                ReDim Preserve keys(1 To count + 100)
                ReDim Preserve vals(1 To count + 100)
            End If
            keys(count) = fullKey
            vals(count) = sVal

        ElseIf ch = "{" Then
            ' Nested object - extract raw JSON then recurse
            Dim objDepth As Long: objDepth = 1
            Dim objStart As Long: objStart = pos
            Dim objInStr As Boolean: objInStr = False
            pos = pos + 1
            Do While pos <= length And objDepth > 0
                ch = Mid(json, pos, 1)
                If objInStr Then
                    If ch = "\" Then
                        pos = pos + 1
                    ElseIf ch = """" Then
                        objInStr = False
                    End If
                Else
                    If ch = """" Then
                        objInStr = True
                    ElseIf ch = "{" Then
                        objDepth = objDepth + 1
                    ElseIf ch = "}" Then
                        objDepth = objDepth - 1
                    End If
                End If
                pos = pos + 1
            Loop
            Dim objJson As String
            objJson = Mid(json, objStart, pos - objStart)
            ' Recurse with key as the new prefix (parent-only naming)
            FlattenJsonObject objJson, key, keys, vals, count

        ElseIf ch = "[" Then
            ' Array - extract raw JSON then recurse
            Dim arrDepth As Long: arrDepth = 1
            Dim arrStart As Long: arrStart = pos
            Dim arrInStr As Boolean: arrInStr = False
            pos = pos + 1
            Do While pos <= length And arrDepth > 0
                ch = Mid(json, pos, 1)
                If arrInStr Then
                    If ch = "\" Then
                        pos = pos + 1
                    ElseIf ch = """" Then
                        arrInStr = False
                    End If
                Else
                    If ch = """" Then
                        arrInStr = True
                    ElseIf ch = "[" Then
                        arrDepth = arrDepth + 1
                    ElseIf ch = "]" Then
                        arrDepth = arrDepth - 1
                    End If
                End If
                pos = pos + 1
            Loop
            Dim arrJson As String
            arrJson = Mid(json, arrStart, pos - arrStart)
            ' Recurse into array with key as prefix
            FlattenJsonArray arrJson, key, keys, vals, count

        ElseIf Mid(json, pos, 4) = "null" Then
            pos = pos + 4
            count = count + 1
            If count > UBound(keys) Then
                ReDim Preserve keys(1 To count + 100)
                ReDim Preserve vals(1 To count + 100)
            End If
            keys(count) = fullKey
            vals(count) = ""

        ElseIf Mid(json, pos, 4) = "true" Then
            pos = pos + 4
            count = count + 1
            If count > UBound(keys) Then
                ReDim Preserve keys(1 To count + 100)
                ReDim Preserve vals(1 To count + 100)
            End If
            keys(count) = fullKey
            vals(count) = "true"

        ElseIf Mid(json, pos, 5) = "false" Then
            pos = pos + 5
            count = count + 1
            If count > UBound(keys) Then
                ReDim Preserve keys(1 To count + 100)
                ReDim Preserve vals(1 To count + 100)
            End If
            keys(count) = fullKey
            vals(count) = "false"

        Else
            ' Numeric value
            Dim numStart As Long: numStart = pos
            Do While pos <= length
                ch = Mid(json, pos, 1)
                If ch = "," Or ch = "}" Or ch = "]" Or ch = " " Or ch = vbCr Or ch = vbLf Then Exit Do
                pos = pos + 1
            Loop
            count = count + 1
            If count > UBound(keys) Then
                ReDim Preserve keys(1 To count + 100)
                ReDim Preserve vals(1 To count + 100)
            End If
            keys(count) = fullKey
            vals(count) = Trim(Mid(json, numStart, pos - numStart))
        End If

FlatObjContinue:
    Loop
End Sub


' =============================================================================
' FlattenJsonArray - parses a JSON array string and flattens each element.
' Arrays are transparent: object elements are flattened with prefix as the
' parent key. Primitive elements use prefix as the header directly.
' Empty arrays produce a single header with an empty-string value.
' =============================================================================
Private Sub FlattenJsonArray(ByVal json As String, _
                              ByVal prefix As String, _
                              ByRef keys() As String, _
                              ByRef vals() As String, _
                              ByRef count As Long)
    Dim pos As Long
    Dim length As Long
    Dim ch As String

    length = Len(json)
    pos = 1

    ' Skip to first [
    Do While pos <= length
        If Mid(json, pos, 1) = "[" Then
            pos = pos + 1
            Exit Do
        End If
        pos = pos + 1
    Loop

    ' Check for empty array []
    Dim tmpPos As Long: tmpPos = pos
    Do While tmpPos <= length
        ch = Mid(json, tmpPos, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        tmpPos = tmpPos + 1
    Loop
    If tmpPos <= length And Mid(json, tmpPos, 1) = "]" Then
        ' Empty array - store header with empty value
        If Len(prefix) > 0 Then
            count = count + 1
            If count > UBound(keys) Then
                ReDim Preserve keys(1 To count + 100)
                ReDim Preserve vals(1 To count + 100)
            End If
            keys(count) = prefix
            vals(count) = ""
        End If
        Exit Sub
    End If

    ' Iterate array elements
    Do While pos <= length
        ' Skip whitespace and commas
        Do While pos <= length
            ch = Mid(json, pos, 1)
            If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf And ch <> "," Then Exit Do
            pos = pos + 1
        Loop

        If pos > length Then Exit Do
        ch = Mid(json, pos, 1)
        If ch = "]" Then Exit Do

        If ch = "{" Then
            ' Object element - extract and recurse
            Dim eDepth As Long: eDepth = 1
            Dim eStart As Long: eStart = pos
            Dim eInStr As Boolean: eInStr = False
            pos = pos + 1
            Do While pos <= length And eDepth > 0
                ch = Mid(json, pos, 1)
                If eInStr Then
                    If ch = "\" Then
                        pos = pos + 1
                    ElseIf ch = """" Then
                        eInStr = False
                    End If
                Else
                    If ch = """" Then
                        eInStr = True
                    ElseIf ch = "{" Then
                        eDepth = eDepth + 1
                    ElseIf ch = "}" Then
                        eDepth = eDepth - 1
                    End If
                End If
                pos = pos + 1
            Loop
            Dim elemJson As String
            elemJson = Mid(json, eStart, pos - eStart)
            ' Recurse with same prefix (arrays are transparent)
            FlattenJsonObject elemJson, prefix, keys, vals, count

        ElseIf ch = "[" Then
            ' Nested array - extract and recurse
            Dim naDepth As Long: naDepth = 1
            Dim naStart As Long: naStart = pos
            Dim naInStr As Boolean: naInStr = False
            pos = pos + 1
            Do While pos <= length And naDepth > 0
                ch = Mid(json, pos, 1)
                If naInStr Then
                    If ch = "\" Then
                        pos = pos + 1
                    ElseIf ch = """" Then
                        naInStr = False
                    End If
                Else
                    If ch = """" Then
                        naInStr = True
                    ElseIf ch = "[" Then
                        naDepth = naDepth + 1
                    ElseIf ch = "]" Then
                        naDepth = naDepth - 1
                    End If
                End If
                pos = pos + 1
            Loop
            Dim naJson As String
            naJson = Mid(json, naStart, pos - naStart)
            FlattenJsonArray naJson, prefix, keys, vals, count

        ElseIf ch = """" Then
            ' Primitive string element in array
            pos = pos + 1
            Dim psStart As Long: psStart = pos
            Do While pos <= length
                ch = Mid(json, pos, 1)
                If ch = "\" Then
                    pos = pos + 2
                ElseIf ch = """" Then
                    Exit Do
                Else
                    pos = pos + 1
                End If
            Loop
            Dim psVal As String
            psVal = Mid(json, psStart, pos - psStart)
            psVal = Replace(psVal, "\""", """")
            psVal = Replace(psVal, "\\", "\")
            pos = pos + 1
            count = count + 1
            If count > UBound(keys) Then
                ReDim Preserve keys(1 To count + 100)
                ReDim Preserve vals(1 To count + 100)
            End If
            keys(count) = prefix
            vals(count) = psVal

        Else
            ' Primitive numeric/boolean/null element
            Dim pStart As Long: pStart = pos
            If Mid(json, pos, 4) = "null" Then
                pos = pos + 4
                count = count + 1
                If count > UBound(keys) Then
                    ReDim Preserve keys(1 To count + 100)
                    ReDim Preserve vals(1 To count + 100)
                End If
                keys(count) = prefix
                vals(count) = ""
            ElseIf Mid(json, pos, 4) = "true" Then
                pos = pos + 4
                count = count + 1
                If count > UBound(keys) Then
                    ReDim Preserve keys(1 To count + 100)
                    ReDim Preserve vals(1 To count + 100)
                End If
                keys(count) = prefix
                vals(count) = "true"
            ElseIf Mid(json, pos, 5) = "false" Then
                pos = pos + 5
                count = count + 1
                If count > UBound(keys) Then
                    ReDim Preserve keys(1 To count + 100)
                    ReDim Preserve vals(1 To count + 100)
                End If
                keys(count) = prefix
                vals(count) = "false"
            Else
                ' Number
                Do While pos <= length
                    ch = Mid(json, pos, 1)
                    If ch = "," Or ch = "]" Or ch = " " Or ch = vbCr Or ch = vbLf Then Exit Do
                    pos = pos + 1
                Loop
                count = count + 1
                If count > UBound(keys) Then
                    ReDim Preserve keys(1 To count + 100)
                    ReDim Preserve vals(1 To count + 100)
                End If
                keys(count) = prefix
                vals(count) = Trim(Mid(json, pStart, pos - pStart))
            End If
        End If
    Loop
End Sub


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