Attribute VB_Name = "Test"
Option Explicit

Dim TWhitespaceTest As String
Dim TWhitespaceResults As String
Dim ConvertHTMLResults As String
Dim ConvertHTMLTests As String
Dim IsQuotedTest As String
Dim IsQuotedResults As String
Dim ConvertVariableExpressionTest As String
Dim ConvertVariableExpressionResults As String
Dim MissingLValueExpressionsResults As String
Dim ValidateVariableIDTest As String
Dim ValidateVariableIDResults As String
Dim UnQuotedDelimiterIndexTest As String
Dim UnQuotedDelimiterIndexResults As String
Dim UnQuotedDelimiterIndexREVTest As String
Dim UnQuotedDelimiterIndexREVResults As String
Dim GetNextNonBlankFieldIndexTest As String
Dim GetNextNonBlankFieldIndexResults As String
Dim GetNonBlankFieldTest As String
Dim GetNonBlankFieldResults As String
Dim GetFieldCountTest As String
Dim GetFieldCountResults As String
Dim GetTrimmedFieldTest As String
Dim GetTrimmedFieldResults As String
Dim GetLiteralQuotedFieldTest As String
Dim GetLiteralQuotedFieldResults As String
Dim GetLiteralQuotedRecordFieldTest As String
Dim GetLiteralQuotedRecordFieldResults As String
Const RULER = ";=== ====1==== ====2==== ====3==== ====4==== ====5==== ====6==== ====7==== ====8"



'returns PASS(True) or FAIL(False)
Function LogStringCompare(testIn As String, bench As String, actual As String, fileref As Integer) As Boolean
    
    Dim prefix As String
    
    If StrComp(bench, actual, vbBinaryCompare) <> 0 Then
        prefix = "FAIL:"
        LogStringCompare = False
    Else
        prefix = " OK :"
        LogStringCompare = True
    End If
    
    Print #fileref, ""
    Print #fileref, "     " & "INPUT:" & testIn & ":"
    Print #fileref, "     " & "BENCH:" & bench & ":"
    Print #fileref, prefix & "RTURN:" & actual & ":"

End Function

Function GetNextLine(fileref As Integer) As String
    Dim record As String
    
    Do While Not EOF(fileref)
        Line Input #fileref, record
        If Len(record) < 1 Then GoTo GNLLoop
        If Left$(record, 1) = ";" Then GoTo GNLLoop
        GetNextLine = record
        Exit Do
GNLLoop:
    Loop
End Function

Function Test_TWhitespace() As Long
    Dim su As StringUtilities
    Dim infileref As Integer
    Dim outfileref As Integer
    Dim testIn As String
    Dim bench As String
    Dim actual As String
    Dim record As String
    Dim index As Long
    Dim passed As Boolean
    Dim errors As Long
    
    Set su = New DDVariableStore.StringUtilities
    
    infileref = FreeFile
    Open TWhitespaceTest For Input Access Read Shared As infileref
    outfileref = FreeFile
    Open TWhitespaceResults For Output As outfileref
    
    Do Until EOF(infileref)
        record = GetNextLine(infileref)
        
        If Len(record) = 0 Then
            Print #outfileref, "     Last TWhitespace record processed."
            Exit Do
        End If
        
        index = InStr(record, ",")
        If index < 1 Then
            Print #outfileref, "FAIL:Not a valid TWhitespace record."
            Exit Do
        End If
        
        testIn = Left$(record, index - 1)
        bench = Mid$(record, index + 1)
        actual = su.TWhitespace(testIn)
        
        passed = LogStringCompare(testIn, bench, actual, outfileref)
        If Not passed Then errors = errors + 1
        
    Loop
    
    Test_TWhitespace = errors
    
    Close #infileref
    Close #outfileref
    Set su = Nothing
    
End Function

Function Test_ConvertHTMLEntities() As Long
    Dim su As StringUtilities
    Dim infileref As Integer
    Dim outfileref As Integer
    Dim testIn As String
    Dim bench As String
    Dim actual As String
    Dim record As String
    Dim index As Long
    Dim passed As Boolean
    Dim errors As Long
    
    Set su = New DDVariableStore.StringUtilities
    Dim var
    
    infileref = FreeFile
    Open ConvertHTMLTests For Input Access Read Shared As infileref
    outfileref = FreeFile
    Open ConvertHTMLResults For Output As outfileref
    
    Do Until EOF(infileref)
        record = GetNextLine(infileref)
        
        If Len(record) = 0 Then
            Print #outfileref, "     Last ConvertHTML record processed."
            Exit Do
        End If
        
        index = InStr(record, ",")
        If index < 1 Then
            Print #outfileref, "FAIL:Not a valid ConvertHTML record."
            Exit Do
        End If
        
        testIn = Left$(record, index - 1)
        bench = Mid$(record, index + 1)
        actual = su.ConvertHTMLEntities(testIn)
        
        passed = LogStringCompare(testIn, bench, actual, outfileref)
        If Not passed Then errors = errors + 1
        
    Loop
    
    Test_ConvertHTMLEntities = errors
    
    Close #infileref
    Close #outfileref
    Set su = Nothing
    
End Function

Function Test_IsQuoted() As Long
    Dim su As StringUtilities
    Dim infileref As Integer
    Dim outfileref As Integer
    Dim prefix As String
    Dim testIn As String
    Dim indexIn As Long
    Dim bench As Boolean
    Dim actual As Boolean
    Dim record As String
    Dim index As Long
    Dim index2 As Long
    Dim passed As Boolean
    Dim errors As Long
    
    Set su = New DDVariableStore.StringUtilities
    
    infileref = FreeFile
    Open IsQuotedTest For Input Access Read Shared As infileref
    outfileref = FreeFile
    Open IsQuotedResults For Output As outfileref
    
    Do Until EOF(infileref)
        record = GetNextLine(infileref)
        
        If Len(record) = 0 Then
            Print #outfileref, "     Last IsQuoted record processed."
            Exit Do
        End If
        
        index = InStr(record, ",")
        If index < 1 Then
            Print #outfileref, "FAIL:Not a valid IsQuoted record."
            Exit Do
        End If
        
        testIn = Left$(record, index - 1)
        
        index2 = InStr(index + 1, record, ",")
        
        If index2 < 1 Then
            Print #outfileref, "FAIL:Not a valid IsQuoted record."
            Exit Do
        End If
        
        indexIn = CLng(Mid$(record, index + 1, index2 - index - 1))
        If indexIn < 1 Then
            Print #outfileref, "FAIL:Not a valid IsQuoted record."
            Exit Do
        End If
        
        bench = CBool(Mid$(record, index2 + 1))
        actual = su.IsQuoted(testIn, indexIn)
        
        If bench <> actual Then
            prefix = "FAIL:"
            passed = False
        Else
            prefix = " OK :"
            passed = True
        End If
        
        Print #outfileref, "           " & RULER
        Print #outfileref, "     " & "INPUT:" & testIn & ":INDEX:" & Str$(indexIn)
        Print #outfileref, "     " & "BENCH:" & bench & ":"
        Print #outfileref, prefix & "RTURN:" & actual & ":"
        
        If Not passed Then errors = errors + 1
        
    Loop
    
    Test_IsQuoted = errors
    
    Close #infileref
    Close #outfileref
    Set su = Nothing
    
End Function


Function Test_ValidateVariableID() As Long
    Dim su As StringUtilities
    Dim infileref As Integer
    Dim outfileref As Integer
    Dim testIn As String
    Dim bench As String
    Dim actual As String
    Dim record As String
    Dim index As Long
    Dim passed As Boolean
    Dim errors As Long
    Dim start As Single, finish As Single
    
    Set su = New DDVariableStore.StringUtilities
    
    infileref = FreeFile
    Open ValidateVariableIDTest For Input Access Read Shared As infileref
    outfileref = FreeFile
    Open ValidateVariableIDResults For Output As outfileref
    
    start = Timer
    Print #outfileref, "Begin Time:" & Str$(start)
    Do Until EOF(infileref)
        record = GetNextLine(infileref)
        
        If Len(record) = 0 Then
            Print #outfileref, "     Last ValidateVariableID record processed."
            Exit Do
        End If
        
        index = InStr(record, ",")
        If index < 1 Then
            Print #outfileref, "FAIL:Not a valid ValidateVariableID record."
            Exit Do
        End If
        
        testIn = Left$(record, index - 1)
        bench = Mid$(record, index + 1)
        actual = su.ValidateVariableID(testIn)
        
        passed = LogStringCompare(testIn, bench, actual, outfileref)
        If Not passed Then errors = errors + 1
        
    Loop
    
    finish = Timer
    Print #outfileref, "End Time:" & Str$(finish)
    Print #outfileref, ""
    Print #outfileref, "Elapsed Time:" & Str$(finish - start)
    
    Test_ValidateVariableID = errors
    
    Close #infileref
    Close #outfileref
    Set su = Nothing
    
End Function

Function Test_UnQuotedDelimiterIndex() As Long
    Dim su As StringUtilities
    Dim infileref As Integer
    Dim outfileref As Integer
    Dim prefix As String
    Dim testIn As String
    Dim indexIn As Long
    Dim delimitersIn As String
    Dim bench As Long
    Dim actual As Long
    Dim record As String
    Dim index As Long
    Dim index2 As Long
    Dim index3 As Long
    Dim passed As Boolean
    Dim errors As Long
    
    Set su = New DDVariableStore.StringUtilities
    
    infileref = FreeFile
    Open UnQuotedDelimiterIndexTest For Input Access Read Shared As infileref
    outfileref = FreeFile
    Open UnQuotedDelimiterIndexResults For Output As outfileref
    
    Do Until EOF(infileref)
        record = GetNextLine(infileref)
        
        If Len(record) = 0 Then
            Print #outfileref, "     Last UnquotedDelimiterIndex record processed."
            Exit Do
        End If
        
        index = InStr(record, ",")
        If index < 1 Then
            Print #outfileref, "FAIL:Not a valid UnquotedDelimiterIndex record."
            Exit Do
        End If
        
        testIn = Left$(record, index - 1)
        
        index2 = InStr(index + 1, record, ",")
        
        If index2 < 1 Then
            Print #outfileref, "FAIL:Not a valid UnquotedDelimiterIndex record."
            Exit Do
        End If
        
        delimitersIn = Mid$(record, index + 1, index2 - index - 1)
        
        index3 = InStr(index2 + 1, record, ",")
        
        If index3 < 1 Then
            Print #outfileref, "FAIL:Not a valid UnquotedDelimiterIndex record."
            Exit Do
        End If
        
        indexIn = CLng(Mid$(record, index2 + 1, index3 - index2 - 1))
'        If indexIn < 1 Then
'            Print #outfileref, "FAIL:Not a valid UnquotedDelimiterIndex record."
'            Exit Do
'        End If
        
        bench = CLng(Mid$(record, index3 + 1))
        actual = su.GetUnquotedDelimiterIndex(indexIn, testIn, delimitersIn)
        
        If bench <> actual Then
            prefix = "FAIL:"
            passed = False
        Else
            prefix = " OK :"
            passed = True
        End If
        
        Print #outfileref, "           " & RULER
        Print #outfileref, "     " & "INPUT:" & testIn & ":INDEX:" & Str$(indexIn)
        Print #outfileref, "     " & "DELIM:" & delimitersIn
        Print #outfileref, "     " & "BENCH:" & bench
        Print #outfileref, prefix & "RTURN:" & actual
        
        If Not passed Then errors = errors + 1
        
    Loop
    
    Test_UnQuotedDelimiterIndex = errors
    
    Close #infileref
    Close #outfileref
    Set su = Nothing
    
End Function


Function Test_UnQuotedDelimiterIndexREV() As Long
    Dim su As StringUtilities
    Dim infileref As Integer
    Dim outfileref As Integer
    Dim prefix As String
    Dim testIn As String
    Dim indexIn As Long
    Dim delimitersIn As String
    Dim bench As Long
    Dim actual As Long
    Dim record As String
    Dim index As Long
    Dim index2 As Long
    Dim index3 As Long
    Dim passed As Boolean
    Dim errors As Long
    
    Set su = New DDVariableStore.StringUtilities
    
    infileref = FreeFile
    Open UnQuotedDelimiterIndexREVTest For Input Access Read Shared As infileref
    outfileref = FreeFile
    Open UnQuotedDelimiterIndexREVResults For Output As outfileref
    
    Do Until EOF(infileref)
        record = GetNextLine(infileref)
        
        If Len(record) = 0 Then
            Print #outfileref, "     Last UnquotedDelimiterIndexREV record processed."
            Exit Do
        End If
        
        index = InStr(record, ",")
        If index < 1 Then
            Print #outfileref, "FAIL:Not a valid UnquotedDelimiterIndexREV record."
            Exit Do
        End If
        
        testIn = Left$(record, index - 1)
        
        index2 = InStr(index + 1, record, ",")
        
        If index2 < 1 Then
            Print #outfileref, "FAIL:Not a valid UnquotedDelimiterIndexREV record."
            Exit Do
        End If
        
        delimitersIn = Mid$(record, index + 1, index2 - index - 1)
        
        index3 = InStr(index2 + 1, record, ",")
        
        If index3 < 1 Then
            Print #outfileref, "FAIL:Not a valid UnquotedDelimiterIndexREV record."
            Exit Do
        End If
        
        indexIn = CLng(Mid$(record, index2 + 1, index3 - index2 - 1))
        If indexIn < 1 Then
            Print #outfileref, "FAIL:Not a valid UnquotedDelimiterIndexREV record."
            Exit Do
        End If
        
        bench = CLng(Mid$(record, index3 + 1))
        actual = su.GetUnquotedDelimiterIndexRev(indexIn, testIn, delimitersIn)
        
        If bench <> actual Then
            prefix = "FAIL:"
            passed = False
        Else
            prefix = " OK :"
            passed = True
        End If
        
        Print #outfileref, "           " & RULER
        Print #outfileref, "     " & "INPUT:" & testIn & ":INDEX:" & Str$(indexIn)
        Print #outfileref, "     " & "DELIM:" & delimitersIn
        Print #outfileref, "     " & "BENCH:" & bench
        Print #outfileref, prefix & "RTURN:" & actual
        
        If Not passed Then errors = errors + 1
        
    Loop
    
    Test_UnQuotedDelimiterIndexREV = errors
    
    Close #infileref
    Close #outfileref
    Set su = Nothing
    
End Function


Function Test_ConvertVariableExpression() As Long
    Dim su As StringUtilities
    Dim infileref As Integer
    Dim outfileref As Integer
    Dim prefix As String
    Dim strPrefix As String
    Dim expressionIn As String
    Dim valueIn As String
    Dim benchStr As String
    Dim benchInt As Integer
    Dim actualStr As String
    Dim actualInt As Integer
    Dim record As String
    Dim index As Long
    Dim index2 As Long
    Dim index3 As Long
    Dim passed As Boolean
    Dim errors As Long
    Dim store As DDVariableStore.GlobalVariables
    
    Set su = New DDVariableStore.StringUtilities
    Set store = New DDVariableStore.GlobalVariables
    
    infileref = FreeFile
    Open ConvertVariableExpressionTest For Input Access Read Shared As infileref
    outfileref = FreeFile
    Open ConvertVariableExpressionResults For Output As outfileref
    
    Do Until EOF(infileref)
        record = GetNextLine(infileref)
        
        If Len(record) = 0 Then
            Print #outfileref, "     Last ConvertVariableExpression record processed."
            Exit Do
        End If
        
        index = InStr(record, ",")
        If index < 1 Then
            Print #outfileref, "FAIL:Not a valid ConvertVariableExpression record."
            Exit Do
        End If
        
        expressionIn = Left$(record, index - 1)
        
        index2 = InStr(index + 1, record, ",")
        
        If index2 < 1 Then
            Print #outfileref, "FAIL:Not a valid ConvertVariableExpression record."
            Exit Do
        End If
        
        valueIn = Mid$(record, index + 1, index2 - index - 1)
        
        index3 = InStr(index2 + 1, record, ",")
        
        If index3 < 1 Then
            Print #outfileref, "FAIL:Not a valid ConvertVariableExpression record."
            Exit Do
        End If
        
        benchStr = Mid$(record, index2 + 1, index3 - index2 - 1)
        
        benchInt = CInt(Mid$(record, index3 + 1))
        actualInt = su.ConvertVariableExpression(expressionIn, store, valueIn)
        
        passed = True
        
        If (benchInt <> actualInt) Then
            prefix = "FAIL:"
            passed = False
        Else
            prefix = " OK :"
        End If
        
        If (benchStr <> valueIn) Then
            strPrefix = "FAIL:"
            passed = False
        Else
            strPrefix = " OK :"
        End If
        
        Print #outfileref, "           " & RULER
        Print #outfileref, "     " & "INPUT:" & expressionIn & ":"
        Print #outfileref, "     " & "BENCHSTR:" & benchStr & ":"
        Print #outfileref, strPrefix & "ACTULSTR:" & valueIn & ":"
        Print #outfileref, "     " & "BENCHINT:" & Str$(benchInt)
        Print #outfileref, prefix & "ACTULINT:" & Str$(actualInt)
        
        If Not passed Then errors = errors + 1
        
    Loop
    
    Test_ConvertVariableExpression = errors
    
    Close #infileref
    Close #outfileref
    Set su = Nothing
    
End Function


Function Test_GetNextNonBlankFieldIndex() As Long
    Dim su As StringUtilities
    Dim infileref As Integer
    Dim outfileref As Integer
    Dim prefix As String
    Dim testIn As String
    Dim indexIn As Long
    Dim delimitersIn As String
    Dim bench As Long
    Dim actual As Long
    Dim record As String
    Dim index As Long
    Dim index2 As Long
    Dim index3 As Long
    Dim passed As Boolean
    Dim errors As Long
    
    Set su = New DDVariableStore.StringUtilities
    
    infileref = FreeFile
    Open GetNextNonBlankFieldIndexTest For Input Access Read Shared As infileref
    outfileref = FreeFile
    Open GetNextNonBlankFieldIndexResults For Output As outfileref
    
    Do Until EOF(infileref)
        record = GetNextLine(infileref)
        
        If Len(record) = 0 Then
            Print #outfileref, "     Last GetNextNonBlankFieldIndex record processed."
            Exit Do
        End If
        
        index = InStr(record, ",")
        If index < 1 Then
            Print #outfileref, "FAIL:Not a valid GetNextNonBlankFieldIndex record."
            Exit Do
        End If
        
        testIn = Left$(record, index - 1)
        
        index2 = InStr(index + 1, record, ",")
        
        If index2 < 1 Then
            Print #outfileref, "FAIL:Not a valid GetNextNonBlankFieldIndex record."
            Exit Do
        End If
        
        delimitersIn = Mid$(record, index + 1, index2 - index - 1)
        
        index3 = InStr(index2 + 1, record, ",")
        
        If index3 < 1 Then
            Print #outfileref, "FAIL:Not a valid GetNextNonBlankFieldIndex record."
            Exit Do
        End If
        
        indexIn = CLng(Mid$(record, index2 + 1, index3 - index2 - 1))
'        If indexIn < 1 Then
'            Print #outfileref, "FAIL:Not a valid GetNextNonBlankFieldIndex record."
'            Exit Do
'        End If
        
        bench = CLng(Mid$(record, index3 + 1))
        actual = su.GetNextNonBlankFieldIndex(indexIn, testIn, delimitersIn)
        
        If bench <> actual Then
            prefix = "FAIL:"
            passed = False
        Else
            prefix = " OK :"
            passed = True
        End If
        
        Print #outfileref, "           " & RULER
        Print #outfileref, "     " & "INPUT:" & testIn & ":INDEX:" & Str$(indexIn)
        Print #outfileref, "     " & "DELIM:" & delimitersIn
        Print #outfileref, "     " & "BENCH:" & bench
        Print #outfileref, prefix & "RTURN:" & actual
        
        If Not passed Then errors = errors + 1
        
    Loop
    
    Test_GetNextNonBlankFieldIndex = errors
    
    Close #infileref
    Close #outfileref
    Set su = Nothing
    
End Function


Function Test_GetFieldCount() As Long
    Dim su As StringUtilities
    Dim infileref As Integer
    Dim outfileref As Integer
    Dim prefix As String
    Dim testIn As String
    Dim indexIn As Long
    Dim delimitersIn As String
    Dim bench As Integer
    Dim actual As Integer
    Dim record As String
    Dim index As Long
    Dim index2 As Long
    Dim index3 As Long
    Dim passed As Boolean
    Dim errors As Long
    
    Set su = New DDVariableStore.StringUtilities
    
    infileref = FreeFile
    Open GetFieldCountTest For Input Access Read Shared As infileref
    outfileref = FreeFile
    Open GetFieldCountResults For Output As outfileref
    
    Do Until EOF(infileref)
        record = GetNextLine(infileref)
        
        If Len(record) = 0 Then
            Print #outfileref, "     Last GetFieldCount record processed."
            Exit Do
        End If
        
        index = InStr(record, ",")
        If index < 1 Then
            Print #outfileref, "FAIL:Not a valid GetFieldCount record."
            Exit Do
        End If
        
        testIn = Left$(record, index - 1)
        
        index2 = InStr(index + 1, record, ",")
        
        If index2 < 1 Then
            Print #outfileref, "FAIL:Not a valid GetFieldCount record."
            Exit Do
        End If
        
        delimitersIn = Mid$(record, index + 1, index2 - index - 1)
        
        index3 = InStr(index2 + 1, record, ",")
        
        If index3 < 1 Then
            Print #outfileref, "FAIL:Not a valid GetFieldCount record."
            Exit Do
        End If
        
        indexIn = CLng(Mid$(record, index2 + 1, index3 - index2 - 1))
'        If indexIn < 1 Then
'            Print #outfileref, "FAIL:Not a valid GetNextNonBlankFieldIndex record."
'            Exit Do
'        End If
        
        bench = CInt(Mid$(record, index3 + 1))
        actual = su.GetFieldCount(indexIn, testIn, delimitersIn)
        
        If bench <> actual Then
            prefix = "FAIL:"
            passed = False
        Else
            prefix = " OK :"
            passed = True
        End If
        
        Print #outfileref, "           " & RULER
        Print #outfileref, "     " & "INPUT:" & testIn & ":INDEX:" & Str$(indexIn)
        Print #outfileref, "     " & "DELIM:" & delimitersIn
        Print #outfileref, "     " & "BENCH:" & bench
        Print #outfileref, prefix & "RTURN:" & actual
        
        If Not passed Then errors = errors + 1
        
    Loop
    
    Test_GetFieldCount = errors
    
    Close #infileref
    Close #outfileref
    Set su = Nothing
    
End Function


Function Test_GetTrimmedField() As Long
    Dim su As StringUtilities
    Dim infileref As Integer
    Dim outfileref As Integer
    Dim prefix As String
    Dim testIn As String
    Dim indexIn As Long
    Dim delimitersIn As String
    Dim fieldIn As Integer
    Dim bench As String
    Dim actual As String
    Dim record As String
    Dim index As Long
    Dim index2 As Long
    Dim index3 As Long
    Dim index4 As Long
    Dim passed As Boolean
    Dim errors As Long
    
    Set su = New DDVariableStore.StringUtilities
    
    infileref = FreeFile
    Open GetTrimmedFieldTest For Input Access Read Shared As infileref
    outfileref = FreeFile
    Open GetTrimmedFieldResults For Output As outfileref
    
    Do Until EOF(infileref)
        record = GetNextLine(infileref)
        
        If Len(record) = 0 Then
            Print #outfileref, "     Last GetTrimmedField record processed."
            Exit Do
        End If
        
        index = InStr(record, ",")
        If index < 1 Then
            Print #outfileref, "FAIL:Not a valid GetTrimmedField record."
            Exit Do
        End If
        
        testIn = Left$(record, index - 1)
        
        index2 = InStr(index + 1, record, ",")
        
        If index2 < 1 Then
            Print #outfileref, "FAIL:Not a valid GetTrimmedField record."
            Exit Do
        End If
        
        delimitersIn = Mid$(record, index + 1, index2 - index - 1)
        
        index3 = InStr(index2 + 1, record, ",")
        
        If index3 < 1 Then
            Print #outfileref, "FAIL:Not a valid GetTrimmedField record."
            Exit Do
        End If
        
        indexIn = CLng(Mid$(record, index2 + 1, index3 - index2 - 1))
'        If indexIn < 1 Then
'            Print #outfileref, "FAIL:Not a valid GetNextNonBlankFieldIndex record."
'            Exit Do
'        End If
        
        index4 = InStr(index3 + 1, record, ",")
        
        If index4 < 1 Then
            Print #outfileref, "FAIL:Not a valid GetTrimmedField record."
            Exit Do
        End If
        
        fieldIn = CInt(Mid$(record, index3 + 1, index4 - index3 - 1))
        
        bench = Mid$(record, index4 + 1)
        actual = su.GetTrimmedField(indexIn, testIn, delimitersIn, fieldIn)
        
        If StrComp(bench, actual, 0) <> 0 Then
            prefix = "FAIL:"
            passed = False
        Else
            prefix = " OK :"
            passed = True
        End If
        
        Print #outfileref, "           " & RULER
        Print #outfileref, "     " & "INPUT:" & testIn & ":INDEX:" & Str$(indexIn)
        Print #outfileref, "     " & "DELIM:" & delimitersIn
        Print #outfileref, "     " & "FIELD:" & Str$(fieldIn)
        Print #outfileref, "     " & "BENCH:" & bench
        Print #outfileref, prefix & "RTURN:" & actual
        
        If Not passed Then errors = errors + 1
        
    Loop
    
    Test_GetTrimmedField = errors
    
    Close #infileref
    Close #outfileref
    Set su = Nothing
    
End Function



Function Test_GetLiteralQuotedRecordField() As Long
    Dim su As StringUtilities
    Dim infileref As Integer
    Dim outfileref As Integer
    Dim prefix As String
    Dim testIn As String
    Dim indexIn As Integer
    Dim bench As String
    Dim actual As String
    Dim record As String
    Dim index As Long
    Dim index2 As Long
    Dim passed As Boolean
    Dim errors As Long
    
    Set su = New DDVariableStore.StringUtilities
    
    infileref = FreeFile
    Open GetLiteralQuotedRecordFieldTest For Input Access Read Shared As infileref
    outfileref = FreeFile
    Open GetLiteralQuotedRecordFieldResults For Output As outfileref
    
    Do Until EOF(infileref)
        record = GetNextLine(infileref)
        
        If Len(record) = 0 Then
            Print #outfileref, "     Last GetLiteralQuotedRecordField record processed."
            Exit Do
        End If
        
        index = InStr(record, ",")
        If index < 1 Then
            Print #outfileref, "FAIL:Not a valid GetLiteralQuotedRecordField record."
            Exit Do
        End If
        
        testIn = Left$(record, index - 1)
        
        index2 = InStr(index + 1, record, ",")
        
        If index2 < 1 Then
            Print #outfileref, "FAIL:Not a valid GetLiteralQuotedRecordField record."
            Exit Do
        End If
        
        indexIn = CInt(Mid$(record, index + 1, index2 - index - 1))
        If indexIn < 1 Then
            Print #outfileref, "FAIL:Not a valid GetLiteralQuotedRecordField record."
            Exit Do
        End If
        
        bench = Mid$(record, index2 + 1)
        actual = su.GetLiteralQuotedRecordField(testIn, indexIn)
        
        If StrComp(bench, actual, vbBinaryCompare) <> 0 Then
            prefix = "FAIL:"
            passed = False
        Else
            prefix = " OK :"
            passed = True
        End If
        
        Print #outfileref, "           " & RULER
        Print #outfileref, "     " & "INPUT:" & testIn & ":INDEX:" & Str$(indexIn)
        Print #outfileref, "     " & "BENCH:" & bench & ":"
        Print #outfileref, prefix & "RTURN:" & actual & ":"
        
        If Not passed Then errors = errors + 1
        
    Loop
    
    Test_GetLiteralQuotedRecordField = errors
    
    Close #infileref
    Close #outfileref
    Set su = Nothing
    
End Function


Function Test_GetLiteralQuotedField() As Long
    Dim su As StringUtilities
    Dim infileref As Integer
    Dim outfileref As Integer
    Dim prefix As String
    Dim testIn As String
    Dim indexIn As Long
    Dim delimitersIn As String
    Dim fieldIn As Integer
    Dim bench As String
    Dim actual As String
    Dim record As String
    Dim index As Long
    Dim index2 As Long
    Dim index3 As Long
    Dim index4 As Long
    Dim passed As Boolean
    Dim errors As Long
    
    Set su = New DDVariableStore.StringUtilities
    
    infileref = FreeFile
    Open GetLiteralQuotedFieldTest For Input Access Read Shared As infileref
    outfileref = FreeFile
    Open GetLiteralQuotedFieldResults For Output As outfileref
    
    Do Until EOF(infileref)
        record = GetNextLine(infileref)
        
        If Len(record) = 0 Then
            Print #outfileref, "     Last GetLiteralQuotedField record processed."
            Exit Do
        End If
        
        index = InStr(record, ",")
        If index < 1 Then
            Print #outfileref, "FAIL:Not a valid GetLiteralQuotedField record."
            Exit Do
        End If
        
        testIn = Left$(record, index - 1)
        
        index2 = InStr(index + 1, record, ",")
        
        If index2 < 1 Then
            Print #outfileref, "FAIL:Not a valid GetLiteralQuotedField record."
            Exit Do
        End If
        
        delimitersIn = Mid$(record, index + 1, index2 - index - 1)
        
        index3 = InStr(index2 + 1, record, ",")
        
        If index3 < 1 Then
            Print #outfileref, "FAIL:Not a valid GetLiteralQuotedField record."
            Exit Do
        End If
        
        indexIn = CLng(Mid$(record, index2 + 1, index3 - index2 - 1))
'        If indexIn < 1 Then
'            Print #outfileref, "FAIL:Not a valid GetNextNonBlankFieldIndex record."
'            Exit Do
'        End If
        
        index4 = InStr(index3 + 1, record, ",")
        
        If index4 < 1 Then
            Print #outfileref, "FAIL:Not a valid GetLiteralQuotedField record."
            Exit Do
        End If
        
        fieldIn = CInt(Mid$(record, index3 + 1, index4 - index3 - 1))
        
        bench = Mid$(record, index4 + 1)
        actual = su.GetLiteralQuotedField(indexIn, testIn, delimitersIn, fieldIn)
        
        If StrComp(bench, actual, 0) <> 0 Then
            prefix = "FAIL:"
            passed = False
        Else
            prefix = " OK :"
            passed = True
        End If
        
        Print #outfileref, "           " & RULER
        Print #outfileref, "     " & "INPUT:" & testIn & ":INDEX:" & Str$(indexIn)
        Print #outfileref, "     " & "DELIM:" & delimitersIn
        Print #outfileref, "     " & "FIELD:" & Str$(fieldIn)
        Print #outfileref, "     " & "BENCH:" & bench
        Print #outfileref, prefix & "RTURN:" & actual
        
        If Not passed Then errors = errors + 1
        
    Loop
    
    Test_GetLiteralQuotedField = errors
    
    Close #infileref
    Close #outfileref
    Set su = Nothing
    
End Function


Function Test_GetNonBlankField() As Long
    Dim su As StringUtilities
    Dim infileref As Integer
    Dim outfileref As Integer
    Dim prefix As String
    Dim testIn As String
    Dim indexIn As Long
    Dim delimitersIn As String
    Dim fieldIn As Integer
    Dim bench As String
    Dim actual As String
    Dim record As String
    Dim index As Long
    Dim index2 As Long
    Dim index3 As Long
    Dim index4 As Long
    Dim passed As Boolean
    Dim errors As Long
    
    Set su = New DDVariableStore.StringUtilities
    
    infileref = FreeFile
    Open GetNonBlankFieldTest For Input Access Read Shared As infileref
    outfileref = FreeFile
    Open GetNonBlankFieldResults For Output As outfileref
    
    Do Until EOF(infileref)
        record = GetNextLine(infileref)
        
        If Len(record) = 0 Then
            Print #outfileref, "     Last GetNonBlankField record processed."
            Exit Do
        End If
        
        index = InStr(record, ",")
        If index < 1 Then
            Print #outfileref, "FAIL:Not a valid GetNonBlankField record."
            Exit Do
        End If
        
        testIn = Left$(record, index - 1)
        
        index2 = InStr(index + 1, record, ",")
        
        If index2 < 1 Then
            Print #outfileref, "FAIL:Not a valid GetNonBlankField record."
            Exit Do
        End If
        
        delimitersIn = Mid$(record, index + 1, index2 - index - 1)
        
        index3 = InStr(index2 + 1, record, ",")
        
        If index3 < 1 Then
            Print #outfileref, "FAIL:Not a valid GetNonBlankField record."
            Exit Do
        End If
        
        indexIn = CLng(Mid$(record, index2 + 1, index3 - index2 - 1))
'        If indexIn < 1 Then
'            Print #outfileref, "FAIL:Not a valid GetNonBlankField record."
'            Exit Do
'        End If
        
        index4 = InStr(index3 + 1, record, ",")
        
        If index4 < 1 Then
            Print #outfileref, "FAIL:Not a valid GetNonBlankField record."
            Exit Do
        End If
        
        fieldIn = CInt(Mid$(record, index3 + 1, index4 - index3 - 1))
        
        bench = Mid$(record, index4 + 1)
        actual = su.GetNonBlankField(indexIn, testIn, delimitersIn, fieldIn)
        
        If StrComp(bench, actual, 0) <> 0 Then
            prefix = "FAIL:"
            passed = False
        Else
            prefix = " OK :"
            passed = True
        End If
        
        Print #outfileref, "           " & RULER
        Print #outfileref, "     " & "INPUT:" & testIn & ":INDEX:" & Str$(indexIn)
        Print #outfileref, "     " & "DELIM:" & delimitersIn
        Print #outfileref, "     " & "FIELD:" & Str$(fieldIn)
        Print #outfileref, "     " & "BENCH:" & bench
        Print #outfileref, prefix & "RTURN:" & actual
        
        If Not passed Then errors = errors + 1
        
    Loop
    
    Test_GetNonBlankField = errors
    
    Close #infileref
    Close #outfileref
    Set su = Nothing
    
End Function

Function Test_LastInStr() As Long
    Dim su As DDVariableStore.StringUtilities
    Dim errors As Long
    
    Dim sstring As String
    Dim mstring As String
    Dim status As Long
    Dim mstatus As Long
    
    Set su = New DDVariableStore.StringUtilities
    
    sstring = "Type=HTMLFrame;Index=1;\;Type=Combobox;Name=K1"
    mstring = "Type=Combobox"
    mstatus = 26
    status = su.LastInStr(-1, sstring, mstring, 1)
    If status <> mstatus Then errors = errors + 1
    
    sstring = "XXpxxpXXPXXX"
    mstring = "P"
    mstatus = 9
    status = su.LastInStr(-1, sstring, mstring, 1)
    If status <> mstatus Then errors = errors + 1
    
    sstring = "XXpxxpXXPXXX"
    mstring = "W"
    mstatus = 0
    status = su.LastInStr(-1, sstring, mstring, 1)
    If status <> mstatus Then errors = errors + 1
    
    sstring = "Type=HTMLFrame;Index=1;\;Type=Combobox;Name=K1"
    mstring = "Type=Combobox"
    mstatus = 26
    status = su.LastInStr(34, sstring, mstring, 1)
    If status <> mstatus Then errors = errors + 1
    
    sstring = "XXpxxpXXPXXX"
    mstring = "P"
    mstatus = 9
    status = su.LastInStr(10, sstring, mstring, 1)
    If status <> mstatus Then errors = errors + 1
    
    Test_LastInStr = errors
    
End Function

Function Test_SAFSVariables() As Long

    Dim vars As GlobalVariables
    Dim su As StringUtilities
    Dim result As Integer
    Dim value As Variant
    Dim expression As String
    Dim out As String
    Dim errors As Long
    
    Set vars = New GlobalVariables
    Set su = New StringUtilities
    
    result = vars.GetVariableValue("empty", value)
    result = vars.SetVariableValue("var1", "varValue1")
    result = vars.SetVariableValue("var2", "varValue2")
    result = vars.GetVariableValue("VAR1", value)
    result = vars.SetVariableValue("vaR3", value)
    result = vars.SetVariableValue("SAFSPROJECTDIRECTORY", "c:\sqarepos\ddengine")
    result = vars.SetVariableValue("SAFSDATAPOOLDIRECTORY", """c:\sqarepos\ddengine\datapool\""")
    result = vars.SetVariableValue("SAFSBENCHDIRECTORY", "c:\sqarepos\ddengine\datapool\bench\")
    result = vars.SetVariableValue("set", "value")
    result = vars.SetVariableValue("value", "3")
    result = vars.SetVariableValue("empty", "")
    result = vars.ResetAndClear()
    result = su.ConvertVariableExpression("^empty", vars, out)
    
    'GoTo T_EXIT
    
    result = su.ConvertVariableExpression("^value = (4&4)& 22", vars, out)
    result = su.ConvertVariableExpression("^value = ""A"" & ^value", vars, out)
    
    result = su.ConvertVariableExpression("^value = ^value1 = ^value2 = 4422 &(""=(4&4)&22"")", vars, out)
    result = su.ConvertVariableExpression("^value = ""A"" & ^value", vars, out)
    result = su.ConvertVariableExpression("^value1 = ""B"" & ^value1", vars, out)
    result = su.ConvertVariableExpression("^value2 = ""C"" & ^value2", vars, out)
    'result = vars.ResetAndClear()
T_EXIT:
    Set vars = Nothing
    Set su = Nothing
End Function

Function Test_MissingLValueExpressions() As Long
    Dim vars As GlobalVariables
    Dim su As StringUtilities
    Dim result As Integer
    Dim value As Variant
    Dim expression As String
    Dim out As String
    Dim errors As Long
    Dim outfileref As Integer
    
    
    Set vars = New GlobalVariables
    Set su = New StringUtilities
    outfileref = FreeFile
    MissingLValueExpressionsResults = "Test_MissingLValueExpressionsRESULTS.txt"
    Open MissingLValueExpressionsResults For Output As outfileref
    
    'result = vars.SetVariableValue("sValue", """Tool Assets""")
    result = su.ConvertVariableExpression("^sValue=""Tool Assets""", vars, out)
    Print #outfileref, "sValue=" & out
    result = su.ConvertVariableExpression("^sValue=""""""Tool Assets""""""", vars, out)
    Print #outfileref, "sValue=" & out
    result = su.ConvertVariableExpression("^sValue", vars, out)
    Print #outfileref, "sValue=" & out
    result = su.ConvertVariableExpression("%EF", vars, out)
    Print #outfileref, "EF=" & out
    result = su.ConvertVariableExpression("^sValue", vars, out)
    Print #outfileref, "sValue=" & out
    
    result = vars.ResetAndClear()
    
    Close #outfileref
    Set su = Nothing
    Set vars = Nothing
    
End Function

Sub Main()
    Dim status As Integer
    Dim errors As Long
    Dim total_errors As Long
    
    'ConvertVariableExpressionTest = App.Path & "\Test_MissingLValues.txt"
    'ConvertVariableExpressionResults = App.Path & "\Test_MissingLValueRESULTS.txt"
    errors = Test_MissingLValueExpressions()
    'If errors > 0 Then Shell "notepad.exe " & ConvertVariableExpressionResults
    If errors > 0 Then MsgBox "MissingLValueExpressions had " & Str$(errors) & " errors.", vbOKOnly, "Test MissingLValueExpressions"
    Shell "notepad.exe " & MissingLValueExpressionsResults
    total_errors = total_errors + errors

GoTo QUICK_EXIT
    
    errors = Test_SAFSVariables()
    If errors > 0 Then MsgBox "Test SAFSVariables had " & Str$(errors) & " errors.", vbOKOnly, "Test SAFSVariables"
    total_errors = total_errors + errors
    'Exit Sub

    errors = Test_LastInStr()
    'If errors > 0 Then Shell "notepad.exe " & TWhitespaceResults
    If errors > 0 Then MsgBox "LastInStr had " & Str$(errors) & " errors.", vbOKOnly, "Test LastInStr"
    total_errors = total_errors + errors
    
    ConvertHTMLTests = App.Path & "\Test_ConvertHTML.txt"
    ConvertHTMLResults = App.Path & "\Test_ConvertHTMLRESULTS.txt"
    errors = Test_ConvertHTMLEntities()
    If errors > 0 Then Shell "notepad.exe " & ConvertHTMLResults
    If errors > 0 Then MsgBox "ConvertHTMLEntities had " & Str$(errors) & " errors.", vbOKOnly, "Test ConverHTMLEntities"
    total_errors = total_errors + errors
        
'GoTo QUICK_EXIT

    TWhitespaceTest = App.Path & "\Test_TWhitespace.txt"
    TWhitespaceResults = App.Path & "\Test_TWhitespaceRESULTS.txt"
    errors = Test_TWhitespace()
    If errors > 0 Then Shell "notepad.exe " & TWhitespaceResults
    If errors > 0 Then MsgBox "TWhitespace had " & Str$(errors) & " errors.", vbOKOnly, "Test TWhitespace"
    total_errors = total_errors + errors
    
    IsQuotedTest = App.Path & "\Test_IsQuoted.txt"
    IsQuotedResults = App.Path & "\Test_IsQuotedRESULTS.txt"
    errors = Test_IsQuoted()
    If errors > 0 Then Shell "notepad.exe " & IsQuotedResults
    If errors > 0 Then MsgBox "IsQuoted had " & Str$(errors) & " errors.", vbOKOnly, "Test IsQuoted"
    total_errors = total_errors + errors

    ValidateVariableIDTest = App.Path & "\Test_ValidateVariableID.txt"
    ValidateVariableIDResults = App.Path & "\Test_ValidateVariableIDRESULTS.txt"
    errors = Test_ValidateVariableID()
    If errors > 0 Then Shell "notepad.exe " & ValidateVariableIDResults
    If errors > 0 Then MsgBox "ValidateVariableID had " & Str$(errors) & " errors.", vbOKOnly, "Test ValidateVariableID"
    total_errors = total_errors + errors

    UnQuotedDelimiterIndexTest = App.Path & "\Test_UnQuotedDelimiterIndex.txt"
    UnQuotedDelimiterIndexResults = App.Path & "\Test_UnQuotedDelimiterIndexRESULTS.txt"
    errors = Test_UnQuotedDelimiterIndex()
    If errors > 0 Then Shell "notepad.exe " & UnQuotedDelimiterIndexResults
    If errors > 0 Then MsgBox "UnQuotedDelimiterIndex had " & Str$(errors) & " errors.", vbOKOnly, "Test UnQuotedDelimiterIndex"
    total_errors = total_errors + errors

    UnQuotedDelimiterIndexREVTest = App.Path & "\Test_UnQuotedDelimiterIndexREV.txt"
    UnQuotedDelimiterIndexREVResults = App.Path & "\Test_UnQuotedDelimiterIndexREVRESULTS.txt"
    errors = Test_UnQuotedDelimiterIndexREV()
    If errors > 0 Then Shell "notepad.exe " & UnQuotedDelimiterIndexREVResults
    If errors > 0 Then MsgBox "UnQuotedDelimiterIndexREV had " & Str$(errors) & " errors.", vbOKOnly, "Test UnQuotedDelimiterIndexREV"
    total_errors = total_errors + errors

    ConvertVariableExpressionTest = App.Path & "\Test_ConvertVariableExpression.txt"
    ConvertVariableExpressionResults = App.Path & "\Test_ConvertVariableExpressionRESULTS.txt"
    errors = Test_ConvertVariableExpression()
    If errors > 0 Then Shell "notepad.exe " & ConvertVariableExpressionResults
    If errors > 0 Then MsgBox "ConvertVariableExpression had " & Str$(errors) & " errors.", vbOKOnly, "Test ConvertVariableExpression"
    total_errors = total_errors + errors

    GetNextNonBlankFieldIndexTest = App.Path & "\Test_GetNextNonBlankFieldIndex.txt"
    GetNextNonBlankFieldIndexResults = App.Path & "\Test_GetNextNonBlankFieldIndexRESULTS.txt"
    errors = Test_GetNextNonBlankFieldIndex()
    If errors > 0 Then Shell "notepad.exe " & GetNextNonBlankFieldIndexResults
    If errors > 0 Then MsgBox "GetNextNonBlankFieldIndex had " & Str$(errors) & " errors.", vbOKOnly, "Test GetNextNonBlankFieldIndex"
    total_errors = total_errors + errors

    GetNonBlankFieldTest = App.Path & "\Test_GetNonBlankField.txt"
    GetNonBlankFieldResults = App.Path & "\Test_GetNonBlankFieldRESULTS.txt"
    errors = Test_GetNonBlankField()
    If errors > 0 Then Shell "notepad.exe " & GetNonBlankFieldResults
    If errors > 0 Then MsgBox "GetNonBlankField had " & Str$(errors) & " errors.", vbOKOnly, "Test GetNonBlankField"
    total_errors = total_errors + errors

    GetFieldCountTest = App.Path & "\Test_GetFieldCount.txt"
    GetFieldCountResults = App.Path & "\Test_GetFieldCountRESULTS.txt"
    errors = Test_GetFieldCount()
    If errors > 0 Then Shell "notepad.exe " & GetFieldCountResults
    If errors > 0 Then MsgBox "GetFieldCount had " & Str$(errors) & " errors.", vbOKOnly, "Test GetFieldCount"
    total_errors = total_errors + errors

    GetTrimmedFieldTest = App.Path & "\Test_GetTrimmedField.txt"
    GetTrimmedFieldResults = App.Path & "\Test_GetTrimmedFieldRESULTS.txt"
    errors = Test_GetTrimmedField()
    If errors > 0 Then Shell "notepad.exe " & GetTrimmedFieldResults
    If errors > 0 Then MsgBox "GetTrimmedField had " & Str$(errors) & " errors.", vbOKOnly, "Test GetTrimmedField"
    total_errors = total_errors + errors

    GetLiteralQuotedFieldTest = App.Path & "\Test_GetLiteralQuotedField.txt"
    GetLiteralQuotedFieldResults = App.Path & "\Test_GetLiteralQuotedFieldRESULTS.txt"
    errors = Test_GetLiteralQuotedField()
    If errors > 0 Then Shell "notepad.exe " & GetLiteralQuotedFieldResults
    If errors > 0 Then MsgBox "GetLiteralQuotedField had " & Str$(errors) & " errors.", vbOKOnly, "Test GetLiteralQuotedField"
    total_errors = total_errors + errors

    GetLiteralQuotedRecordFieldTest = App.Path & "\Test_GetLiteralQuotedRecordField.txt"
    GetLiteralQuotedRecordFieldResults = App.Path & "\Test_GetLiteralQuotedRecordFieldRESULTS.txt"
    errors = Test_GetLiteralQuotedRecordField()
    If errors > 0 Then Shell "notepad.exe " & GetLiteralQuotedRecordFieldResults
    If errors > 0 Then MsgBox "GetLiteralQuotedRecordField had " & Str$(errors) & " errors.", vbOKOnly, "Test GetLiteralQuotedRecordField"
    total_errors = total_errors + errors

QUICK_EXIT:

    MsgBox "DDVariableStore had" & Str$(total_errors) & " total testing errors.", vbOKOnly, "Test DDVariableStore"
    
    Close
End Sub


