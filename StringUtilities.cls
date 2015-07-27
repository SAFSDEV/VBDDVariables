VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "String functions and String  expression processing"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'##############################################################################
'# <a name="stringutilities"></a>
'#
'#    StringUtilities.CLS  SOURCECODE
'#
'# MODULE DESCRIPTION:
'#
'#      Utility routines for working with Strings.
'#      Create an instance of this class through standard ActiveX means.
'#
'#      This Public class is made available via the DDVariableStore.DLL server.
'#
'#      Typical VisualBasic invocations:
'#
'#          Set oStringUtils = DDVariableStore.StringUtilities
'#
'#              OR
'#
'#          Set oStringUtils = CreateObject("DDVariableStore.StringUtilities")
'#
'#
'# This library is defined by:
'#
'# StringUtilities.CLS      ActiveX Class
'# DDVariableStore.DLL      Compiled executable library
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#      NOV 04, 2001    (Carl Nagle) Fixed Variable Assignment in Groupings bug.
'#      APR 11, 2002    (Carl Nagle) LTrim numeric values in ConvertVariableExpression.
'#                               Recoded LastInStr to eliminate use of VB InStrRev
'#                               which apparently has a bug.
'#      JAN 16, 2003    (Carl Nagle) Added GetMultiDelimitedField.
'#      MAR 28, 2003    (Carl Nagle) Slight perf. improvement and more doc.
'#      NOV 03, 2003    '(Carl Nagle) Fixed bug: variable with no stored value in QTTextOfVariable.
'#
'# Copyright (2002,2003) SAS Institute Inc. All rights reserved.
'#  General Public License: http://www.opensource.org/licenses/gpl-license.php
'#
'##############################################################################


'# FORWARD DECLARATIONS


'############################################################################
'# <a name="ltwhitespace"></a>
'#
'#  Function LTWhitespace (strText As String) As String
'#
'# DESCRIPTION:
'#
'#      This routine removes leading (Left) whitespace (spaces and tabs)
'#      from the input string.
'#
'#
'# PARAMETERS:
'#
'#      strText = the string to trim of leading tabs and spaces.
'#
'#
'# RETURNS:
'#
'#      A new string trimmed of leading tabs and spaces.
'#
'#
'# ERRORS:
'#
'#       none
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function LTWhitespace(strText As String) As String
    Dim strStart As String
    Dim strTemp As String
    Dim notDone As Boolean
    
    If Len(strText) = 0 Then Exit Function
    
    notDone = True
    strTemp = strText
    Do While notDone And (Len(strTemp) > 0)
        strStart = strTemp
        strTemp = LTrim$(strStart)
        If Left$(strTemp, 1) = Chr$(9) Then strTemp = Mid$(strTemp, 2)
        notDone = Len(strStart) <> Len(strTemp)
    Loop
    
    LTWhitespace = strTemp
End Function


'############################################################################
'# <a name="rtwhitespace"></a>
'#
'#  Function RTWhitespace (strText As String) As String
'#
'# DESCRIPTION:
'#
'#      This routine removes trailing (Right) whitespace (spaces and tabs)
'#      from the input string.
'#
'#
'# PARAMETERS:
'#
'#      strText = the string to trim of trailing tabs and spaces.
'#
'#
'# RETURNS:
'#
'#      A new string trimmed of trailing tabs and spaces.
'#
'#
'# ERRORS:
'#
'#       none
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function RTWhitespace(strText As String) As String
    Dim strStart As String
    Dim strTemp As String
    Dim notDone As Boolean
    
    If Len(strText) = 0 Then Exit Function
    
    notDone = True
    strTemp = strText
    Do While notDone And (Len(strTemp) > 0)
        strStart = strTemp
        strTemp = RTrim$(strStart)
        If Right$(strTemp, 1) = Chr$(9) Then strTemp = Left$(strTemp, Len(strTemp) - 1)
        notDone = Len(strStart) <> Len(strTemp)
    Loop
    
    RTWhitespace = strTemp
End Function


'############################################################################
'# <a name="twhitespace"></a>
'#
'#  Function TWhitespace (strText As String) As String
'#
'# DESCRIPTION:
'#
'#      This routine removes both leading and trailing whitespace (spaces and tabs)
'#      from the input string.
'#
'#
'# PARAMETERS:
'#
'#      strText = the string to trim of leading and trailing tabs and spaces.
'#
'#
'# RETURNS:
'#
'#      A new string trimmed of leading and trailing tabs and spaces.
'#
'#
'# ERRORS:
'#
'#       none
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function TWhitespace(strText As String) As String
    Dim strTemp As String
    strTemp = LTWhitespace(strText)
    TWhitespace = RTWhitespace(strTemp)
End Function




'############################################################################
'# <a name="firstinstr"></a>
'#
'#  Function FirstInStr (startIndex as Long, text As String,
'#                      match As String, noCase As Integer) As Long
'#
'# DESCRIPTION:
'#
'#      This surfaces the InStr function of VB.
'#      Search the text string for the first occurence of the match string.
'#      The search begins at startIndex. If startIndex is < 1 then the
'#      search begins at index 1.
'#
'#
'# PARAMETERS:
'#
'#      startIndex  the index to begin at. If < 1 then we will start at 1.
'#                  Index is 1 based.
'#      text        the string to examine for the last occurrence of match.
'#      match       the substring to find the last occurrence within source.
'#      noCase      0 means compare is case-sensitive(binary).
'#                  1 means it is not (text-based).
'#
'# RETURNS:
'#
'#      0   on failure or no match found.
'#          Otherwise returns the index of the last occurrence of the substring
'#          in the test string.
'#
'# ERRORS:
'#
'#       none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function FirstInStr(startindex As Long, text As String, _
                          match As String, noCase As Integer) As Long

    Dim lIndex As Long
    
    On Error GoTo FISError
    
    If startindex < 1 Then startindex = 1
    
    FirstInStr = InStr(startindex, text, match, noCase)
    
FISError:
    On Error GoTo 0
    Exit Function
End Function


'############################################################################
'# <a name="lastinstr"></a>
'#
'#  Function LastInStr (startIndex as Long, text As String, match As String,
'#                      noCase As Integer) As Long
'#
'# DESCRIPTION:
'#
'#      This mimics the FindInStr function only indexes backwards through
'#      the String.  Generally, this is just making VBs InStrREV function
'#      available.
'#
'# PARAMETERS:
'#
'#      startIndex  the index to begin at. If < 1 or > Len(source)
'#                  then we will start at the end of source.  Index is 1 based.
'#      text        the string to examine for the last occurrence of match.
'#      match       the substring to find the last occurrence within source.
'#      noCase      0 means compare is case-sensitive(binary).
'#                  1 means it is not (text-based).
'#
'# RETURNS:
'#
'#      0   on failure or no match found.
'#          Otherwise returns the index of the last occurrence of the substring
'#          in the test string.
'#
'# ERRORS:
'#
'#       none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#      APR 11, 2002    (Carl Nagle) Remove use of MS InStrRev which has defect.
'#                               Recoded to use InStr repetitively.
'#
'############################################################################
Public Function LastInStr(startindex As Long, text As String, match As String, _
                          noCase As Integer) As Long

    Dim index As Integer
    Dim status As Integer
    Dim lenSub As Integer
    Dim lenTest As Integer
    Dim maxStart As Integer
    
    LastInStr = 0
    lenSub = Len(match)
    lenTest = Len(text)
        
    'if either string is invalid then exit
    If (lenSub = 0) Or (lenTest = 0) Then Exit Function
    
    'if the substring is larger than the test string then exit
    If lenSub > lenTest Then Exit Function
    
    'if noCase is not a valid value then exit
    If noCase < 0 Or noCase > 1 Then Exit Function
    
    'find the first index into text that could possibly contain match
    maxStart = lenTest - lenSub + 1
    
    If startindex < 1 Or startindex > maxStart Then startindex = maxStart
    
    'try to find the last occurrence of sSub in sTest
    For index = startindex To 1 Step -1
        status = InStr(index, text, match, noCase)
        If status > 0 Then Exit For
    Next index
    
    LastInStr = index
        
End Function


'############################################################################
'# <a name="isquoted"></a>
'#
'#  Function IsQuoted (strText As String, position As Long) As Boolean
'#
'# DESCRIPTION:
'#
'#      Returns TRUE if the provided position within the string is considered
'#      to be "inside" quoted (literal) text.  The routine evaluates all the
'#      string from left to right looking for quote marks.
'#
'#      The outermost quote marks are not considered to be "inside" quoted text.
'#      Two back-to-back double-quote marks indicate a single double-quote mark
'#      inside quoted text.  That means:
'#
'#          " text containing a quote mark "" looks like this"
'#
'#      The above gets interpretted as:
'#
'#           text containing a quote mark " looks like this
'#
'#      Thus, the quotes in the middle of the input string are both considered
'#      "inside" the quoted text.
'#
'#
'# PARAMETERS:
'#
'#      strText     the string to evaluate for quote marks preceding position.
'#
'#      position    the position or index to determine if it is "quoted" or
'#                  "inside" literal quoted text.
'#
'#
'# RETURNS:
'#
'#      Boolean     TRUE (-1) if the position is inside quoted text.
'#                  FALSE (0) otherwise.
'#
'#
'# ERRORS:
'#
'#       none
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function IsQuoted(strText As String, position As Long) As Boolean
    Dim lvalue As String
    Dim lenlvalue As Integer
    Dim Item As String
    Dim locQT As Long
    Dim quoted As Boolean
    Dim doubles As Boolean
    Dim itemIsQuote As Boolean
    
    IsQuoted = False
    
    'check special case of first character only
    'this shouldn't happen, hopefully
    If position < 2 Then Exit Function  'can't possibly be quoted
    
    'get everything to the left of position
    lvalue = Left$(strText, position - 1)
    lenlvalue = position - 1
    Item = Mid$(strText, position, 1)
    
    itemIsQuote = False
    If Item = gQT Then itemIsQuote = True
    
    locQT = 1
    quoted = False
    Do While (locQT > 0) And (locQT <= lenlvalue)
        
        doubles = False
        'if no (more) quotes then we're done
        locQT = InStr(locQT, lvalue, gQT, vbTextCompare)
        If locQT < 1 Then Exit Do
        
        'we found (another) one, toggle the state of quoted
        quoted = Not quoted
        
        'exit if this quote is position - 1
        If locQT = lenlvalue Then Exit Do
        
        'see if this is two double-quotes--those don't count
        If Mid$(lvalue, locQT + 1, 1) = gQT Then
            quoted = Not quoted
            doubles = True
        End If
        
        'we tested the next location already,
        'so add 2 instead of 1 for the next loop
        locQT = locQT + 2
    Loop
        
    If itemIsQuote Then 'that changes everything!
        
        'if the last char was a quote
        If locQT = lenlvalue Then
            If Not doubles Then 'it is now
                quoted = Not quoted
            Else
                'OK we are following double quote marks
                If quoted Then
                    'a closing quote is NOT inside a quoted string
                    If position = Len(strText) Then
                        quoted = False
                    Else
                        'see if position is a closing quote mark
                        If Mid$(strText, position + 1, 1) <> gQT Then
                            quoted = False
                        End If
                    End If
                End If
            End If
        
        'last one wasn't a quote, but I am
        Else
            'check to see if position is a closing quote
            'a closing quote is NOT inside a quoted string
            If quoted Then
                If position = Len(strText) Then
                    quoted = False
                Else
                    'see if position is a closing quote mark
                    If Mid$(strText, position + 1, 1) <> gQT Then
                        quoted = False
                    End If
                End If
            End If
        End If
        
    End If
    IsQuoted = quoted
    
End Function


'############################################################################
'# <a name="validatevariableid"></a>
'#
'#  Function ValidateVariableID (variable As String) As String
'#
'# DESCRIPTION:
'#
'#      Evaluates if the input variable string conforms to the naming convention
'#      for variables.  That means it must contain the leading variable
'#      identifier (default is caret ^) and contain no spaces or invalid
'#      variable name characters.  The string is first trimmed of whitespace
'#      before evaluated.
'#
'#      If the provided string conforms, then the string is returned without
'#      the leading variable identifier--just the variable name.
'#      Otherwise, the return string is empty.
'#
'#      Variable naming convention:
'#
'#      1) Must start with the variable identifier (default is ^ caret)
'#      2) Contains only the following characters, in any order:
'#
'#              'a-z
'#              codes 97 To 122
'#
'#              '0-9
'#              codes 48 To 57
'#
'#              'A-Z
'#              codes 65 To 90
'#
'#              ' . (period)
'#              code 46
'#
'#              ' _ (underscore)
'#              code 95
'#
'#              'extended accented characters
'#              codes 192 To 246
'#              codes 248 To 255
'#              code 131            ' f  (Florin)
'#              code 140            ' OE  (OE ligature)
'#              code 156            ' oe  (oe ligature)
'#              code 159            ' Y  (Y umlaut)
'#
'#
'# PARAMETERS:
'#
'#      variable    the string to validate as a variable ID
'#
'#
'# RETURNS:
'#
'#      The variable name (without the variable identifier) if successful.
'#      Otherwise, an empty string.
'#
'#
'# ERRORS:
'#
'#       none
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function ValidateVariableID(variable As String) As String
    Dim strText As String
    Dim index As Long
    Dim endstr As Long
    Dim code As Long
    
    strText = TWhitespace(variable)
    If Left$(strText, 1) <> gVID Then Exit Function
    
    'strip off the leading caret ^
    strText = Mid$(strText, 2)
    endstr = Len(strText)
    
    For index = 1 To endstr

        'gValidDVNameCharacters is populated during Class_Initialize
        'If InStr(gValidDVNameCharacters, Mid$(strText, index, 1)) > 0 Then GoTo VVIDNext
        
        'if we made it to here, the item is invalid
        'Exit Function
                
        code = Asc(Mid$(strText, index, 1))
        
        If code < 46 Then Exit Function
        If code = 47 Then Exit Function
        If code < 58 Then GoTo VVIDNext     '46, 48-57   OK  "." "0-9"
        If code < 65 Then Exit Function     '58-64   NOT OK
        If code < 91 Then GoTo VVIDNext     '65-90       OK  "A-Z"
        If code < 95 Then Exit Function     '91-94   NOT OK
        If code < 123 Then GoTo VVIDNext    '95-122      OK  "a-z"
        If code < 129 Then Exit Function    '123-128 NOT OK
        
        If code = 131 Then Exit Function
        If code = 136 Then Exit Function
        If code = 139 Then Exit Function
        
        If code < 145 Then GoTo VVIDNext    '129-144 OK except 131, 136, 139
        If code < 154 Then Exit Function    '145-153 NOT OK
        If code < 160 Then GoTo VVIDNext    '154-159 OK
        
        If code = 168 Then GoTo VVIDNext
        If code < 180 Then Exit Function    '160-179 NOT OK except 168
        If code < 182 Then GoTo VVIDNext    '180-181 OK
        If code < 192 Then Exit Function    '182-191 NOT OK
                    
        If code = 215 Then Exit Function
        If code = 247 Then Exit Function    '192-255 OK except 215, 247
                
VVIDNext:
    Next
    
    ValidateVariableID = strText
    
End Function

Private Function optGetDelimiterIndex(startindex As Long, _
                                      text As String, lenText As Long, _
                                      delimiters As String, lenDelimiters As Long) As Long
    Dim ascText As Integer
    Dim ascDelimiter As Integer
    Dim textIndex As Long
    Dim delimiterIndex As Long
    Dim done As Boolean
    
    optGetDelimiterIndex = -1
    
    done = False
    textIndex = startindex
    Do While (Not done) And (textIndex <= lenText)
        delimiterIndex = 1
        Do While (Not done) And (delimiterIndex <= lenDelimiters)
            If Mid$(text, textIndex, 1) = Mid$(delimiters, delimiterIndex, 1) Then
                done = True
            Else
                delimiterIndex = delimiterIndex + 1
            End If
        Loop
        If Not done Then textIndex = textIndex + 1
    Loop
    If done Then optGetDelimiterIndex = textIndex 'otherwise -1

End Function


'############################################################################
'# <a name="getdelimiterindex"></a>
'#
'#  Function GetDelimiterIndex (startindex As Long, text As String,
'#                                  delimiters As String) As Long
'#
'# DESCRIPTION:
'#
'#      Finds the index of the first character matching one of the provided
'#      delimiter characters.  The search begins at startindex within the text.
'#      The index is 1-based.
'#
'#
'# PARAMETERS:
'#
'#      startindex      Long (1-based) start position for search in text.
'#                      If startindex is < 1 then we use startindex = 1.
'#
'#      text            String to search for first delimiter character
'#
'#      delimiters      String list of delimiters to find.
'#
'#
'# RETURNS:
'#
'#      -1   on error, invalid parameters, or end of text reached.
'#       N   (Long) index of first delimiter character found.
'#
'#
'# ERRORS:
'#
'#      none
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function GetDelimiterIndex(startindex As Long, text As String, delimiters As String) As Long
    
    Dim lenText As Long
    Dim lenDelimiters As Long
    
    GetDelimiterIndex = -1
    lenText = Len(text)
    If lenText = 0 Then Exit Function
    lenDelimiters = Len(delimiters)
    If lenDelimiters = 0 Then Exit Function
    If startindex > lenText Then Exit Function
    If startindex < 1 Then startindex = 1
    
    GetDelimiterIndex = optGetDelimiterIndex(startindex, text, lenText, delimiters, lenDelimiters)

End Function


Private Function optGetUnquotedDelimiterIndex(startindex As Long, _
                                              text As String, lenText As Long, _
                                              delimiters As String, lenDelimiters As Long) As Long

    Dim index As Long
    
    index = startindex
    Do
        index = optGetDelimiterIndex(index, text, lenText, delimiters, lenDelimiters)
        If index > 0 Then
            If IsQuoted(text, index) Then
                index = index + 1
            Else
                Exit Do
            End If
        End If
    Loop Until index < 1 Or index > lenText
    If index > lenText Then index = -1
    optGetUnquotedDelimiterIndex = index

End Function


'############################################################################
'# <a name="getunquoteddelimiterindex"></a>
'#
'#  Function GetUnquotedDelimiterIndex (startindex As Long, text As String,
'#                                  delimiters As String) As Long
'#
'# DESCRIPTION:
'#
'#      Finds the index of the first character matching one of the provided delimiter
'#      characters that is NOT inside a quoted literal string.  Essentially,
'#      this is the same as calling GetDelimiterIndex followed by IsQuoted
'#      and only returning the index if the delimiter is NOT quoted.
'#      The search for the delimiter begins at startindex within the text.
'#      The index is 1-based.
'#      The test for whether the delimiter is quoted uses the full text string.
'#
'#
'# PARAMETERS:
'#
'#      startindex      Long (1-based) start position for search in text
'#                      If startindex is < 1 then we use startindex = 1.
'#
'#      text            String to search for first unquoted delimiter character
'#
'#      delimiters      String list of delimiters to find.
'#
'#
'# RETURNS:
'#
'#      -1   on error, invalid parameters, or end of text reached.
'#       N   (Long) index of first delimiter character found.
'#
'#
'# ERRORS:
'#
'#      none
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function GetUnquotedDelimiterIndex(startindex As Long, text As String, delimiters As String) As Long
    
    Dim lenText As Long
    Dim lenDelimiters As Long
    
    GetUnquotedDelimiterIndex = -1
    lenText = Len(text)
    lenDelimiters = Len(delimiters)
    If startindex > lenText Then Exit Function
    If startindex < 1 Then startindex = 1
    
    GetUnquotedDelimiterIndex = optGetUnquotedDelimiterIndex(startindex, text, lenText, _
                                                             delimiters, lenDelimiters)
                                                             
End Function


Private Function optGetFieldIndex(startindex As Long, text As String, lenText As Long, _
                                  delimiters As String, lenDelimiters As Long, _
                                  fieldID As Integer) As Long
    Dim lI1 As Long
    Dim status As Long
    Dim i As Integer
    
    'explicit set for failure
    optGetFieldIndex = -1
    
    'the first field is always startIndex
    If fieldID = 1 Then
        optGetFieldIndex = startindex
        Exit Function
    End If
    
    lI1 = startindex
           
    For i = 2 To fieldID
        status = optGetDelimiterIndex(lI1, text, lenText, delimiters, lenDelimiters)
        If status = -1 Then Exit Function                'exit with error -1
        
        'increment to the actual first char of the next field
        lI1 = status + 1
    Next i
    
    'possible errors
    If lI1 > lenText Then Exit Function     'exit with error -1
    
    optGetFieldIndex = lI1
                                  
End Function


'############################################################################
'# <a name="getfieldindex"></a>
'#
'#  Function GetFieldIndex (startindex As Long, text As String,
'#                          delimiters As String, fieldID As Integer) As Long
'#
'# DESCRIPTION:
'#
'#      Finds the index of the nth field within the text.
'#      The search and field count begins at startindex within the text.
'#      The field and index are 1-based.
'#
'#      Blank fields count, so finding any one delimiter in our delimiters string
'#      will indicate the end of a field. The next field starts at the next
'#      character index even if that char is also a delimiter.
'#
'#      Field #1 is always the same as startindex--even if the character at
'#      the startindex position is a delimiter. In that event, field #1 is
'#      a zero-length string.
'#
'# PARAMETERS:
'#
'#      startindex      Long (1-based) start position for search in text
'#
'#      text            String to search for first delimiter character
'#
'#      delimiters      String list of delimiters that separate fields.
'#                      delimiters CAN be zero length if field #1 is requested.
'#                      (Kinda silly, but it is allowed.)
'#
'#      fieldID         Integer (1-based) of field # to find the start of.
'#                      field #1 begins at startindex regardless of the true
'#                      start of the text.
'#
'# RETURNS:
'#
'#     -1   on error, invalid parameters, or end of text reached.
'#      N   (Long) index of first non-blank field found.
'#
'# ERRORS:
'#
'#      none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function GetFieldIndex(startindex As Long, text As String, _
                              delimiters As String, fieldID As Integer) As Long

    Dim lenText As Long
    Dim lenDelimiters As Long
    
    'explicit set for failure
    GetFieldIndex = -1
    
    'we don't necessarily have to have delimiters if we want the first field
    lenDelimiters = Len(delimiters)
    If (lenDelimiters = 0) And (fieldID > 1) Then Exit Function
    
    lenText = Len(text)
    If lenText = 0 Then Exit Function
    
    'insure startindex has valid setting
    If startindex > lenText Then Exit Function
    If startindex < 1 Then startindex = 1
    
    'error if fieldId < 1
    If fieldID < 1 Then Exit Function
    
    GetFieldIndex = optGetFieldIndex(startindex, text, lenText, delimiters, lenDelimiters, fieldID)

End Function


Private Function optGetFieldCount(startindex As Long, _
                               text As String, lenText As Long, _
                               delimiters As String, lenDelimiters As Long) As Integer

    Dim i As Integer
    Dim rIndex As Long
    
    i = 0                   'field counter
    rIndex = startindex
    Do
        'always have at least 1 field
        i = i + 1
        rIndex = optGetDelimiterIndex(rIndex, text, lenText, delimiters, lenDelimiters)
        
        'force finish if last character is a delimiter
        If rIndex = lenText Then
            rIndex = 0
        Else
            rIndex = rIndex + 1
        End If
    Loop Until rIndex < 1
    
    optGetFieldCount = i

End Function


'############################################################################
'# <a name="getfieldcount"></a>
'#
'#  Function GetFieldCount (startindex As Long, text As String,
'#                          delimiters As String) As Integer
'#
'# DESCRIPTION:
'#
'#      Finds the count of all fields within the text from
'#      startindex to the end of the text.
'#      The field and index are 1-based.
'#
'#      Blank fields count, so finding any one delimiter in our delimiters string
'#      will indicate the end of a field. The next field starts at the next
'#      character index even if that char is also a delimiter.
'#      If the last character in the text is a delimiter, then the
'#      field it terminates is the last field in the text.
'#
'#      Field #1 index is always the same as startindex--even if the character at
'#      the startindex position is a delimiter.
'#
'# PARAMETERS:
'#
'#      startindex      Long (1-based) start position for field count in
'#                      text.  If startIndex < 1 then we will use
'#                      startindex = 1.
'#
'#      text            String to count fields beginning at startindex.
'#
'#      delimiters      String list of possible delimiters that separate fields.
'#                      If there are no delimiters, then there is only 1 field.
'#
'# RETURNS:
'#
'#      0   Indicates bad input (like 0-length text or a startindex
'#          greater than the text length).
'#
'#      N   (Integer) The number of fields found beginning at startindex.
'#
'# ERRORS:
'#
'#      none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: JUL 19, 2000
'# History:
'#
'#      JUL 19, 2000    Original Release
'#      SEP 28, 2000    (Carl Nagle) Optimized
'#      AUG 16, 2001    Converted to DLL use
'#
'############################################################################
Public Function GetFieldCount(startindex As Long, text As String, _
                              delimiters As String) As Integer

    Dim lenText As Long, lenDelimiters As Long

    ' default return code GetFieldCount = 0
    lenText = Len(text)
    If lenText = 0 Then Exit Function
    If startindex > lenText Then Exit Function
    lenDelimiters = Len(delimiters)
    If startindex < 1 Then startindex = 1
    
    GetFieldCount = optGetFieldCount(startindex, text, lenText, delimiters, lenDelimiters)

End Function


Private Function optGetField(startindex As Long, _
                             text As String, lenText As Long, _
                             delimiters As String, lenDelimiters As Long, _
                             fieldID As Integer) As String
                         
    Dim lI1 As Long
    Dim lI2 As Long
    
    lI1 = optGetFieldIndex(startindex, text, lenText, delimiters, lenDelimiters, fieldID)
    If lI1 < 1 Then Exit Function
    
    'ANY field whose first char is a delimiter is an empty field! Isn't it?
    If InStr(delimiters, Mid$(text, lI1, 1)) > 0 Then Exit Function
        
    lI2 = optGetDelimiterIndex(lI1 + 1, text, lenText, delimiters, lenDelimiters)
    If lI2 > lI1 Then
        optGetField = Mid$(text, lI1, lI2 - lI1)
    Else
        optGetField = Mid$(text, lI1)
    End If
                         
End Function


'############################################################################
'# <a name="getfield"></a>
'#
'#  Function GetField (startIndex as Long, text As String,
'#                     delimiters As String, fieldID As Integer) As String
'#
'# DESCRIPTION:
'#
'#      Given text of delimited fields, return the nth(FieldID) Field
'#      in the record from startIndex.  The index is 1 based.  The delimiter
'#      string can contain multiple characters but each character is individually
'#      treated as a possible delimiter.
'#
'# PARAMETERS:
'#
'#      startIndex      the one based index to begin search.  if < 1 then we
'#                      will start at index = 1.
'#      text            string of delimited fields
'#      delimiters       the string (>0 characters) that delimits each field.
'#                      As indicated this can be more than one character in
'#                      length--each character is separately a possible delimiter.
'#      fieldID         the one based index of the field to find and return. If < 1
'#                      then we will use fieldID = 1.
'#
'#
'# RETURNS:
'#
'#      A string of the indexed field.
'#      An empty string ("") is returned if the fieldID is out of range for
'#      the provided text or the provided delimiter is invalid.
'#
'# ERRORS:
'#
'#       none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: Aug 15, 2001
'# History:
'#
'#      Aug 15, 2001    Original Release
'#
'############################################################################
Public Function GetField(startindex As Long, text As String, _
                         delimiters As String, fieldID As Integer) As String
    
    Dim lenText As Long, lenDelimiters As Long
    
    lenText = Len(text)
    If lenText = 0 Then Exit Function
    If startindex > lenText Then Exit Function
    lenDelimiters = Len(delimiters)
    If startindex < 1 Then startindex = 1
    If fieldID < 1 Then fieldID = 1
    
    GetField = optGetField(startindex, text, lenText, delimiters, lenDelimiters, fieldID)
    
End Function


'############################################################################
'# <a name="getmultidelimitedfieldcount"></a>
'#
'#  Function GetMultiDelimitedFieldCount (startindex As Long, text As String,
'#                                        delimiter As String) As Integer
'#
'# DESCRIPTION:
'#
'#      Finds the count of all fields within the text from
'#      startindex to the end of the text. The field and index are 1-based.
'#
'#      The delimiter string can contain multiple characters and this string
'#      is treated as a single delimiter.
'#
'#      For example, if delimiter = "->" then the fields of the text string
'#      are expected to be delimited by the full delimiter string as in:
'#
'#      "field1->field2->field3"
'#
'#      Blank fields count. The next field starts at the next
'#      character index even if that char is also a delimiter.
'#      If the last character in the text is a delimiter, then the
'#      field it terminates is the last field in the text.
'#
'#      Field #1 index is always the same as startindex--even if the character at
'#      the startindex position is a delimiter.
'#
'# PARAMETERS:
'#
'#      startindex      Long (1-based) start position for field count in
'#                      text.  If startIndex < 1 then we will use
'#                      startindex = 1.
'#
'#      text            String to count fields beginning at startindex.
'#
'#      delimiter       Delimiter string.  The entire string is one delimiter.
'#
'# RETURNS:
'#
'#      0   Indicates bad input (like 0-length text or a startindex
'#          greater than the text length).
'#
'#      N   (Integer) The number of fields found beginning at startindex.
'#
'# ERRORS:
'#
'#      none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: JAN 16, 2003
'# History:
'#
'#      JAN 16, 2003    Original Release
'#
'############################################################################
Public Function GetMultiDelimitedFieldCount(startindex As Long, text As String, _
                                            delimiter As String) As Integer

    Dim fieldIndex As Long
    Dim nextIndex As Long
    Dim lenDelimiter As Long
    Dim lenText As Long
    Dim fieldCount As Integer
    
    lenText = Len(text)
    lenDelimiter = Len(delimiter)
    
    If (lenText = 0) Or (lenDelimiter = 0) Then Exit Function
    
    fieldIndex = startindex
    If fieldIndex < 1 Then fieldIndex = 1
    If fieldIndex > lenText Then Exit Function
    
    fieldCount = 0
    
    Do
        fieldCount = fieldCount + 1
        nextIndex = InStr(fieldIndex, text, delimiter, 0)
            
        If nextIndex > 0 Then fieldIndex = nextIndex + lenDelimiter
    
    Loop While ((fieldIndex <= lenText) And (nextIndex > 0))
        
    GetMultiDelimitedFieldCount = fieldCount
        
End Function


'############################################################################
'# <a name="getmultidelimitedfield"></a>
'#
'#  Function GetMultiDelimitedField (startIndex as Long, text As String,
'#                     delimiter As String, fieldID As Integer) As String
'#
'# DESCRIPTION:
'#
'#      Given text of delimited fields, return the nth(FieldID) Field
'#      in the record from startIndex.  The index is 1 based.  The delimiter
'#      string can contain multiple characters and this string is treated
'#      as a single delimiter.
'#
'#      For example, if delimiter = "->" then the fields of the text string
'#      are expected to be delimited by the full delimiter string as in:
'#
'#      "field1->field2->field3"
'#
'#
'# PARAMETERS:
'#
'#      startIndex      the one based index to begin search.  if < 1 then we
'#                      will start at index = 1.
'#
'#      text            string of delimited fields
'#
'#      delimiter       the string (>0 characters) that delimits each field.
'#                      As indicated this can be more than one character in
'#                      length.  The entire string counts as one delimiter.
'#
'#      fieldID         the one based index of the field to find and return. If < 1
'#                      then we will use fieldID = 1.
'#
'#
'# RETURNS:
'#
'#      A string of the indexed field.
'#      An empty string ("") is returned if the fieldID is out of range for
'#      the provided text or the provided delimiter is invalid.
'#
'# ERRORS:
'#
'#       none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: JAN 16, 2003
'# History:
'#
'#      JAN 16, 2003    Original Release
'#
'############################################################################
Public Function GetMultiDelimitedField(startindex As Long, text As String, _
                         delimiter As String, fieldID As Integer) As String
    
    Dim fieldIndex As Long
    Dim nextIndex As Long
    Dim lenDelimiter As Long
    Dim lenText As Long
    Dim lenField As Long
    Dim fieldCount As Integer
    Dim fieldTarget As Integer
    
    lenText = Len(text)
    lenDelimiter = Len(delimiter)
    
    fieldTarget = fieldID
    If fieldTarget < 1 Then fieldTarget = 1
    If (lenText = 0) Or (lenDelimiter = 0) Then Exit Function
    
    fieldIndex = startindex
    If fieldIndex < 1 Then fieldIndex = 1
    If fieldIndex > lenText Then Exit Function
    
    fieldCount = 1
    
    Do
        nextIndex = InStr(fieldIndex, text, delimiter, 0)
        
        If nextIndex < 1 Then      'no delimiter found
            If fieldCount = fieldTarget Then
                lenField = lenText - fieldIndex + 1
                GetMultiDelimitedField = Mid$(text, fieldIndex, lenField)
            End If
            Exit Function
        Else                        'delimiter was found
            If fieldCount = fieldTarget Then
                lenField = nextIndex - fieldIndex
                GetMultiDelimitedField = Mid$(text, fieldIndex, lenField)
                Exit Function
            Else
                fieldCount = fieldCount + 1
                fieldIndex = nextIndex + lenDelimiter
            End If
        End If
        
    Loop While ((fieldCount <= fieldTarget) And (fieldIndex <= lenText))
        
End Function


Private Function optGetFieldArray(startindex As Long, text As String, lenText As Long, _
                        delimiters As String, lenDelimiters As Long, fields() As String) As Integer
    
    Dim count As Integer
    Dim index As Integer
    Dim sindex As Long
    Dim eindex As Long
    Dim nodelimit As Integer
    Dim field As String
    
    
    count = optGetFieldCount(startindex, text, lenText, delimiters, lenDelimiters)
    'If count = 0 Then Exit Function  'don't think this is possible from optGetFieldCount

    ReDim fields(1 To count) As String
    sindex = startindex
    
    For index = 1 To count
            
        eindex = optGetDelimiterIndex(sindex, text, lenText, delimiters, lenDelimiters)
        
        'if no delimiters then set to end of string
        'this can only be on the LAST field
        If eindex < sindex Then
            eindex = lenText
            nodelimit = 1
        End If
        
        If sindex = eindex Then
            'if also the last char in the string
            If nodelimit Then
                field = Mid$(text, sindex, 1)
            'otherwise it is just an empty field
            Else
                field = ""
            End If
        Else
            'normal extraction of the field
            field = Mid$(text, sindex, eindex - sindex + nodelimit)
        End If
        
        fields(index) = field
        sindex = eindex + 1

    Next
    
    optGetFieldArray = count
    
End Function


'############################################################################
'#
'#  Function GetFieldArray (startindex As Long, text As String, _
'#                          delimiters As String, fields() As String) As Integer
'#
'# DESCRIPTION:
'#
'#      Process the inputRecord and fill the provided array with all the
'#      fields in the record subsequent to startindex.  To process the
'#      entire inputRecord into an array of fields use startIndex = 1.
'#
'#      Blank fields count, so finding any one delimiter in our delimiters string
'#      will indicate the end of a field. The next field starts at the next
'#      character index even if that char is also a delimiter.  Back-to-back
'#      delimiters indicate an empty field.
'#
'#      While blank fields are allowed, if the entire text string is empty,
'#      then we exit with no found fields (0).
'#
'#      Fields are extracted unmodified.
'#
'#      If the last character in the inputRecord is a delimiter, then the
'#      field it terminates is the last field in the inputRecord.
'#
'#
'# PARAMETERS:
'#
'#      startindex      Integer (1-based) start character position for field count in
'#                      inputRecord. If startindex < 1 then we will use
'#                      startindex = 1.
'#
'#      text            String to parse for fields beginning at startindex.
'#
'#      delimiters      String list of possible delimiters that separate fields.
'#                      If there are no delimiters, then there is only 1 field.
'#
'#      array           String array to populate with the field data.
'#                      Should be allocated or dimensioned as:
'#
'#                         Dim array() As String
'#
'#                      This routine will ReDim the array appropriately.
'#                      The returned array will be 1-based.  That is, the
'#                      first element of the array will be at index = 1.
'#
'#      (NOTE:Not sure if this will work passing arrays outside of the DLL)
'#      (Seems SQABasic won't even let me attempt to pass an array!)
'#
'#
'# RETURNS:
'#
'#      0   Indicates bad input (like 0-length inputRecord or a startindex
'#          greater than the record length).
'#
'#      N   The number of fields returned in the array.
'#
'# ERRORS:
'#
'#      none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: SEP 28, 2001
'# History:
'#
'#      SEP 28, 2001    Original Release
'#
'############################################################################
Public Function GetFieldArray(startindex As Long, text As String, _
                        delimiters As String, fields() As String) As Integer
    
    Dim lenText As Long
    Dim lenDelimiters As Long
    
    lenText = Len(text)
    If lenText = 0 Then Exit Function
    If startindex > lenText Then Exit Function
    lenDelimiters = Len(delimiters)
    If startindex < 1 Then startindex = 1
    
    GetFieldArray = optGetFieldArray(startindex, text, lenText, delimiters, lenDelimiters, fields)
    
End Function



'############################################################################
'# <a name="gettrimmedfield"></a>
'#
'#  Function GetTrimmedField (startIndex as Long, text As String,
'#                            delimiters As String, fieldID As Integer) As String
'#
'# DESCRIPTION:
'#
'#      Given text of delimited fields, return the nth(FieldID) Field
'#      in the record from startIndex.  The index is 1 based.  The delimiters
'#      string can contain multiple characters but each character is individually
'#      treated as a possible delimiter.  The field is trimmed of leading and
'#      trailing whitespace.
'#
'# PARAMETERS:
'#
'#      startIndex      the one based index to begin search.  if < 1 then we
'#                      will start at index = 1.
'#      text            string of delimited fields
'#      delimiters       the string (>0 characters) that delimits each field.
'#                      As indicated this can be more than one character in
'#                      length--each character is separately a possible delimiter.
'#      fieldID         the one based index of the field to find and return.
'#
'# RETURNS:
'#
'#      A string of the indexed field trimmed of whitespace.
'#      An empty string ("") is returned if the fieldID is out of range for
'#      the provided text or the provided delimiters are invalid.
'#
'# ERRORS:
'#
'#       none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: Aug 15, 2001
'# History:
'#
'#      Aug 15, 2001    Original Release
'#
'############################################################################
Public Function GetTrimmedField(startindex As Long, text As String, _
                         delimiters As String, fieldID As Integer) As String
    
    Dim fieldVal As String
    
    fieldVal = GetField(startindex, text, delimiters, fieldID)
    If Len(fieldVal) > 0 Then GetTrimmedField = TWhitespace(fieldVal)
    
End Function


'############################################################################
'# <a name="getquotedfield"></a>
'#
'#  Function GetQuotedField(startindex as Long, text As String,
'#                          delimiters As String, fieldID As Integer) As String
'#
'# DESCRIPTION:
'#
'#      Given an text of delimited fields, return the nth(FieldID) Field
'#      in the record.  The index is 1 based.  The field is processed with
'#      <a href="#GetQuotedString">GetQuotedString</a> before it is returned.
'#
'# PARAMETERS:
'#
'#      startIndex      the one based index to begin search.  if < 1 then we
'#                      will start at index = 1.
'#      text            string of delimited fields
'#      delimiter       the string (>0 characters) that delimits each field.
'#      fieldID         the one based index of the desired field
'#
'# RETURNS:
'#
'#      A string of the indexed field processed with <a href="#GetQuotedString">GetQuotedString</a>.
'#      An empty string ("") is returned if the fieldID is out of range for
'#      the provided text or the provided delimiter is invalid.
'#
'# ERRORS:
'#
'#       none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: APR 24, 2000
'# History:
'#
'#      APR 24, 2000    Original Release
'#      JUL 22, 2000    (Carl Nagle) Added error trapping.
'#      AUG 16, 2001    Converted to DLL use
'#
'############################################################################
Public Function GetQuotedField(startindex As Long, text As String, _
                         delimiters As String, fieldID As Integer) As String
    
    Dim fieldVal As String
        
    fieldVal = GetField(startindex, text, delimiters, fieldID)
    GetQuotedField = GetQuotedString(fieldVal)
    
End Function


Private Function optGetDelimiterIndexRev(startindex As Long, text As String, lenText, _
                                         delimiters As String, lenDelimiters As Long) As Long
    Dim midText As String
    Dim textIndex As Long
    Dim delimiterIndex As Long
    Dim done As Boolean
    
    optGetDelimiterIndexRev = -1
    
    done = False
    textIndex = startindex
    Do While (Not done) And (textIndex >= 1)
        delimiterIndex = 1
        midText = Mid$(text, textIndex, 1)
        Do While (Not done) And (delimiterIndex <= lenDelimiters)
            If midText = Mid$(delimiters, delimiterIndex, 1) Then
                done = True
            Else
                delimiterIndex = delimiterIndex + 1
            End If
        Loop
        If Not done Then textIndex = textIndex - 1
    Loop
    If done Then optGetDelimiterIndexRev = textIndex
End Function


'############################################################################
'# <a name="getdelimiterindexrev"></a>
'#
'#  Function GetDelimiterIndexREV (startindex As Long, text As String,
'#                                  delimiters As String) As Long
'#
'# DESCRIPTION:
'#
'#      Finds the index of the first character matching one of the provided delimiter
'#      characters.  The search begins at startindex within the text and
'#      goes BACKWARDS.
'#      The index is 1-based.
'#
'#
'# PARAMETERS:
'#
'#      startindex      Integer (1-based) start position for search in text
'#                      If startindex is < 1 then we use startindex = the last
'#                      character in the text string.
'#
'#      text            String to search for first delimiter character found
'#                      during the backwards search.
'#
'#      delimiters      String list of delimiters to find.
'#
'#
'# RETURNS:
'#
'#      -1   on error, invalid parameters, or end of text reached.
'#      N   index of first delimiter character found.
'#
'#
'# ERRORS:
'#
'#      none
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function GetDelimiterIndexRev(startindex As Long, text As String, delimiters As String) As Long
    
    Dim lenText As Long
    Dim lenDelimiters As Long
    
    GetDelimiterIndexRev = -1
    lenText = Len(text)
    If lenText = 0 Then Exit Function
    lenDelimiters = Len(delimiters)
    If lenDelimiters = 0 Then Exit Function
    If startindex < 1 Then startindex = lenText
    If startindex > lenText Then Exit Function
    
    GetDelimiterIndexRev = optGetDelimiterIndexRev(startindex, text, lenText, delimiters, lenDelimiters)
    
End Function


Private Function optGetUnquotedDelimiterIndexRev(startindex As Long, _
                        text As String, lenText As Long, _
                        delimiters As String, lenDelimiters As Long) As Long
    Dim index As Long
        
    index = startindex
    Do
        index = optGetDelimiterIndexRev(index, text, lenText, delimiters, lenDelimiters)
        If index > 0 Then
            If IsQuoted(text, index) Then
                index = index - 1
            Else
                Exit Do
            End If
        End If
    Loop Until index < 1
    If index < 1 Then index = -1
    optGetUnquotedDelimiterIndexRev = index
End Function


'############################################################################
'# <a name="getunquoteddelimiterindexrev"></a>
'#
'#  Function GetUnquotedDelimiterIndexREV (startindex As Integer, text As String,
'#                                  delimiters As String) As Long
'#
'# DESCRIPTION:
'#
'#      Finds the index of the first character matching one of the provided delimiter
'#      characters that is NOT inside a quoted literal string.  The search
'#      begins at startindex within the text and goes BACKWARDS.  Essentially,
'#      this is the same as calling GetDelimiterIndexREV followed by IsQuoted
'#      and only returning the index if the delimiter is NOT quoted.
'#      The index is 1-based.
'#
'#
'# PARAMETERS:
'#
'#      startindex      Integer (1-based) start position for search in text
'#                      If startindex is < 1 then we use startindex = the last
'#                      character in the text string.
'#
'#      text            String to search for first unquoted delimiter character
'#
'#      delimiters      String list of delimiters to find.
'#
'#
'# RETURNS:
'#
'#      -1   on error, invalid parameters, or end of text reached.
'#      N   index of first delimiter character found.
'#
'#
'# ERRORS:
'#
'#      none
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function GetUnquotedDelimiterIndexRev(startindex As Long, text As String, delimiters As String) As Long
    
    Dim lenText As Long
    Dim lenDelimiters As Long
    
    GetUnquotedDelimiterIndexRev = -1
    lenText = Len(text)
    If lenText = 0 Then Exit Function
    lenDelimiters = Len(delimiters)
    If lenDelimiters = 0 Then Exit Function
    If startindex < 1 Then startindex = lenText
    If startindex > lenText Then Exit Function
    
    GetUnquotedDelimiterIndexRev = optGetUnquotedDelimiterIndexRev(startindex, text, _
                                        lenText, delimiters, lenDelimiters)
End Function


'############################################################################
'# <a name="gettrimmedquotedstring"></a>
'#
'#  Function GetTrimmedQuotedString (text as String) As String
'#
'# DESCRIPTION:
'#
'#      Trim leading and trailing whitespace from the text string AND remove
'#      the outermost double-quote marks in the text, if they exist.
'#
'#
'# PARAMETERS:
'#
'#      text    The text string to trim and remove quotes.
'#
'#
'# RETURNS:
'#
'#      Text trimmed of leading and trailing whitespace but retaining
'#      whatever exists between a leading and/or trailing double-quote mark.
'#      Text not containing leading or trailing quote marks (after trimming)
'#      will be returned merely trimmed of leading and trailing whitespace.
'#
'#
'# ERRORS:
'#
'#      none
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function GetTrimmedQuotedString(text As String) As String
    Dim strText As String
    
    strText = TWhitespace(text)
    If Left$(strText, 1) = gQT Then strText = Mid$(strText, 2)
    If Right$(strText, 1) = gQT Then strText = Left$(strText, Len(strText) - 1)
    GetTrimmedQuotedString = strText
    
End Function


'############################################################################
'# <a name="gettrimmedquotedfield"></a>
'#
'#  Function GetTrimmedQuotedField(startindex As Long, text As String,
'#                          delimiters As String, fieldID As Integer) As String
'#
'# DESCRIPTION:
'#
'#      Given an text of delimited fields, return the nth(FieldID) Field
'#      in the record.  The index is 1 based.  The field is processed with
'#      <a href="#GetTrimmedQuotedString">GetTrimmedQuotedString</a> before it is returned.
'#
'# PARAMETERS:
'#
'#      startIndex      the index to begin at. If < 1 then we will start at 1.
'#                      Index is 1 based.
'#      text            string of delimited fields
'#      delimiters       the string (>0 characters) that delimits each field.
'#      fieldID         the one based index of the desired field
'#
'# RETURNS:
'#
'#      The processed string of the indexed field.
'#      An empty string ("") is returned if the fieldID is out of range for
'#      the provided text or the provided delimiter is invalid.
'#
'# ERRORS:
'#
'#       none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function GetTrimmedQuotedField(startindex As Long, text As String, _
                         delimiters As String, fieldID As Integer) As String
    
    Dim fieldVal As String
    
    fieldVal = GetField(startindex, text, delimiters, fieldID)
    
    GetTrimmedQuotedField = GetTrimmedQuotedString(fieldVal)
    
End Function


'extract variable value if text is a variable.
'quote the whole returned string IF it is a variable value.
'(Carl Nagle) NOV 03, 2003 Fixed bug where a variable with no stored value or constant
'                      was returned with a string value of the variable expression.
'                      EX: ^NoVal was receiving and returned with value "^NoVal"
Private Function QTTextOfVariable(text As String, store As Object) As String
    Dim strTemp As String
    Dim status As Integer
    Dim varvalue As Variant
    
    strTemp = ValidateVariableID(text)
    
    'if it is a valid variable name, get the value
    If Len(strTemp) > 0 Then
        status = store.GetVariableValue(strTemp, varvalue)
        
        'if variable never initialized status <> 0
        If status <> 0 Then
            text = ""
        Else
            text = gQT & varvalue & gQT
        End If
    'not a variable, so we don't quote the trimmed string
    Else
        text = TWhitespace(text)
    End If
    
    
    QTTextOfVariable = text
End Function


'############################################################################
'# <a name="findandreplace"></a>
'#
'#  Function FindAndReplace(source As String, find As String,
'#                          replace As String, Optional noCase as Integer = 0) As String
'#
'# DESCRIPTION:
'#
'#      Search the source string for all instances of the find string and
'#      replace all instances with the replace string.  This is done on a
'#      copy of the source string.  The routine returns this new string on
'#      exit.  All strings are treated as case-sensitive unless the optional
'#      noCase parameter is provided and is <> 0.
'#
'#      NOTE: the routine can be used to only delete unwanted find strings
'#      by providing a zero-length ("") replacement string.
'#
'# PARAMETERS:
'#
'#      source      the String to copy and do find and replace on.
'#                  It is not modified.
'#
'#      find        string to locate for replacement within source.
'#
'#      replace     string to substitute in place of find string.
'#
'#      noCase      (Optional) Set <> 0 to use case-insensitive searches.
'#
'# RETURNS:
'#
'#      New String after FindAndReplace performed.  The input source string
'#      is not modified.
'#
'#      If either the source or the find strings are invalid ("") then the
'#      return value is "".
'#
'# ERRORS:
'#
'#       none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#      APR 29, 2002    (Carl Nagle) Added the noCase option.
'#
'############################################################################
Public Function FindAndReplace(source As String, find As String, replace As String, Optional noCase As Integer = 0) As String

    FindAndReplace = ""
    
    If source = "" Then Exit Function
    If find = "" Then Exit Function
        
    Dim workstring As String
    Dim lenfind    As Integer
    Dim lenreplace As Integer
    Dim lensource  As Integer
    Dim start As Integer
    Dim index As Integer
    
    workstring = ""
    lensource = Len(source)
    lenfind = Len(find)
    lenreplace = Len(replace)
    If noCase <> 0 Then noCase = 1
    
    start = 1
    Do
        index = InStr(start, source, find, noCase)
        If index = 0 Then
            workstring = workstring & Right$(source, lensource - start + 1)
            start = lensource + 1
        Else
            workstring = workstring & Mid$(source, start, index - start) & replace
            start = index + lenfind
        End If
    Loop While start <= lensource
    
    FindAndReplace = workstring
        
End Function


'############################################################################
'# <a name="getnextnonwhitespacecharacterindex"></a>
'#
'#  Function GetNextNonWhitespaceCharacterIndex (startindex As Integer,
'#                                              text As String) As Integer
'#
'# DESCRIPTION:
'#
'#      Finds the index of the first character that is not a whitespace
'#      character.  The search begins at startindex within the text.
'#      The index is 1-based.
'#
'#
'# PARAMETERS:
'#
'#      startindex      Integer (1-based) start position for search in text.
'#                      If startindex is < 1 then we use startindex = 1.
'#
'#      text            String to search for first non-whitespace character
'#
'#
'# RETURNS:
'#
'#      -1   on error, invalid parameters, or end of text reached.
'#      N   index of first non-whitespace character found.
'#
'#
'# ERRORS:
'#
'#      none
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function GetNextNonWhitespaceCharacterIndex(startindex As Long, text As String) As Long
    GetNextNonWhitespaceCharacterIndex = -1

'    If startIndex < 1 Then Exit Function
    If startindex < 1 Then startindex = 1
    
    Do Until (startindex > Len(text))
        If InStr(gWhitespace, Mid$(text, startindex, 1)) < 1 Then
            GetNextNonWhitespaceCharacterIndex = startindex
            Exit Do
        End If
        startindex = startindex + 1
    Loop
    
End Function


'############################################################################
'# <a name="getnextnonwhitespacecharacterindexrev"></a>
'#
'#  Function GetNextNonWhitespaceCharacterIndexREV (startindex As Integer,
'#                                              text As String) As Integer
'#
'# DESCRIPTION:
'#
'#      Finds the index of the first character that is not a whitespace
'#      character.  The search begins at startindex within the text and
'#      moves BACKWARDS.
'#      The index is 1-based.
'#
'#
'# PARAMETERS:
'#
'#      startindex      Integer (1-based) start position for search in text.
'#                      If startindex is < 1 then we use startindex = the last
'#                      character in the string.
'#
'#      text            String to search for first non-whitespace character
'#
'#
'# RETURNS:
'#
'#      -1   on error, invalid parameters, or end of text reached.
'#      N   index of first non-whitespace character found.
'#
'#
'# ERRORS:
'#
'#      none
'#
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function GetNextNonWhitespaceCharacterIndexREV(startindex As Long, text As String) As Long
    GetNextNonWhitespaceCharacterIndexREV = -1
    
    If startindex < 1 Then startindex = Len(text)
    
    Do Until (startindex < 1)
        If InStr(gWhitespace, Mid$(text, startindex, 1)) < 1 Then
            GetNextNonWhitespaceCharacterIndexREV = startindex
            Exit Do
        End If
        startindex = startindex - 1
    Loop
    
End Function


Private Function optGetNextNonBlankFieldIndex(startindex As Long, _
                                text As String, lenText As Long, _
                                delimiters As String, lenDelimiters As Long) As Long

    Dim lI1 As Long
    Dim lI2 As Long
    Dim isDelimiter As Integer
    Dim fieldTest As String
    Dim endindex As Long
        
    'explicit set for failure
    optGetNextNonBlankFieldIndex = -1
    
    lI1 = startindex
    
    'determine if lI1 itself is a delimiter
    If InStr(1, delimiters, Mid$(text, lI1, 1), vbBinaryCompare) > 0 Then
        isDelimiter = 1
    Else
        isDelimiter = 0
    End If
    
    Do
        'find next delimiter
        lI2 = optGetDelimiterIndex(lI1 + 1, text, lenText, delimiters, lenDelimiters)
        
        'set the appropriate endIndex based on
        'if we found a delimiter or not
        If lI2 >= lI1 + 1 Then
            endindex = lI2
        Else
            endindex = lenText
        End If
            
        'extract the field
        fieldTest = Mid$(text, lI1 + isDelimiter, endindex - lI1 - isDelimiter)
        
        'see if we have a non-blank field
        fieldTest = LTWhitespace(fieldTest)
        
        'if it is non-blank
        If Len(fieldTest) > 0 Then
            optGetNextNonBlankFieldIndex = lI1 + isDelimiter
            Exit Do
        
        'otherwise set lI1 to lI2 index (or end of text)
        Else
            lI1 = endindex
            isDelimiter = 1     'won't matter if we are at end of text
        End If
                    
    'if no endIndex delimiter was found, we exit because there are
    'no more fields to check
    Loop Until lI2 < lI1

End Function


'############################################################################
'# <a name="getnextnonblankfieldindex"></a>
'#
'#  Function GetNextNonBlankFieldIndex (startindex As Integer, text As String,
'#                                      delimiters As String) As Long
'#
'# DESCRIPTION:
'#
'#      This routine skips over blank fields looking for the index of the
'#      first non-blank field.
'#
'#      The search begins at startindex within the text.
'#      The index is 1-based.
'#
'# PARAMETERS:
'#
'#      text            String to search for first non-blank field
'#
'#      startindex      Integer (1-based) start position for search in text.
'#                      if < 1 then start at 1
'#
'#      delimiters      String list of delimiters to use for fields.
'#
'# RETURNS:
'#
'#     -1   on error, invalid parameters, or end of text reached.
'#      N   index of first non-blank field found.
'#
'# ERRORS:
'#
'#      none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function GetNextNonBlankFieldIndex(startindex As Long, text As String, _
                            delimiters As String) As Long

    Dim lenText As Long, lenDelimiters As Long
        
    'explicit set for failure
    GetNextNonBlankFieldIndex = -1
    
    lenDelimiters = Len(delimiters)
    If lenDelimiters = 0 Then Exit Function
    lenText = Len(text)
    If lenText = 0 Then Exit Function
    If startindex > lenText Then Exit Function
    If startindex < 1 Then startindex = 1
    
    GetNextNonBlankFieldIndex = optGetNextNonBlankFieldIndex(startindex, text, _
                                                    lenText, delimiters, lenDelimiters)

End Function


'############################################################################
'# <a name="getnonblankfield"></a>
'#
'#  Function GetNonBlankField(startindex as Long, text As String,
'#                            delimiters As String, fieldID as Integer) As String
'#
'# DESCRIPTION:
'#
'#      Given text of delimited fields, return the nth(FieldID)
'#      non-blank field in the record.  The index is 1 based.
'#      Thus if you have back-to-back delimiters they are ignored and instead of
'#      returning "" we continue our field search until the nth non-blank field
'#      is found.
'#      One example would be fields that are space delimited and multiple (and
'#      variable numbers of) spaces separate the fields.
'#
'#      Use GetField when back-to-back delimiters are allowed indicating
'#      a desirable zero-length (missing) string field.
'#
'# PARAMETERS:
'#
'#      startindex      begin text search at 1-based startindex
'#      text            string of delimited fields
'#      fieldID         the 1-based index of the desired non-blank field
'#      delimiters      the string (>0 characters) that delimits each field.
'#                      As indicated this can be more than one character in
'#                      length.
'#
'# RETURNS:
'#
'#      A string of the indexed field trimmed of leading and trailing spaces.
'#      An empty string ("") is returned if the fieldID is out of range for
'#      the provided text or the provided delimiter is invalid.
'#
'# ERRORS:
'#
'#       none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function GetNonBlankField(startindex As Long, text As String, _
                         delimiters As String, fieldID As Integer) As String

    'default return code GetNonBlankField = ""
    
    Dim beginindex As Long
    Dim endindex As Long
    Dim tempIndex As Long
    Dim position As Long
    Dim count As Integer
    Dim lenText As Long, lenDelimiters As Long
    
    'If fieldID < 1 Then Exit Function
    If fieldID < 1 Then fieldID = 1
    
    lenDelimiters = Len(delimiters)
    If lenDelimiters = 0 Then Exit Function
    lenText = Len(text)
    If lenText = 0 Then Exit Function
    If startindex > lenText Then Exit Function
    If startindex < 1 Then startindex = 1
    
    
    endindex = startindex - 1
    position = 1
    Do While position < lenText
        
        'find the start of the next field
        beginindex = endindex + 1
        tempIndex = optGetNextNonBlankFieldIndex(beginindex, text, lenText, delimiters, lenDelimiters)
        If tempIndex >= beginindex Then
            beginindex = tempIndex
            position = tempIndex
            count = count + 1
            
            'now find the end of the field
            tempIndex = optGetDelimiterIndex(beginindex, text, lenText, delimiters, lenDelimiters)
            If tempIndex > beginindex Then
                endindex = tempIndex
                position = tempIndex
            Else
                endindex = lenText + 1
                position = endindex
            End If
            
            'exit if we are on the right fieldID
            If count = fieldID Then
                GetNonBlankField = Mid$(text, beginindex, endindex - beginindex)
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
End Function


'############################################################################
'# <a name="getliteralquotedfield"></a>
'#
'#  Function GetLiteralQuotedField (startindex As Long, text As String,
'#                          delimiters As String, fieldID As Integer) As String
'#
'# DESCRIPTION:
'#
'#      Return the nth field from the text parameter.
'#      Parse the field as a literal quoted string regardless of embedded
'#      delimiters.  The search and field count begins at startindex within
'#      the text.  The field and index are 1-based using GetFieldIndex
'#      to locate the beginning of the nth field.
'#
'#      Blank fields count, so finding any one delimiter in our delimiters string
'#      will indicate the end of a field. The next field starts at the next
'#      character index even if that char is also a delimiter.
'#
'#      A literal quoted string is defined by a leading double-quote and a
'#      trailing double-quote irregardless of any field separator currently in
'#      force.  To include a double-quote within this string two double-quote
'#      characters must exist with no intervening characters and neither of these
'#      can be the outermost leading or trailing double-quote.
'#
'#      Field #1 is always the same as startindex--even if the character at
'#      the startindex position is a delimiter.
'#
'#      The outermost quotes will not be returned as part of the string.
'#      A leading quote with no field ending quote will return everything
'#      after the leading quote all the way to the end of the text.
'#
'#      Quoted String Examples (assume comma separated record fields):
'#
'#           Input:"Every day I say, ""Hello."""
'#          Output:Every day I say, "Hello."
'#          (note the 3 quotes at the end)
'#
'#           Input: Enter this: "Every day I say, ""Hello."""
'#          Output:Every day I say, "Hello."
'#          (note initial "Enter this: " is discarded)
'#
'#           Input: Enter this: "Every day I say, ""Hello.""" then exit
'#          Output:Every day I say, "Hello."
'#          (note initial "Enter this: " AND trailing " then exit" is discarded)
'#
'#      Note: Even though the record (not shown) is supposedly a group of comma
'#      separated fields, this Get routine ignores the comma in the text and
'#      continues on until the closing quote mark.
'#
'#
'# PARAMETERS:
'#
'#      startindex      Integer (1-based) start position for text field
'#                      search.
'#
'#      text            record to extract field from.
'#
'#      delimiters      String list of delimiters that separate fields.  This is
'#                      used to find the start of our field and then ignored.
'#                      If fieldID = 1 then delimiters can be any string or empty
'#                      since it will be ignored.
'#
'#      fieldID         Integer (1-based) of field # to find.
'#                      field #1 begins at startindex regardless of the true
'#                      start of the text.  If the field # is not found
'#                      to exist (due to record length) then it is treated as
'#                      an empty (optional?) field ("").
'#                      A fieldID < 1 is an error.
'#
'#
'# RETURNS:
'#
'#      The text inside the literal quoted string with all embedded quotes
'#      converted to single double-quotes.
'#
'# ERRORS:
'#
'#      none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function GetLiteralQuotedField(startindex As Long, text As String, _
                         delimiters As String, fieldID As Integer) As String

    Dim fIndex As Long, q1 As Long, q2 As Long
    Dim tempRecord As String
    Dim lenText As Long, lenDelimiters As Long
    
    'default: invalid parameters
    'If fieldID < 1 Then Exit Function
    If fieldID < 1 Then fieldID = 1
    lenDelimiters = Len(delimiters)
    If lenDelimiters = 0 Then If fieldID > 1 Then Exit Function
    lenText = Len(text)
    If lenText = 0 Then Exit Function
    If startindex > lenText - 1 Then Exit Function
    If startindex < 1 Then startindex = 1
    
    'finds the true starting index of our field n
    fIndex = optGetFieldIndex(startindex, text, lenText, delimiters, lenDelimiters, fieldID)
    If fIndex < 1 Then Exit Function
    
    'cut down the string to start at fIndex
    tempRecord = Mid$(text, fIndex)
    
    'find the first quote mark index q1
    q1 = InStr(tempRecord, gQT)
    
    'if not a quoted string then exit with error
    If (q1 < 1) Or (q1 = Len(tempRecord)) Then Exit Function

    'cut down the string to start at the quote
    tempRecord = Mid$(tempRecord, q1)
    
    'find the terminating quote mark
    q2 = optGetUnquotedDelimiterIndex(2, tempRecord, Len(tempRecord), gQT, 1)
    
    'if none, the rest of the string is included
    If q2 < 1 Then
        tempRecord = Mid$(tempRecord, 2)
    Else
        tempRecord = Mid$(tempRecord, 2, q2 - 2)
    End If
    
    'reduce double double-quotes to single double-quotes
    GetLiteralQuotedField = FindAndReplace(tempRecord, gQT & gQT, gQT)
    
End Function


'############################################################################
'# <a name="getliteralquotedrecordfield"></a>
'#
'#  Function GetLiteralQuotedRecordField (text As String, fieldID As Integer) As String
'#
'# DESCRIPTION:
'#
'#      Return the nth field from the text.
'#      Parse ALL fields as Literal Quoted Strings using GetLiteralQuotedField.
'#      This means that anything can separate the fields as long as their is
'#      at least one non-quote character between the end of one field and the
'#      beginning of the next.  No characters inside the Literal Quoted Strings
'#      are treated as delimiters (except the leading and trailing quote that
'#      identifies the beginning and ending of the field).
'#
'#      A Literal Quoted String is defined by a leading double-quote and a
'#      trailing double-quote irregardless of any field separator currently in
'#      force.  To include a double-quote within this string two double-quote
'#      characters must exist with no intervening characters and neither of these
'#      can be the outermost leading or trailing double-quote.
'#
'#      The outermost quotes will not be returned as part of the string.
'#      A leading quote with no field ending quote will return everything
'#      after the leading quote all the way to the end of the text.
'#
'#      LiteralQuotedString Examples:
'#
'#          FIELD #1        FIELD #2            FIELD #3
'#      ----------------    ----------------    -------------------
'#      "This, is Field1" : "Then, Field #2" ,, "Lastly, comes Field #3"
'#
'#      Note the colons and commas between the fields?  They are ignored since
'#      they are not within a quoted string.
'#
'#
'# PARAMETERS:
'#
'#      text            Literal Quoted String record to extract field from.
'#
'#      fieldID         Integer (1-based) of field # to find.
'#                      field #1 begins at index 1 of the text.
'#                      If the field # is not found to exist (due to record length)
'#                      then it is treated as an empty (optional?) field ("").
'#                      A fieldID < 1 is an error.
'#
'#
'# RETURNS:
'#
'#      The text inside literal quoted field N with all embedded quotes
'#      converted to single double-quotes.  See the examples in
'#      GetLiteralQuotedField.
'#
'#
'# ERRORS:
'#
'#      none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function GetLiteralQuotedRecordField(text As String, fieldID As Integer) As String

    Dim fIndex As Integer, q1 As Long, q2 As Long
    Dim tempRecord As String
    
    If Len(text) = 0 Or fieldID < 1 Then Exit Function
    
    tempRecord = text
    
    Do
        'increment for which field we are looking for
        fIndex = fIndex + 1
        
        'find the first quote mark index q1
        q1 = InStr(tempRecord, gQT)
        
        'if not a quoted string then exit with error
        If (q1 < 1) Or (q1 = Len(tempRecord)) Then Exit Function
    
        'cut down the string to start at the quote
        tempRecord = Mid$(tempRecord, q1)
        
        'find the terminating quote mark
        q2 = optGetUnquotedDelimiterIndex(2, tempRecord, Len(tempRecord), gQT, 1)
        
        'exit if fieldID is not met by end of text
        If (q2 < 1) And (fIndex < fieldID) Then Exit Function
        
        'if we are at end of text then we are also at fieldID
        If q2 < 1 Then Exit Do
                
        'cut down the string removing the current quoted field
        If fIndex < fieldID Then tempRecord = Mid$(tempRecord, q2 + 1)
    
    Loop While fIndex < fieldID
    
    'if none, the rest of the string is included
    If q2 < 1 Then
        tempRecord = Mid$(tempRecord, 2)
    Else
        tempRecord = Mid$(tempRecord, 2, q2 - 2)
    End If
    'reduce double double-quotes to single double-quotes
    GetLiteralQuotedRecordField = FindAndReplace(tempRecord, gQT & gQT, gQT)
    
End Function


'############################################################################
'# <a name="getquotedstring"></a>
'#
'#  Function GetQuotedString (sInput as String) As String
'#
'# DESCRIPTION:
'#
'#      This routine is used to extract possibly quoted text from within the provided
'#      text string.  It first trims the input text of any leading or trailing
'#      whitespace looking for leading or trailing double-quotes.  If the trimmed text
'#      contains leading or trailing quotes then the routine removes a single leading
'#      and/or trailing double-quote mark from the text (whichever exist).
'#
'#      If there are no leading or trailing quote marks then this routine merely
'#      returns the untrimmed text with all whitespace intact.
'#
'#      The purpose for this is to allow leading and trailing whitespace to exist
'#      in a string and extract only the text from within the quote marks (which
'#      might also include required leading and trailing whitespace).  It also allows
'#      us to retain leading and trailing spaces on input text not containing leading
'#      or trailing quote marks (for backward compatibility).
'#
'#      (Note: that in the examples below the outermost quotes shown on the
'#      return value are shown only to convey the string start and stop points.
'#      They are not actually returned as part of the return value.)
'#
'#               ----------------  returns:-----------------
'#      Examples:   text field             "   text field"
'#                 "text field"            "text field"
'#                     "  text  "          "  text  "
'#
'#      Quote marks other than the leading and/or trailing quote mark are
'#      retained.
'#
'#
'# PARAMETERS:
'#
'#      sInput    The input string to trim and extract text from.
'#
'# RETURNS:
'#
'#      Text trimmed of leading and trailing whitespacespaces but retaining
'#      whatever exists between a leading and/or trailing double-quote mark.
'#
'#      Text not containing leading or trailing quote marks (after trimming)
'#      will be returned unmodified (untrimmed).
'#
'# ERRORS:
'#
'#      none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#
'############################################################################
Public Function GetQuotedString(sInput As String) As String

    Dim value As String
    Dim leftquoted As Integer, rightquoted As Integer
    
    GetQuotedString = sInput
        
    If sInput = "" Then Exit Function
    
    value = TWhitespace(sInput)
    If value = "" Then Exit Function
    
    If Left$(value, 1) = gQT Then leftquoted = 1
    If Right$(value, 1) = gQT Then rightquoted = 1
    
    If ((leftquoted = 0) And (rightquoted = 0)) Then Exit Function
    
    If rightquoted = 1 Then value = Left$(value, Len(value) - 1)

    rightquoted = Len(value)
    If ((leftquoted = 1) And (rightquoted > 0)) Then value = Right$(value, rightquoted - 1)

    GetQuotedString = value

End Function


'############################################################################
'#
'#  Function ConvertHTMLEntities(source As String) As String
'#
'# DESCRIPTION:
'#
'#      If source string is of type HTML, this function will search for and
'#      eliminate instances of HTML markup.  This is done on a
'#      copy of the source string.  The routine returns this new string on
'#      exit.  All strings are treated as case-insensitive.
'#
'#      Currently converts:
'#
'#          &nbsp;          &lt;
'#          &#160;          &#060;
'#          &quot;          &gt;
'#          &#034;          &#062;
'#          &amp;           &#032;
'#          &#038;
'#
'# PARAMETERS:
'#
'#      source      the String to seek and convert HTML tags on.
'#                  It is not modified.
'#
'# RETURNS:
'#
'#      New String after ConvertHTMLEntities is performed.  The input source string
'#      is not modified.
'#
'#      If the source string is invalid ("") then the return value is "".
'#
'# ERRORS:
'#
'#       none
'#
'# Orig Author: Dawn Coen/Carl Nagle
'# Orig Date: FEB 11, 2002
'# History:
'#
'#      FEB 12, 2002    Original Release
'#
'############################################################################
Function ConvertHTMLEntities(source As String) As String

    Dim find As String
    Dim replace As String
    Dim text As String
    Dim index As Integer
    
    If source = "" Then Exit Function
    text = source
    
    On Error Resume Next

    For index = 1 To UBound(gHTMLEntityValues, 1)
    
        find = gHTMLEntityValues(index, 0)
        replace = gHTMLEntityValues(index, 1)
        text = FindAndReplace(text, find, replace, 1)
    
    Next
    
    ConvertHTMLEntities = text
    
End Function


'############################################################################
'#
'#  Function GetUnsupportedExpressionOperators() As String
'#
'# DESCRIPTION:
'#
'#      A string of characters representing expression operators often found
'#      in other languages or formats but not supported at this time.
'#
'#      Currently: "[]<>!|"
'#
'# PARAMETERS:
'#
'#      none
'#
'# RETURNS:
'#
'#      String of operators known to be unsupported for expression processing
'#
'# ERRORS:
'#
'#       none
'#
'# Orig Author: Carl Nagle
'# Orig Date: Nov 15, 2006
'# History:
'#
'#      Nov 15, 2006    Original Release
'#
'############################################################################
Public Function GetUnsupportedExpressionOperators() As String
    GetUnsupportedExpressionOperators = gUnsupportedOperators
End Function


'############################################################################
'#
'#  Function GetSupportedExpressionOperators() As String
'#
'# DESCRIPTION:
'#
'#      A string of characters representing supported expression operators.
'#
'#      Currently: "=&+-*/%()"
'#
'# PARAMETERS:
'#
'#      none
'#
'# RETURNS:
'#
'#      String of operators known to be supported for expression processing
'#
'# ERRORS:
'#
'#       none
'#
'# Orig Author: Carl Nagle
'# Orig Date: Nov 15, 2006
'# History:
'#
'#      Nov 15, 2006    Original Release
'#
'############################################################################
Public Function GetSupportedExpressionOperators() As String
    GetSupportedExpressionOperators = gSupportedOperators
End Function


'############################################################################
'#
'#  Function GetOperatorPrecedence() As String
'#
'# DESCRIPTION:
'#
'#      A string of characters representing operator precedence in expressions.
'#
'#      Currently: "*/%+-&"
'#
'# PARAMETERS:
'#
'#      none
'#
'# RETURNS:
'#
'#      A string of characters representing operator precedence in expressions.
'#
'# ERRORS:
'#
'#       none
'#
'# Orig Author: Carl Nagle
'# Orig Date: Nov 15, 2006
'# History:
'#
'#      Nov 15, 2006    Original Release
'#
'############################################################################
Public Function GetOperatorPrecedence() As String
    GetOperatorPrecedence = gOperatorPrecedence
End Function


'############################################################################
'# <a name="convertvariableexpression"></a>
'#
'#  Function ConvertVariableExpression (expression As String,
'#                                      store As Object,
'#                                      value as String) As Integer
'#
'# DESCRIPTION:
'#
'#      Given an expression attempt to locate variable assignment references
'#      and other operators and process the expressions into a result string.
'#
'#      Valid operators (default values):
'#
'#              ^                    '(Caret) Variable Prefix
'#              =                    'Assignment operator
'#              "                    'A single Double-Quote mark
'#              &                    'String concatenate operator
'#              +                    'Addition operator
'#              -                    'Subtraction operator
'#              *                    'Multiplication operator
'#              /                    'Division operator
'#              %                    'Modulus/Remainder operator
'#              (                    'Open Group operator
'#              )                    'Close Group operator
'#
'#      Double-quote marks are used to enclose literal text strings or values.
'#
'#      Examples:
'#
'#          INPUT:                      OUTPUT:
'#          ======================      =======================
'#          "some literal text"         some literal text
'#          "82"                        82
'#
'#      To include a double-quote inside a literal text string use two
'#      double-quote marks with no intervening characters.  These must appear
'#      inside the outermost quote marks that delimit a literal text string.
'#
'#      Examples:
'#
'#          INPUT:                      OUTPUT:
'#          """some literal text"""     "some literal text"
'#          "some ""literal"" text"     some "literal" text
'#
'#      Oddly quoted strings that do not have operators can produce unexpected or
'#      undefined results.  The results of some of these odd circumstances should
'#      not be considered reliable and may change as the handling becomes better
'#      defined in later releases.  To ensure proper results, always properly
'#      quote input values and use appropriate expression operators.
'#
'#      Recent Odd Examples (Don't do stuff like this):
'#
'#          INPUT:  "nothing" "complicated"
'#          OUTPUT: nothing" "complicated
'#
'#          INPUT:  "nothing" complicated"
'#          OUTPUT: nothing" complicated
'#
'#
'#      Precedence:
'#
'#      Groupings are processed first.  Within groupings, and after all groups
'#      have been processed, operators are processed in the following order:
'#
'#          */%+-&
'#
'#
'#      The expression will be trimmed of leading and trailing whitespace on
'#      entry.  Use double-quote marks (literal text) to retain significant
'#      leading and trailing whitespace.
'#
'#      Variable assignment references are identified by a leading caret (^)
'#      immediately followed by the name of the variable, an equal sign (=),
'#      and the value to assign to that variable. Whitespace can exist on either
'#      side of the equal sign.
'#
'#      Variable names must conform to the ValidateVariableID routine.
'#
'#      Variable lookups(GET) and assignments(SET) will be done via the provided
'#      DDVariables or GlobalVariables store.
'#
'#      If the proposed value is NOT another variable reference then the proposed
'#      value will be trimmed of leading and trailing whitespace (spaces and tabs).
'#      And, yes, a variable can be assigned the value of another variable.
'#
'#      Additionally, if there is text or operators that are not to be considered as
'#      variables they should be enclosed in double-quote marks.
'#
'#
'#      IMPORTANT NOTE:
'#      ===============
'#      Expressions containing groupings but no other legitimate operators may
'#      produce unexpected results that may not be the same on later versions of
'#      these routines.
'#
'#      Recent Odd Examples (Don't do stuff like this):
'#
'#          INPUT:  Now ( 2 ( groups ) with no quotes).
'#          OUTPUT: Now "2 "groups" with no quotes".
'#
'#          INPUT:  "Now " ( "2" ( " groups" ) " with quotes" ) "."
'#          OUTPUT: Now " "2" " groups" " with quotes" ".
'#
'#      To ensure proper results, always properly quote input values and use
'#      appropriate expression operators.
'#
'#
'# PARAMETERS:
'#
'#      expression      string expression containing references to variables.
'#                      This String will not be modified by this routine.  A copy
'#                      will be created and modified for return.
'#
'#      store           A DDVariables or GlobalVariables Object for GETTING and
'#                      SETTING variable values.  The routine will exit with error
'#                      if the provided Object is not of one of these types.
'#
'#      value           String to receive the result of the expression processing.
'#                      This will NOT be modified if the routine exits with error.
'#
'# RETURNS:
'#
'#       0   successful processing of the expression
'#      -1   general failure of the routine
'#      -2   store Object parameter invalid type
'#
'# ERRORS:
'#
'#       none
'#
'# Orig Author: Carl Nagle
'# Orig   Date: AUG 15, 2001
'# History:
'#
'#      AUG 15, 2001    Original Release
'#      NOV 04, 2001    (Carl Nagle) Fixed Variable Assignment in Groupings bug.
'#      APR 11, 2002    (Carl Nagle) LTrim numeric values.
'#      MAR 28, 2003    (Carl Nagle) Slight perf. improvement and more doc.
'#
'############################################################################
Public Function ConvertVariableExpression(expression As String, store As Object, value As String) As Integer
    
    Static entry As Integer          'increments on reentry. decrements on exit
    Static assignentry  As Integer   '0= no assignment lvalue
    
    Dim field As String
    Dim lenField As Long
    Dim BQIndex As Long
    Dim EQindex As Long
    Dim EGRPindex As Long
    Dim BGRPindex As Long
    Dim tempIndex As Long
    
    Dim lvalue As String
    Dim rvalue As String
    Dim operator1 As String
    Dim operator2 As String
    Dim status As Integer
    
    'trim leading and trailing whitespace
    field = TWhitespace(expression)
    
    'preset return value
    ConvertVariableExpression = -1
        
    'store must be either a DDVariables or GlobalVariables object
    If (Not TypeOf store Is DDVariables) And _
       (Not TypeOf store Is GlobalVariables) Then
       ConvertVariableExpression = -2
       Exit Function
    End If
    
    'increment our re-entry counter
    entry = entry + 1
    
    lenField = Len(field)               'added Carl Nagle
    If lenField < 1 Then GoTo CVEExit   'added Carl Nagle
    
    'first exit if we contain unquoted, unsupported operators
    'If GetUnquotedDelimiterIndex(1, field, gUnsupportedOperators) > 0 Then GoTo CVEExit
    If optGetUnquotedDelimiterIndex(1, field, lenField, _
            gUnsupportedOperators, lengUnsupportedOperators) > 0 Then GoTo CVEExit
    
    'now handle groupings til they are all gone
    Do
        'find unquoted end group
        EGRPindex = optGetUnquotedDelimiterIndex(1, field, lenField, gEGRP, 1)
        'find unquoted begin group backwards from end group (or text end)
        If EGRPindex > 0 Then
            tempIndex = EGRPindex
        Else
            tempIndex = lenField
        End If
        BGRPindex = optGetUnquotedDelimiterIndexRev(tempIndex, field, lenField, gBGRP, 1)
        
        'better have both or none
        If (BGRPindex > 0) Xor (EGRPindex > 0) Then GoTo CVEExit
        'recurse this routine with that substring if necessary
        If BGRPindex > 0 Then
            status = ConvertVariableExpression(Mid$(field, BGRPindex + 1, EGRPindex - BGRPindex - 1), store, rvalue)
            'exit if some error occurred - don't know what that could be at this point
            If status <> 0 Then GoTo CVEExit
            
            'substitute the resulting text of the group into the field string
            field = Left$(field, BGRPindex - 1) & rvalue & Right$(field, lenField - EGRPindex)
            'recalc for each iteration
            lenField = Len(field)
        End If
    'keep going to we find no more groupings
    Loop While (BGRPindex > 0) And (EGRPindex > 0)
        
    'check for an lvalue assignment operation
    EQindex = InStr(field, gEQ)
    If EQindex > 0 Then
        Dim storeval As String
        'is first EQ quoted? If so, process this as a regular rvalue
        If IsQuoted(field, EQindex) Then GoTo CVEEQBypass
        'if  0 < EQ < 2 then it cannot be a valid assignment.
        If EQindex < 2 Then GoTo CVEExit
    
        'exit if the LValue is not valid variableID format( ^Avariable )
        'returns with the variableID "Avariable" (no caret) if successful
        lvalue = ValidateVariableID(Left$(field, EQindex - 1))
        If Len(lvalue) = 0 Then GoTo CVEExit
        
        assignentry = assignentry + 1
        
        'get rvalue by recursing this function
        status = ConvertVariableExpression(Mid$(field, EQindex + 1), store, rvalue)
        
        'we must remove any extra quotes for DDVariableStore purposes
        'this fixes the bug reported by Y.Wang on multi-entry (grouped) variable assignments
        If (status = 0) And (entry > assignentry) Then
            storeval = GetTrimmedQuotedString(FindAndReplace(rvalue, gQT & gQT, gQT))
        Else
            storeval = rvalue
        End If
        
        assignentry = assignentry - 1
        
        If status = 0 Then
            value = rvalue
            status = store.SetVariableValue(lvalue, storeval)
            ConvertVariableExpression = 0
        End If
        GoTo CVEExit
    End If
    
CVEEQBypass:
        
    'now just start piecing things together.
    'all groupings are gone. All assignments are gone (sort of).
    'just variables, text, and supported operators.
                    
    BQIndex = 1     'the start of our current lvalue
    lenField = Len(field)
    Dim mCol As New DDVariableStore.DDVariables
    Dim mColSize As Long
    Dim lenOp2 As Integer
    
    'extract all variable values and literal text and operators into Collection
    'storage is value, operator, value, operator, value, etc..
    Do
        lvalue = ""
        rvalue = ""
        BGRPindex = -1
        EGRPindex = -1
        operator1 = ""
        operator2 = ""
        lenOp2 = 0
        
        'find the next unquoted operator for our lvalue
        BGRPindex = optGetUnquotedDelimiterIndex(BQIndex, field, lenField, _
                                                    gSupportedOperators, lengSupportedOperators)
        
        'if valid, get the operator and find the next unquoted operator for our rvalue
        If BGRPindex > 0 Then
            operator1 = Mid$(field, BGRPindex, 1)
            If BGRPindex < lenField Then
                EGRPindex = optGetUnquotedDelimiterIndex(BGRPindex + 1, field, lenField, _
                                                     gSupportedOperators, lengSupportedOperators)
                If EGRPindex > 0 Then
                    operator2 = Mid$(field, EGRPindex, 1)
                    lenOp2 = 1
                Else
                    EGRPindex = lenField
                End If
            Else
                EGRPindex = lenField
            End If
        Else
            BGRPindex = lenField
        End If
        
        'BQIndex = start of lvalue
        'BGRPindex = end of lvalue, index of operator (if any)
        'operator = operator string character (if any)
        'EGRPindex = end of rvalue after operator (if any)
        
        'extract any variable value and quote the lvalue
        If EGRPindex > BGRPindex Then
            lvalue = Mid$(field, BQIndex, BGRPindex - BQIndex)
            'do not need to mess with value if it contains quotes
            If InStr(1, lvalue, gQT, vbTextCompare) < 1 Then
                'lvalue = FindAndReplace(lvalue, gQT & gQT, gQT)
                lvalue = QTTextOfVariable(lvalue, store)
            Else
                lvalue = FindAndReplace(lvalue, gQT & gQT, gQT)
                'lvalue = FindAndReplace(lvalue, gQT & gQT, Chr$(1))
                lvalue = TWhitespace(lvalue)
            End If
            BQIndex = EGRPindex + 1
            
            'extract any variable value and quote the rvalue
            rvalue = Mid$(field, BGRPindex + 1, EGRPindex - BGRPindex - lenOp2)
            'do not need to mess with value if it contains quotes
            If InStr(1, rvalue, gQT, vbTextCompare) < 1 Then
                'rvalue = FindAndReplace(rvalue, gQT & gQT, gQT)
                rvalue = QTTextOfVariable(rvalue, store)
            Else
                rvalue = FindAndReplace(rvalue, gQT & gQT, gQT)
                'rvalue = FindAndReplace(rvalue, gQT & gQT, Chr$(1))
                rvalue = TWhitespace(rvalue)
            End If
        Else
            'extract any variable value and quote the lvalue
            lvalue = Mid$(field, BQIndex, BGRPindex - BQIndex + 1)
            'do not need to mess with value if it contains quotes
            If InStr(1, lvalue, gQT, vbTextCompare) < 1 Then
                'lvalue = FindAndReplace(lvalue, gQT & gQT, gQT)
                lvalue = QTTextOfVariable(lvalue, store)
            Else
                lvalue = FindAndReplace(lvalue, gQT & gQT, gQT)
                'lvalue = FindAndReplace(lvalue, gQT & gQT, Chr$(1))
                lvalue = TWhitespace(lvalue)
            End If
            BQIndex = BGRPindex + 1
        End If
        
        'place all values and operators into storage
        mColSize = mCol.count
        mCol.SetVariableValue Str$(mColSize + 1), lvalue
        If EGRPindex > BGRPindex Then
            mCol.SetVariableValue Str$(mColSize + 2), operator1
            mCol.SetVariableValue Str$(mColSize + 3), rvalue
            If lenOp2 Then mCol.SetVariableValue Str$(mColSize + 4), operator2
        End If
                
    Loop Until BQIndex > lenField
        
    'next handle all right to left operators
    'we have none yet
    
    
    'now handle left to right operators
    EQindex = mCol.count                    'the total number of values and operators
    
    'remove outside whitespace and quotes
    For BQIndex = 1 To EQindex
        mCol.Item(BQIndex) = GetTrimmedQuotedString(mCol.Item(BQIndex))
    Next
    
    Dim dlvalue As Double
    Dim drvalue As Double
    Dim operatorIndex As Integer
    Dim precedenceIndex As Integer
    Dim precedenceTest As String
    
    Do While EQindex > 1
        
        rvalue = ""
        operator1 = ""
        precedenceIndex = 1
        Do Until (precedenceIndex > lengOperatorPrecedence) Or (operator1 <> "")
            precedenceTest = Mid$(gOperatorPrecedence, precedenceIndex, 1)
            For operatorIndex = 2 To EQindex Step 2
                If precedenceTest = mCol.Item(operatorIndex) Then
                    operator1 = precedenceTest
                    Exit For
                End If
            Next
            If operator1 = "" Then precedenceIndex = precedenceIndex + 1
        Loop
        
        'we have to have matched an operator. So, we have an operatorIndex
        lvalue = mCol.Item(operatorIndex - 1)
        If operatorIndex < EQindex Then rvalue = mCol.Item(operatorIndex + 1)
        
        Select Case operator1
        
            '&
            Case gJN
                mCol.Item(operatorIndex - 1) = lvalue & rvalue
            
            '+
            Case gADD
                dlvalue = 0
                drvalue = 0
                On Error Resume Next
                dlvalue = CDbl(lvalue)
                drvalue = CDbl(rvalue)
                On Error GoTo 0
                mCol.Item(operatorIndex - 1) = LTrim$(Str$(dlvalue + drvalue))
                
            '-
            Case gSUB
                dlvalue = 0
                drvalue = 0
                On Error Resume Next
                dlvalue = CDbl(lvalue)
                drvalue = CDbl(rvalue)
                On Error GoTo 0
                mCol.Item(operatorIndex - 1) = LTrim$(Str$(dlvalue - drvalue))
                
            '*
            Case gMUL
                dlvalue = 0
                drvalue = 0
                On Error Resume Next
                dlvalue = CDbl(lvalue)
                drvalue = CDbl(rvalue)
                On Error GoTo 0
                mCol.Item(operatorIndex - 1) = LTrim$(Str$(dlvalue * drvalue))
                
            ' / (division)
            Case gDIV
                dlvalue = 0
                drvalue = 0
                On Error Resume Next
                dlvalue = CDbl(lvalue)
                drvalue = CDbl(rvalue)
                If drvalue = 0 Then
                    'cannot divide by 0
                    'should we change stored value?
                    mCol.Item(operatorIndex - 1) = field
                Else
                    mCol.Item(operatorIndex - 1) = LTrim$(Str$(dlvalue / drvalue))
                End If
                On Error GoTo 0
                
            ' % (remainder)
            Case gREM
                dlvalue = 0
                drvalue = 0
                On Error Resume Next
                dlvalue = CDbl(lvalue)
                drvalue = CDbl(rvalue)
                If drvalue = 0 Then
                    'cannot divide by 0
                    'should we change stored value?
                    mCol.Item(operatorIndex - 1) = field
                Else
                    mCol.Item(operatorIndex - 1) = LTrim$(Str$(dlvalue Mod drvalue))
                End If
                On Error GoTo 0
            
            Case Else
            
                'what do we do on error here?  Should we ever get here?
                'only if there was an error in our programming logic above.
                MsgBox "Error in expression storage!", vbCritical Or vbOKOnly, "DDVariableStore Storage Error"
                
        End Select
        
        If operatorIndex < EQindex Then mCol.Remove (operatorIndex + 1)
        mCol.Remove (operatorIndex)
        'we keep and work with the item to the left of the operator
        EQindex = mCol.count
    Loop
    
    If EQindex = 1 Then
        value = mCol.Item(1)
        
        'if this value will be used again in this routine we need
        'to add back in any double-quotes
        
        If entry > (1 + assignentry) Then
        'If entry > 1 + assignentry Then
            value = gQT & FindAndReplace(value, gQT, gQT & gQT) & gQT
        End If
        
        ConvertVariableExpression = 0
    End If
    
CVEExit:
    
    entry = entry - 1
    If entry < 0 Then entry = 0
    Set mCol = Nothing

End Function



Private Sub Class_Initialize()

    InitializeStringUtilities       'SharedModule
    
End Sub
