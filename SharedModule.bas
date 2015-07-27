Attribute VB_Name = "SharedModule"
Option Explicit

'# GLOBAL STORAGE for cross-process class instances.

'for GlobalVariables
Public gDDVariables As DDVariables
Public glngDDVRefCount As Long

Public bUseVARS As Boolean
Public bUseMAPS As Boolean

'for GlobalMappings
Public gDDMappings As DDVariables
Public glngDDMRefCount As Long

Public Const gDD_DEFAULT_MAP_SECTION = "ApplicationConstants"
Public gDDDefaultMapSection As String
Public gDDCurrentAppMap As String
Public DEBUG_ENABLED As Boolean
Public debugfile As Object
Public Const DEBUG_FILE = "C:\Debug_ddvariable.txt"
Public debugbinaryfile As Object
Public Const DEBUG_BINARY_FILE = "C:\Debug_ddvariableBinary.txt"

'for StringUtilities
Public gValidDVNameCharacters As String   'filled by StringUtilities.Class_Initialize
Public gHTMLEntityValues(1 To 11, 0 To 1) As String   'filled by StringUtilities.Class_Initialize (1-based 2D array)

Public gVID As String                     'Variable Prefix
Public gEQ  As String                     'Assignment operator
Public gQT  As String
Public gJN  As String                     'Concatenate operator
Public gADD As String                     'Addition operator
Public gSUB  As String                    'Subtraction operator
Public gMUL  As String                    'Multiplication operator
Public gDIV  As String                    'Division operator
Public gREM  As String                    'Modulus/Remainder operator
Public gBGRP  As String                   'Open Group operator
Public gEGRP  As String                   'Close Group operator

Public gAllOperators  As String           'some are reserved
Public gSupportedOperators  As String     'supported operators
Public gUnsupportedOperators  As String   'not yet supported
Public lengAllOperators  As Long          'the length of that string
Public lengSupportedOperators  As Long    'the length of that string
Public lengUnsupportedOperators  As Long  'the length of that string

Public gOperatorPrecedence  As String     'the order they are processed
Public lengOperatorPrecedence As Long     'the length of that string

Public gWhitespace As String              'space and tab in Class_Initialize

'DEFAULT VALUES FOR THESE THINGS
Private Const cgVID = "^"                    'Variable Prefix
Private Const cgEQ = "="                     'Assignment operator
Private Const cgQT = """"                    'A single Double-Quote mark
Private Const cgJN = "&"                     'Concatenate operator
Private Const cgADD = "+"                    'Addition operator
Private Const cgSUB = "-"                    'Subtraction operator
Private Const cgMUL = "*"                    'Multiplication operator
Private Const cgDIV = "/"                    'Division operator
Private Const cgREM = "%"                    'Modulus/Remainder operator
Private Const cgBGRP = "("                   'Open Group operator
Private Const cgEGRP = ")"                   'Close Group operator

Private Const cgAllOperators = "=&+-*/%()[]<>!|"     'some are reserved
Private Const cgSupportedOperators = "=&+-*/%()"     'supported operators
Private Const cgUnsupportedOperators = "[]<>!|"      'not yet supported

Private Const cgOperatorPrecedence = "*/%+-&"        'the order they are processed
Private Const clengOperatorPrecedence = 6            'the length of that string

Public Enum KnownCodePage
CP_UNKNOWN = -1
CP_ACP = 0
CP_OEMCP = 1
CP_MACCP = 2
CP_THREAD_ACP = 3
CP_SYMBOL = 42
' ARABIC
CP_AWIN = 101 ' Bidi Windows codepage
CP_709 = 102 ' MS-DOS Arabic Support CP 709
CP_720 = 103 ' MS-DOS Arabic Support CP 720
CP_A708 = 104 ' ASMO 708
CP_A449 = 105 ' ASMO 449+
CP_TARB = 106 ' MS Transparent Arabic
CP_NAE = 107 ' Nafitha Enhanced Arabic Char Set
CP_V4 = 108 ' Nafitha v 4.0
CP_MA2 = 109 ' Mussaed Al Arabi (MA/2) CP 786
CP_I864 = 110 ' IBM Arabic Supplement CP 864
CP_A437 = 111 ' Ansi 437 codepage
CP_AMAC = 112 ' Macintosh Code Page
' HEBREW
CP_HWIN = 201 ' Bidi Windows codepage
CP_862I = 202 ' IBM Hebrew Supplement CP 862
CP_7BIT = 203 ' IBM Hebrew Supplement CP 862 Folded
CP_ISO = 204 ' ISO Hebrew 8859-8 Character Set
CP_H437 = 205 ' Ansi 437 codepage
CP_HMAC = 206 ' Macintosh Code Page
' CODE PAGES
CP_OEM_437 = 437
CP_ARABICDOS = 708
CP_DOS720 = 720
CP_DOS737 = 737
CP_DOS775 = 775
CP_IBM850 = 850
CP_IBM852 = 852
CP_DOS861 = 861
CP_DOS862 = 862
CP_IBM866 = 866
CP_DOS869 = 869
CP_THAI = 874
CP_EBCDIC = 875
CP_JAPAN = 932
CP_CHINA = 936
CP_KOREA = 949
CP_TAIWAN = 950
' UNICODE
CP_UNICODELITTLE = 1200
CP_UNICODEBIG = 1201
' CODE PAGES
CP_EASTEUROPE = 1250
CP_RUSSIAN = 1251
CP_WESTEUROPE = 1252
CP_GREEK = 1253
CP_TURKISH = 1254
CP_HEBREW = 1255
CP_ARABIC = 1256
CP_BALTIC = 1257
CP_VIETNAMESE = 1258
' KOREAN
CP_JOHAB = 1361
' MAC
CP_MAC_ROMAN = 10000
CP_MAC_JAPAN = 10001
CP_MAC_ARABIC = 10004
CP_MAC_GREEK = 10006
CP_MAC_CYRILLIC = 10007
CP_MAC_LATIN2 = 10029
CP_MAC_TURKISH = 10081
' CODE PAGES
CP_CHINESECNS = 20000
CP_CHINESEETEN = 20002
CP_IA5WEST = 20105
CP_IA5GERMAN = 20106
CP_IA5SWEDISH = 20107
CP_IA5NORWEGIAN = 20108
CP_ASCII = 20127
CP_RUSSIANKOI8R = 20866
CP_RUSSIANKOI8U = 21866
CP_ISOLATIN1 = 28591
CP_ISOEASTEUROPE = 28592
CP_ISOTURKISH = 28593
CP_ISOBALTIC = 28594
CP_ISORUSSIAN = 28595
CP_ISOARABIC = 28596
CP_ISOGREEK = 28597
CP_ISOHEBREW = 28598
CP_ISOTURKISH2 = 28599
CP_ISOLATIN9 = 28605
CP_HEBREWLOG = 38598
CP_USER = 50000
CP_AUTOALL = 50001
CP_JAPANNHK = 50220
CP_JAPANESC = 50221
CP_JAPANISO = 50222
CP_KOREAISO = 50225
CP_TAIWANISO = 50227
CP_CHINAISO = 50229
CP_AUTOJAPAN = 50932
CP_AUTOCHINA = 50936
CP_AUTOKOREA = 50949
CP_AUTOTAIWAN = 50950
CP_AUTORUSSIAN = 51251
CP_AUTOGREEK = 51253
CP_AUTOARABIC = 51256
CP_JAPANEUC = 51932
CP_CHINAEUC = 51936
CP_KOREAEUC = 51949
CP_TAIWANEUC = 51950
CP_CHINAHZ = 52936
CP_GB18030 = 54936
' UNICODE
CP_UTF7 = 65000
CP_UTF8 = 65001
End Enum

' Flags
Public Const MB_PRECOMPOSED = &H1
Public Const MB_COMPOSITE = &H2
Public Const MB_USEGLYPHCHARS = &H4
Public Const MB_ERR_INVALID_CHARS = &H8

Public Const WC_DEFAULTCHECK = &H100 ' check for default char
Public Const WC_COMPOSITECHECK = &H200 ' convert composite to precomposed
Public Const WC_DISCARDNS = &H10 ' discard non-spacing chars
Public Const WC_SEPCHARS = &H20 ' generate separate chars
Public Const WC_DEFAULTCHAR = &H40 ' replace with default char

Public Declare Function GetACP Lib "kernel32" () As Long

Public Const UTF8 = 65001

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, _
                                                             ByVal dwFlags As Long, _
                                                             ByVal lpMultiByteStr As Long, _
                                                             ByVal cchMultiByte As Long, _
                                                             ByVal lpWideCharStr As Long, _
                                                             ByVal cchWideChar As Long) As Long

Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, _
                                                             ByVal dwFlags As Long, _
                                                             ByVal lpWideCharStr As Long, _
                                                             ByVal cchWideChar As Long, _
                                                             ByVal lpMultiByteStr As Long, _
                                                             ByVal cchMultiByte As Long, _
                                                             ByVal lpDefaultChar As Long, _
                                                             lpUsedDefaultChar As Long) As Long


'called by each initialization of a StringUtilities instance.
'However, only the first one will actually initialize the data.
Public Sub InitializeStringUtilities()
    
    Dim index As Integer
    
    'if the global storage has not been initialized, do it
    If Len(gValidDVNameCharacters) = 0 Then
        
        'define what is whitespace
        gWhitespace = " " & Chr$(9)
        
        'set default values for operators
        gVID = cgVID                                        'Variable Prefix
        gEQ = cgEQ                                          'Assignment operator
        gQT = cgQT                                          'single double-quote mark
        gJN = cgJN                                          'Concatenate operator
        gADD = cgADD                                        'Addition operator
        gSUB = cgSUB                                        'Subtraction operator
        gMUL = cgMUL                                        'Multiplication operator
        gDIV = cgDIV                                        'Division operator
        gREM = cgREM
        gBGRP = cgBGRP                                      'Open Group operator
        gEGRP = cgEGRP                                      'Close Group operator
        
        gAllOperators = cgAllOperators                      'some are reserved
        lengAllOperators = Len(gAllOperators)
        gSupportedOperators = cgSupportedOperators          'supported operators
        lengSupportedOperators = Len(gSupportedOperators)
        gUnsupportedOperators = cgUnsupportedOperators      'not yet supported
        lengUnsupportedOperators = Len(gUnsupportedOperators)

        gOperatorPrecedence = cgOperatorPrecedence          'the order they are processed
        lengOperatorPrecedence = clengOperatorPrecedence    'the length of that string
        
        'a-z
        For index = 97 To 122
            gValidDVNameCharacters = gValidDVNameCharacters & Chr$(index)
        Next
            
        '0-9
        For index = 48 To 57
            gValidDVNameCharacters = gValidDVNameCharacters & Chr$(index)
        Next
            
        'A-Z
        For index = 65 To 90
            gValidDVNameCharacters = gValidDVNameCharacters & Chr$(index)
        Next
            
        ' . (period)
        gValidDVNameCharacters = gValidDVNameCharacters & Chr$(46)
            
        ' _ (underscore)
        gValidDVNameCharacters = gValidDVNameCharacters & Chr$(95)
            
        'weird foreign stuff
        For index = 192 To 246
            gValidDVNameCharacters = gValidDVNameCharacters & Chr$(index)
        Next
            
        'weird foreign stuff
        For index = 248 To 255
            gValidDVNameCharacters = gValidDVNameCharacters & Chr$(index)
        Next
            
        ' f  (Florin)
        gValidDVNameCharacters = gValidDVNameCharacters & Chr$(131)
        ' OE  (OE ligature)
        gValidDVNameCharacters = gValidDVNameCharacters & Chr$(140)
        ' oe  (oe ligature)
        gValidDVNameCharacters = gValidDVNameCharacters & Chr$(156)
        ' Y  (Y umlaut)
        gValidDVNameCharacters = gValidDVNameCharacters & Chr$(159)
            
    End If

    'don't forget to set dimension values up in Public area when adding more entities
    If gHTMLEntityValues(1, 1) <> Chr(160) Then
    
        gHTMLEntityValues(1, 0) = "&nbsp;"      '&nbsp;
        gHTMLEntityValues(1, 1) = Chr(160)
        gHTMLEntityValues(2, 0) = "&#160;"
        gHTMLEntityValues(2, 1) = Chr(160)
        gHTMLEntityValues(3, 0) = "&quot;"      '&quot;
        gHTMLEntityValues(3, 1) = Chr(34)
        gHTMLEntityValues(4, 0) = "&#034;"
        gHTMLEntityValues(4, 1) = Chr(34)
        gHTMLEntityValues(5, 0) = "&amp;"       '&amp;
        gHTMLEntityValues(5, 1) = Chr(38)
        gHTMLEntityValues(6, 0) = "&#038;"
        gHTMLEntityValues(6, 1) = Chr(38)
        gHTMLEntityValues(7, 0) = "&lt;"        '&lt;
        gHTMLEntityValues(7, 1) = Chr(60)
        gHTMLEntityValues(8, 0) = "&#060;"
        gHTMLEntityValues(8, 1) = Chr(60)
        gHTMLEntityValues(9, 0) = "&gt;"        '&gt;
        gHTMLEntityValues(9, 1) = Chr(62)
        gHTMLEntityValues(10, 0) = "&#062;"
        gHTMLEntityValues(10, 1) = Chr(62)
        gHTMLEntityValues(11, 0) = "&#032;"     'space
        gHTMLEntityValues(11, 1) = Chr(32)
                
    End If


End Sub

'*** THIS FUNCTION DOES NOT YET WORK PROPERLY -- NEEDS DEBUGGING
Private Function ANSItoUTF8(sText As String, Optional ByVal cPage As KnownCodePage = CP_UNKNOWN) As String
            
    Static tmpArr() As Byte, textStr As String
    Static textArr() As Byte
    Dim tmpLen As Long, textLen As Long, A As Long
    Dim text() As Byte, sUTF8 As String
    
    ' set code page to a valid one
    If cPage = CP_UNKNOWN Then cPage = GetACP
    
    'convert Text() to Unicode tmpArr
    If cPage = CP_ACP Or cPage = CP_WESTEUROPE Then
        textLen = Len(sText)
        tmpLen = textLen + textLen + 1
        If (Not tmpArr) = True Then ReDim Preserve tmpArr(tmpLen)
        If UBound(tmpArr) <> tmpLen Then ReDim Preserve tmpArr(tmpLen)
        For A = 0 To textLen - 1
            tmpArr(A + A) = AscW(Mid$(sText, A + 1, 1))
        Next A
    
    'convert Text() to Unicode tmpArr
    Else
        textLen = LenB(sText)
        tmpLen = textLen + textLen
        ReDim Preserve tmpArr(tmpLen + 1)
        
        'get the new string to tmpArr
        tmpLen = MultiByteToWideChar(CLng(cPage), ByVal 0&, ByVal StrPtr(sText), -1, _
                                     ByVal VarPtr(tmpArr(0)), tmpLen)
        If tmpLen = 0 Then Exit Function
        tmpLen = tmpLen + tmpLen - 5
        'If tmpArr(tmpLen - 1) = 0 And tmpArr(tmpLen) = 0 Then tmpLen = tmpLen - 2
        If UBound(tmpArr) <> tmpLen Then ReDim Preserve tmpArr(tmpLen)
    End If
    
    'convert Unicode tmpArr to UTF-8 textArr
    If (Not textArr) = True Then ReDim Preserve textArr(tmpLen)
    If UBound(textArr) <> tmpLen Then ReDim Preserve textArr(tmpLen)
    textLen = tmpLen + tmpLen + tmpLen + 1
    tmpLen = WideCharToMultiByte(CP_UTF8, ByVal 0&, ByVal VarPtr(tmpArr(0)), tmpLen, ByVal VarPtr(textArr(0)), _
                                    textLen, ByVal 0&, ByVal 0&)
    ' a hopeless try to correct a weird error?
    ReDim Preserve textArr(tmpLen - 1)
    sUTF8 = CStr(textArr)
    ANSItoUTF8 = sUTF8
End Function

'*** THIS FUNCTION DOES NOT YET WORK PROPERLY -- NEEDS DEBUGGING
Public Function UTF8toANSI(sText As String, Optional ByVal cPage As KnownCodePage = CP_UNKNOWN, _
                            Optional lFlags As Long) As String
    Static tmpArr() As Byte, textArr() As Byte, textStr As String
    Dim tmpLen As Long, textLen As Long, A As Long
    Dim text() As Byte, sUTF8 As String
    
    textLen = LenB(sText)
    tmpLen = textLen + textLen
    ReDim Preserve tmpArr(tmpLen + 1)
    tmpLen = MultiByteToWideChar(CP_UTF8, ByVal 0&, ByVal StrPtr(sText), -1, _
                                 ByVal VarPtr(tmpArr(0)), tmpLen)
                                 
    ' set code page to a valid one
    If cPage = CP_UNKNOWN Then cPage = GetACP
    If cPage = CP_ACP Or cPage = CP_WESTEUROPE Then
        textLen = UBound(tmpArr)
        tmpLen = (textLen + 1) \ 2 - 1
        If (Not textArr) = True Then ReDim Preserve textArr(tmpLen)
        If UBound(textArr) <> tmpLen Then ReDim Preserve textArr(tmpLen)
        For A = 0 To tmpLen
            textArr(A) = tmpArr(A + A)
        Next A
    Else
        textLen = (UBound(tmpArr) + 1)
        ' at maximum ANSI can be four bytes per character in new Chinese encoding GB18030–2000
        tmpLen = textLen + textLen
        ReDim Preserve textArr(tmpLen - 1)
        ' get the new string to tmpArr
        tmpLen = WideCharToMultiByte(CLng(cPage), lFlags, ByVal VarPtr(tmpArr(0)), textLen, ByVal VarPtr(textArr(0)), _
                                    tmpLen, ByVal 0&, ByVal 0&)
        If tmpLen = 0 Then Exit Function
        ' a hopeless try to correct a weird error?
        ReDim Preserve textArr(tmpLen - 1)
    End If
    ' return the result
    UTF8toANSI = CStr(textArr)
End Function


Public Function UTF8BytesToUnicodeString(data() As Byte) As String
    Dim objStream As Object
    Dim strTmp As String
    On Error Resume Next
    
    ' init stream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8" 'We must set "utf-8" charset before open the stream
    objStream.Mode = 3 'adModeReadWrite
    objStream.Type = 1 'adTypeBinary
    objStream.Open
    
    ' write bytes into stream
    objStream.Write data
    objStream.Flush
    
    ' rewind stream and read texts
    objStream.position = 0
    objStream.Type = 2 'adTypeText
    strTmp = objStream.ReadText
    
    ' close up and return
    objStream.Close
    UTF8BytesToUnicodeString = strTmp

End Function

Public Function UnicodeStringToUTF8Bytes(strText As String) As Byte()
    Dim objStream As Object
    Dim data() As Byte
    On Error Resume Next
    
    ' init stream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Mode = 3 'adModeReadWrite
    objStream.Type = 2 ' adTypeText
    objStream.Open
    
    ' write strings into stream
    objStream.WriteText strText
    objStream.Flush
    
    ' rewind stream and read text
    objStream.position = 0
    objStream.Type = 1 'adTypeBinary
    'Test the UTF8 BOM 0xEF,0xBB,0xBF, skip first 3 bytes if BOM exists
    If Not HasUTF8BOM(objStream) Then
        objStream.position = 0
    End If
    data = objStream.Read()

    ' close up and return
    objStream.Close
    UnicodeStringToUTF8Bytes = data

End Function

'# Read a file as 'UTF8-encoded-text-file' and return result as Unicode string
Public Function ReadUTF8FileString(FileName As String) As String
    Dim objStream As Object
    Dim data As String
    On Error Resume Next
    
    ' init stream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Mode = 3 'adModeReadWrite
    objStream.Type = 2 ' adTypeText
    objStream.Open
    
    objStream.LoadFromFile FileName
    data = objStream.ReadText()
    
    ' close up and return
    objStream.Close
    ReadUTF8FileString = data

End Function

'# Read a file as binary file and return result as an array of bytes
'# If the file contains UTF8 BOM, these 3 bytes will NOT be included in resut.
Public Function ReadUTF8FileBytes(FileName As String) As Byte()
    Dim objStream As Object
    Dim data() As Byte
    On Error Resume Next
    
    ' init stream
    Set objStream = CreateObject("ADODB.Stream")
    'objStream.Charset = "utf-8"
    objStream.Mode = 3 'adModeReadWrite
    objStream.Type = 1 'adTypeBinary
    objStream.Open
    
    objStream.LoadFromFile FileName
    'Test the UTF8 BOM 0xEF,0xBB,0xBF, skip first 3 bytes if BOM exists
    If Not HasUTF8BOM(objStream) Then
        objStream.position = 0
    End If
    
    data = objStream.Read()
    'Call normalizeArray(data)
    
    ' close up and return
    objStream.Close
    ReadUTF8FileBytes = data

End Function

'# Write Unicode string to a 'UTF8-encoded-text-file'
Public Function WriteUTF8FileString(FileName As String, ByVal content As String)
    Dim objStream As Object
    On Error Resume Next
    
    ' init stream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Mode = 3 'adModeReadWrite
    objStream.Type = 2 ' adTypeText
    objStream.Open
    
    objStream.WriteText content
    objStream.SaveToFile FileName, 2 'adSaveCreateOverWrite
    
    ' close up and return
    objStream.Close

End Function

'# Write an array of UTF8 bytes to a binary file
'# This function will write UTF8 BOM (3 bytes &HEF &HBB &HBF) to the file
Public Function WriteUTF8FileBytes(FileName As String, content() As Byte)
    Dim objStream As Object
    On Error Resume Next
    
    ' init stream
    Set objStream = CreateObject("ADODB.Stream")
    'objStream.Charset = "utf-8"
    objStream.Mode = 3 'adModeReadWrite
    objStream.Type = 1 'adTypeBinary
    objStream.Open
    
    If ArraySize(content) > 3 And content(0) = &HEF And content(1) = &HBB And content(2) = &HBF Then
        'HasUTF8BOM
    Else
        Call writeUTF8BOM(objStream)
    End If
    
    objStream.Write content
    objStream.SaveToFile FileName, 2 'adSaveCreateOverWrite
    
    ' close up and return
    objStream.Close

End Function

'#objStream is an  opened readable ADODB.Stream
'#objStream must be opened as binary stream and the stream pointer must be at the begining
'#otherwise do nothing
Public Function writeUTF8BOM(ByVal objStream As Object)
    Dim BOM(2) As Byte 'an array holding 3 items
    
    On Error Resume Next
    BOM(0) = &HEF
    BOM(1) = &HBB
    BOM(2) = &HBF
    
    If objStream.position = 0 And objStream.Type = 1 Then
        objStream.Write BOM
    End If
End Function

'#objStream is an  opened readable ADODB.Stream
'#This function will detect if this stream has UTF8 BOM
'#Return True if has BOM; False otherwise
'#
'#Note: After calling this function, the stream pointer will be moved 3 bytes foreward
'#      if you want to read from the first byte, you must move the pointer back to 0
Private Function HasUTF8BOM(ByVal objStream As Object) As Boolean
    Dim BOM() As Byte
    
    On Error GoTo ErrHandler
    'Test the UTF8 BOM: &HEF, &HBB, &HBF, skip first 3 bytes if BOM exists
    BOM = objStream.Read(3)
    If ArraySize(BOM) = 3 And BOM(0) = &HEF And BOM(1) = &HBB And BOM(2) = &HBF Then
        HasUTF8BOM = True
    Else
        HasUTF8BOM = False
    End If
    Exit Function
    
ErrHandler:
    Debug.Print "Error occured with code " & Err.Number
    HasUTF8BOM = False
    Exit Function
    
End Function


Public Function ArraySize(anArray As Variant) As Integer
    ArraySize = 0
    On Error Resume Next
    ArraySize = UBound(anArray) - LBound(anArray)
    If (ArraySize > 0) Then
        ArraySize = ArraySize + 1
    Else
        ArraySize = 0
    End If
End Function

Public Function AddNullTerminatorToArray(anArray As Variant)
    On Error Resume Next
    Dim tmpLen As Integer
    tmpLen = ArraySize(anArray)
    'ReDim Preserve anArray(tmpLen) will preserve array value and add one more byte 'null'
    ReDim Preserve anArray(tmpLen)
    'anArray(tmpLen) = 0 (set to null terminator)
    
End Function

Public Function openDebugFile()
    On Error Resume Next
    If DEBUG_ENABLED Then
        Set debugfile = CreateObject("ADODB.Stream")
        debugfile.Type = 2 'Specify stream type - we want To save text/string data.
        debugfile.Charset = "utf-8" 'Specify charset For the source text data.
        debugfile.Mode = 3 'adModeReadWrite
        debugfile.Open 'Open the stream And write binary data To the object
        
        Set debugbinaryfile = CreateObject("ADODB.Stream")
        debugbinaryfile.Type = 1 'Specify stream type - binary file
        debugbinaryfile.Mode = 3 'adModeWrite
        debugbinaryfile.Open 'Open the stream And write binary data To the object
        
        Call writeUTF8BOM(debugbinaryfile)
    End If
End Function

Public Function closeDebugFile()
    On Error Resume Next
    If DEBUG_ENABLED Then
        debugfile.SaveToFile DEBUG_FILE, 2
        debugfile.Close
        
        debugbinaryfile.SaveToFile DEBUG_BINARY_FILE, 2
        debugbinaryfile.Close
    End If
End Function

Public Function DebugString(message As String)
    On Error Resume Next
    If DEBUG_ENABLED Then
        'WriteText 2th parameter can be 'adWriteLine 1' or 'adWriteChar 0', use 1 to write a 'EOL'
        debugfile.WriteText message, 1
    End If
End Function

Public Function DebugBinary(bytes() As Byte)
    On Error Resume Next
    If DEBUG_ENABLED Then
        debugbinaryfile.Write bytes
    End If
End Function

'DLL Initialization routine
Sub Main()
    On Error Resume Next
    'try to catch the DLL not found Exception
    On Error GoTo 0
End Sub


