VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pStringOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Function ByteArrayToString(bytArray() As Byte)
    On Error GoTo ErrHandler
    
    Dim sAns As String
    Dim iPos As String
    
    sAns = StrConv(bytArray, vbUnicode)
    iPos = InStr(sAns, Chr(0))
    If iPos > 0 Then sAns = Left(sAns, iPos - 1)
    
    ByteArrayToString = sAns
    Exit Function
    
ErrHandler:
    ByteArrayToString = 0
    Exit Function
End Function

Public Function BytesToMegabytes(Bytes As Double) As Double
Attribute BytesToMegabytes.VB_Description = "Converts a figure represented in bytes to the corresponding figure in megabytes."
    Dim dblAns As Double
    
    dblAns = (Bytes / 1024) / 1024
    BytesToMegabytes = Format(dblAns, "###,###,##0.00")
End Function
 
Public Function Decrypt(Expression As String, Key As String)
    'Decrypts the specified string using the given numerical decryption key
    On Error GoTo ErrHandler

    'Declare variables
    Dim i As Long
    Dim sChar As String
    Dim lChrCrypt As Long
    Dim sFront As String
    Dim sBack As String

    Expression = Trim(Expression)

    'Decrypt the string
    lChrCrypt = 0
    For i = 1 To Len(Expression)
        lChrCrypt = lChrCrypt + 1
        If lChrCrypt > Len(Key) Then lChrCrypt = 1

        'Pull the string apart and decrypt one character at a time by moving it down the ASCII chart
        sChar = Mid(Expression, i, 1)
        
        'Change to prevent geting Ascii Code < 1
        'Thanks to Patrick Di Martino [pitc@costarica.net] for this fix!
        If Asc(sChar) - Mid(Key, lChrCrypt, 1) < 1 Then
            sChar = Chr(Asc(sChar) - Mid(Key, lChrCrypt, 1) + 255)
        Else
            sChar = Chr(Asc(sChar) - Mid(Key, lChrCrypt, 1))
        End If
        
        sFront = Left(Expression, i - 1)
        sBack = Right(Expression, Len(Expression) - i)

        'Put the string back together and move on to the next character
        Expression = sFront & sChar & sBack
        DoEvents
    Next

    'Return the encrypted string
    Decrypt = Expression
    Exit Function

ErrHandler:
    Decrypt = 0
    Exit Function
End Function

Public Function Encrypt(Expression As String, Key As String)
Attribute Encrypt.VB_Description = "Encrypts the specified expression using the specified encryption key."
    'Encrypts the specified string using the given numerical encryption key
    On Error GoTo ErrHandler

    'Declare variables
    Dim i As Long
    Dim sChar As String
    Dim lChrCrypt As Long
    Dim sFront As String
    Dim sBack As String
    
    Expression = Trim(Expression)
    
    'Encrypt the string
    lChrCrypt = 0
    For i = 1 To Len(Expression)
        lChrCrypt = lChrCrypt + 1
        If lChrCrypt > Len(Key) Then lChrCrypt = 1
        
        'Pull the string apart and encrypt one character at a time by moving it up the ASCII chart
        sChar = Mid(Expression, i, 1)
        
        'Change to prevent Ascii values > 255
        'Thanks to Patrick Di Martino [pitc@costarica.net] for this fix!
        If Asc(sChar) + Mid(Key, lChrCrypt, 1) > 255 Then
            sChar = Chr(Asc(sChar) + Mid(Key, lChrCrypt, 1) - 255)
        Else
            sChar = Chr(Asc(sChar) + Mid(Key, lChrCrypt, 1))
        End If
        
        sFront = Left(Expression, i - 1)
        sBack = Right(Expression, Len(Expression) - i)
        
        'Put the string back together and move on to the next character
        Expression = sFront & sChar & sBack
        DoEvents
    Next
    
    'Return the encrypted string
    Encrypt = Expression
    Exit Function
    
ErrHandler:
    Encrypt = 0
    Exit Function
End Function

Public Function GetAllAfter(Expression As String, Delimiter As String)
Attribute GetAllAfter.VB_Description = "Returns all text in the expression after the specified delimiter."
    'Returns all text after the first instance of the specified delimiter character(s)
    On Error GoTo ErrHandler
    
    'Declare variables
    Dim lCharPos As Long
    
    'Return the text after the delimiter
    lCharPos = InStr(1, Expression, Delimiter, vbTextCompare) + (Len(Delimiter) - 1)
    GetAllAfter = Right(Expression, Len(Expression) - lCharPos)
    Exit Function
    
ErrHandler:
    GetAllAfter = 0
    Exit Function
End Function

Public Function GetAllBefore(Expression As String, Delimiter As String)
Attribute GetAllBefore.VB_Description = "Returns all text in the expression before the specified delimiter."
    'Returns all text before the first instance of the specified delimiter character(s)
    On Error GoTo ErrHandler
    
    'Declare variables
    Dim lCharPos As Long
    
    'Return the text after the delimiter
    lCharPos = InStr(1, Expression, Delimiter, vbTextCompare)
    GetAllBefore = Left(Expression, lCharPos - 1)
    Exit Function
    
ErrHandler:
    GetAllBefore = 0
    Exit Function
End Function

Public Function GetDelimited(Expression As String, ReturnCollection As Collection, OpenDelimiter As String, CloseDelimiter As String)
Attribute GetDelimited.VB_Description = "Returns a collection containing all instances of text in the expression between the two specified delimiters."
    'Returns a collection containing each instance of text between the specified delimiters
    On Error GoTo ErrHandler
    
    'Declare variables
    Dim startPos As Long
    Dim endPos As Long
    Dim tmpString As String
    Dim i As Long
    
    'Set variable values
    tmpString = Expression
    startPos = 1
    
    'Get the text between the delimiters
    Do While startPos > 0
        startPos = InStr(startPos, Expression, OpenDelimiter, vbTextCompare)
        endPos = InStr(startPos + 1, Expression, CloseDelimiter, vbTextCompare)
        
        If (startPos > 0) And (endPos > startPos) Then
            tmpString = Mid(Expression, startPos + Len(OpenDelimiter), (endPos - startPos) - Len(OpenDelimiter))
            
            'Populate the specified collection with every instance of delimited text
            ReturnCollection.Add tmpString
            startPos = startPos + 1
        Else
            startPos = 0
        End If
        DoEvents
    Loop
    
    'Return the new string.
    Exit Function
    
ErrHandler:
    GetDelimited = 0
    Exit Function
End Function

Public Function GetDelimitedSingle(Expression As String, OpenDelimiter As String, CloseDelimiter As String)
Attribute GetDelimitedSingle.VB_Description = "Returns a string containing only the first instance of text in the expression between the two specified delimiters."
    'Returns the first instance of text between the specified delimiters
    On Error GoTo ErrHandler
    
    'Declare variables
    Dim tmpString As String
    Dim lOpenCharPos As Long
    Dim lCloseCharPos As Long
    
    'Set variable values
    tmpString = Expression
    
    lOpenCharPos = InStr(1, tmpString, OpenDelimiter, vbTextCompare) + (Len(OpenDelimiter) - 1)
    tmpString = Right(tmpString, Len(tmpString) - lOpenCharPos)
    
    lCloseCharPos = InStr(1, tmpString, CloseDelimiter, vbTextCompare)
    tmpString = Left(tmpString, lCloseCharPos - 1)
    
    'Return the new string.
    GetDelimitedSingle = tmpString
    Exit Function
    
ErrHandler:
    GetDelimitedSingle = 0
    Exit Function
End Function

Public Function GetFilenameFromPath(FullPath As String) As String
Attribute GetFilenameFromPath.VB_Description = "Returns a filename only, given the file's full path."
    'Returns a filename only, given the file's full path.
  

    Dim sPath As String
    Dim sList() As String
    Dim sAns As String
    Dim iArrayLen As Integer

    If Len(FullPath) = 0 Then Exit Function
    sList = Split(FullPath, "\")
    iArrayLen = UBound(sList)
    sAns = IIf(iArrayLen = 0, "", sList(iArrayLen))
    
    GetFilenameFromPath = sAns
End Function

Public Function IsAlphaNumeric(sChr As String) As Boolean
Attribute IsAlphaNumeric.VB_Description = "Determines whether or not the specified character is alphanumeric."
    'Determines whether or not a character is alphanumeric
  
    IsAlphaNumeric = sChr Like "[0-9A-Za-z]"
End Function

Public Function IsEven(Number As Long) As Boolean
Attribute IsEven.VB_Description = "Determines whether the specified number is even or odd."
    'Determines whether the specified number is even or odd.
    IsEven = (Number Mod 2 = 0)
End Function

Public Function IsValidCreditCardNumber(CardNumber As String) As Boolean
   CardNumber = Replace(CardNumber, "-", "")
   CardNumber = Replace(CardNumber, " ", "")
   
   Const MAX_DIGITS = 20  ' actually don't know any
                            'card using more than 16 digits
   
   Dim anDigits(1 To MAX_DIGITS) As Byte
   Dim nDigits As Long
   
   Dim ofsCurrentDigit As Long
   Dim ofsCurrentCharacter As Long
   
   Dim CurrentCharacter As String
   
   Dim Multiplier As Long
   Dim CheckSum As Long
   Dim DigitValue As Long
   Dim Result As Boolean
   
   Dim ValidDigits As String
   
   If Not IsNumeric(CardNumber) Then Exit Function
   If Len(Trim$(CardNumber)) < 1 Then
      Result = False
      GoTo Exit_Point
   End If
   
   ValidDigits = "0123456789"
   
   For ofsCurrentCharacter = 1 To Len(CardNumber)
      CurrentCharacter = Mid$(CardNumber, ofsCurrentCharacter, 1)
      If InStr(1, ValidDigits, CurrentCharacter, vbBinaryCompare) Then
         nDigits = nDigits + 1
         If nDigits > MAX_DIGITS Then
            Result = False
            GoTo Exit_Point
         End If
         anDigits(nDigits) = Val(CurrentCharacter)
      End If
   Next ofsCurrentCharacter
   
   CheckSum = anDigits(nDigits)
   
   For ofsCurrentDigit = nDigits - 1 To 1 Step -1
      
      If Multiplier = 2 Then
         Multiplier = 1
      Else
         Multiplier = 2
      End If
      DigitValue = anDigits(ofsCurrentDigit) * Multiplier
      CheckSum = CheckSum + DigitValue
      If DigitValue > 9 Then
         CheckSum = CheckSum - 9
      End If
      
   Next ofsCurrentDigit
   
   Result = ((CheckSum Mod 10) = 0)
   
Exit_Point:
   
   IsValidCreditCardNumber = Result
   Exit Function
   
End Function

Public Function IsValidZipCode(ZipString As String) As Boolean
    'Determines whether or not a string of numbers is a valid U.S. zip code.
  
    Dim iLen As Integer
    Dim bAns As Boolean
    
    iLen = Len(ZipString)
    
    Select Case iLen
        Case 5
           bAns = ZipString Like "#####"
        Case 9
            bAns = ZipString Like "#########"
        Case 10
            bAns = ZipString Like "#####-####"
        End Select
                
    IsValidZipCode = bAns
End Function

Public Function MegabytesToBytes(Megabytes As Double) As Double
   'Converts a figure represented in megabytes to the corresponding figure in bytes.
    Dim dblAns As Double
    
    dblAns = (Megabytes * 1024) * 1024
    MegabytesToBytes = Format(dblAns, "###,###,##0.00")
End Function
 
Public Function PhoneFormat(ByVal PhoneNumber As String) As String
    'Converts a number in any format to a phone number in the format '(###) ###-####'. If
    'the code can't figure out what format your number is in, it will simply return the
    'number you gave it.
  

  Dim strResult As String
  Dim iLength As Integer
  Dim strExtraChar As String
  Dim strOriginal As String
  Dim iSpaceResult As Integer
  Dim i As Integer
  
  strOriginal = PhoneNumber
      
  ' Remove any style characters from the user input
  PhoneNumber = Replace(PhoneNumber, ")", "")
  PhoneNumber = Replace(PhoneNumber, "(", "")
  PhoneNumber = Replace(PhoneNumber, "-", "")
  PhoneNumber = Replace(PhoneNumber, ".", "")
  PhoneNumber = Replace(PhoneNumber, Space(1), "")
      
  iLength = Len(PhoneNumber)
  
  'convert any letters to numbers
  For i = 1 To iLength
    Mid$(PhoneNumber, i, i) = _
        PhoneLetterToDigit(Mid$(PhoneNumber, i, i))
  Next i
  
  ' now, if any other chars besides numbers exist, return original string to user
  For i = 1 To iLength
    Select Case Asc(Mid$(PhoneNumber, i, i))
      Case Is < 48, Is > 57
        strResult = strOriginal
    End Select
  Next i
  
  Select Case iLength
' user entered a lot of numbers;only format the first 10
    Case Is > 11
      If Left$(PhoneNumber, 1) = "1" Then
        strExtraChar = Mid$(PhoneNumber, 12)
        PhoneNumber = Mid$(PhoneNumber, 2, 10)
      Else
        strExtraChar = Mid$(PhoneNumber, 11)
        PhoneNumber = Mid$(PhoneNumber, 1, 10)
      End If
 
' if user included the number 1 before the area code.
'We drop this number
   
    Case Is = 11
      If Left$(PhoneNumber, 1) = "1" Then
        PhoneNumber = Mid$(PhoneNumber, 2)
      Else
        ' check for a space character
        iSpaceResult = InStrRev(strOriginal, Space(1))
        
        If iSpaceResult = 0 Then
          ' we have no idea what they entered
          strResult = strOriginal
          GoTo Exit_Proc
        Else
          strExtraChar = Mid$(PhoneNumber, iSpaceResult)
          PhoneNumber = Mid$(PhoneNumber, 1, iSpaceResult - 1)
        End If
      
      End If
    
    Case Is = 10 ' area code and phone
      PhoneNumber = PhoneNumber
 ' user did not include an area code; add 3 spaces
         
    Case Is = 7
        PhoneNumber = Space(3) & PhoneNumber
 
   ' unable to figure out what the user typed
   ' must be an extentsion and not a 'real' phone number

      Case Else
         strResult = strOriginal
         GoTo Exit_Proc
  
  End Select
    
  'Add sytle characters into phone number (format)
  strResult = Format(PhoneNumber, "\(@@@\)\ @@@\-@@@@") & Space(1) & strExtraChar
 
Exit_Proc:
  PhoneFormat = strResult
    
End Function

Function PhoneLetterToDigit(ByVal PhoneLetter As String) As String
Attribute PhoneLetterToDigit.VB_Description = "Converts letters to the corresponding digits on a telephone dialpad."
  'Converts a letter to the corresponding digit on a telephone dialpad.

  Dim intDigit As Integer
  
  intDigit = Asc(UCase$(PhoneLetter))
    
  If intDigit >= 65 And intDigit <= 90 Then

    If intDigit = 81 Or 90 Then ' Q or Z
      intDigit = intDigit - 1
    End If

    intDigit = (((intDigit - 65) \ 3) + 2)
    PhoneLetterToDigit = intDigit
  Else
    PhoneLetterToDigit = PhoneLetter
  End If

End Function

Public Function RandomNumber(u_bound As Long, Optional l_bound As Long = 0) As Long
Attribute RandomNumber.VB_Description = "Returns a MUCH more random number than VB's RND function. Pass in the upper bound, as well as the optional lower bound values for the range in which you wish to select the random number in."
    'Returns a MUCH more random number than VB's RND function.
  
    
    Dim lngSeed As Long
    Dim i As Integer, intIterCount As Integer
    
    lngSeed = ((u_bound - l_bound + 1) * Rnd + l_bound) * Second(Now)
    If lngSeed = 0 Then
        lngSeed = (u_bound - l_bound + 1) * Rnd + l_bound
    End If
    
    If Len(CStr(lngSeed * lngSeed)) > 4 Then
        intIterCount = CInt(Left(CStr(lngSeed * lngSeed), 4))
    Else
        intIterCount = lngSeed * lngSeed
    End If
    
    For i = 0 To intIterCount
        If Len(CStr(lngSeed * lngSeed)) > 4 Then
            lngSeed = CLng(Mid(CStr(lngSeed * lngSeed), 2, 2))
        Else
            lngSeed = CLng(Left(CStr(lngSeed * lngSeed), 2)) + 1
        End If
    Next i
    
    RandomNumber = (u_bound - l_bound) * (lngSeed / 100) + l_bound
End Function

Public Function RandomString(Expression As String)
    'Rearranges the characters in the string randomly.
    On Error GoTo ErrHandler
    
    Dim i As Long
    Dim sFront As String
    Dim sBack As String
    Dim sChar As String
    
    'Randomize the characters in the string
    For i = 1 To Len(Expression)
        'Pull the string apart
        sChar = Mid(Expression, i, 1)
        sFront = Left(Expression, i - 1)
        sBack = Right(Expression, Len(Expression) - i)
        
        'Put it back together
        Expression = sFront & sBack & sChar
        DoEvents
    Next
    
    RandomString = Expression
    Exit Function
    
ErrHandler:
    RandomString = 0
    Exit Function
End Function

Public Function ReverseString(Expression As String) As String
Attribute ReverseString.VB_Description = "Reverses a string"
    'Reverses a string
  
    
    Dim lLen As Long, lCtr As Long
    Dim sChar As String
    Dim sAns As String
    
    lLen = Len(Expression)
    For lCtr = lLen To 1 Step -1
        sChar = Mid(Expression, lCtr, 1)
        sAns = sAns & sChar
    Next
    
    ReverseString = sAns

End Function

Public Function SearchReplace(Expression As String, Find As String, sReplace As String)
Attribute SearchReplace.VB_Description = "Searches the expression for the specified string and, if found, replaces every occurance with the specified replacement string."
    'Searches a string for the specified text and replaces it with the specified replacement text.
    On Error GoTo ErrHandler
    
    Dim tmpString As String
    tmpString = Expression
    SearchReplace = Replace(tmpString, Find, sReplace)
    Exit Function
    
ErrHandler:
    SearchReplace = 0
    Exit Function
End Function

Public Function StripDelimited(Expression As String, OpenDelimiter As String, CloseDelimiter As String)
Attribute StripDelimited.VB_Description = "Removes all instances of text from the expression that are between the specified delimiters."
    'Strips all text from a string which is inside the specified delimiters. Can be used
    'to strip HTML tags from a document and other such things. Returns the new string if
    'successful, or 0 for failure.
    On Error GoTo ErrHandler
    
    'Declare variables
    Dim startPos As Long
    Dim oldStartPos As Long
    Dim endPos As Long
    Dim sFront As String
    Dim sBack As String
    
    'Set variable values
    startPos = 1
    
    'Perform the strip
    Do While startPos > 0
        startPos = InStr(startPos, Expression, OpenDelimiter, vbTextCompare)
        endPos = InStr(startPos + 1, Expression, CloseDelimiter, vbTextCompare)
        
        If (startPos > 0) And (endPos > startPos) Then
            sFront = Left(Expression, startPos - 1)
            sBack = Right(Expression, Len(Expression) - (endPos + (Len(CloseDelimiter) - 1)))
            Expression = sFront & sBack
        Else
            startPos = 0
        End If
        DoEvents
    Loop
    
    'Return the new string.
    StripDelimited = Expression
    Exit Function
    
ErrHandler:
    StripDelimited = 0
    Exit Function
End Function

Public Function StripSingle(Expression As String, Strip As String)
Attribute StripSingle.VB_Description = "Removes all instances of a single character from the expression."
    'Strips the specified character(s) from a string. Returns the new string if
    'successful, or 0 for failure.
    On Error GoTo ErrHandler
    
    'Declare variables
    Dim startPos As Long
    Dim sFront As String
    Dim sBack As String
    
    'Set variable values
    startPos = 1
    
    'Perform the strip
    Do While startPos > 0
        startPos = InStr(startPos, Expression, Strip, vbTextCompare)
        
        If startPos > 0 Then
            sFront = Left(Expression, startPos - 1)
            sBack = Right(Expression, Len(Expression) - (startPos + (Len(Strip) - 1)))
            Expression = sFront & sBack
        Else
            startPos = 0
        End If
        DoEvents
    Loop
    
    'Return the new string.
    StripSingle = Expression
    Exit Function
    
ErrHandler:
    StripSingle = 0
    Exit Function
End Function

Public Function StripSpaces(Expression As String) As String
Attribute StripSpaces.VB_Description = "Removes all spaces from the specified string."
    'Removes all spaces from the specified string.
  

    Dim sAns As String
    Dim lLen As String
    Dim lCtr As Long, lCtr2 As Long
    Dim sChar As String
    
    lLen = Len(Expression)
    sAns = Expression
    lCtr2 = 1
    
    For lCtr = 1 To lLen
        sChar = Mid(Expression, lCtr, 1)
        If sChar <> " " Then
            Mid(sAns, lCtr2, 1) = sChar
            lCtr2 = lCtr2 + 1
        End If
    Next
    
    If lCtr2 > 1 Then
        sAns = Left(sAns, lCtr2 - 1)
    Else
        sAns = ""
    End If
    
    StripSpaces = sAns
End Function

Public Function StripSpecialCharacters(Expression As String) As String
Attribute StripSpecialCharacters.VB_Description = "This function strips all characters with an ASCII value of 128 or higher from a string. In other words, it converts an 8-bit ASCII value into a 7-bit ASCII value."
    'This function strips all characters with an ASCII value of 128 or higher from a string.
    'In other words, it converts an 8-bit ASCII value into a 7-bit ASCII value.
  

    Dim sWkg As String, sOrigString As String, sNewString As String
    
    Dim lLen As Long
    Dim lCtr As Long, lCtr2 As Long
    Dim sChar As String
    
    lLen = Len(Expression)
    'create buffer
    
    sOrigString = Space$(lLen)
    sOrigString = Expression
    sNewString = sOrigString
    lCtr2 = 1
    
    For lCtr = 1 To lLen
        sChar = Mid(sOrigString, lCtr, 1)
        If Asc(sChar) < 128 Then
            Mid(sNewString, lCtr2, 1) = sChar
            lCtr2 = lCtr2 + 1
        End If
    Next
    
    If lCtr2 > 1 Then
        sNewString = Left(sNewString, lCtr2 - 1)
    Else
        sNewString = ""
    End If
    
    StripSpecialCharacters = sNewString
End Function

Public Function TrimWithoutPrejudice(Expression As String) As String
Attribute TrimWithoutPrejudice.VB_Description = "VB's Trim function only trims space characters ("" "") from the beginning or the end of a string. This alternative trims all nonprinting characters (tabs, character returns, etc.)"
    'VB's Trim function only trims space characters (" ") from the beginning or
    'the end of a string. This alternative trims all nonprinting characters
    '(tabs, character returns, etc.)
  
    
    Dim sAns As String
    Dim sWkg As String
    Dim sChar As String
    Dim lLen As Long
    Dim lCtr As Long
    
    sAns = Expression
    lLen = Len(Expression)
    
    If lLen > 0 Then
    'Ltrim
        For lCtr = 1 To lLen
            sChar = Mid(sAns, lCtr, 1)
            If Asc(sChar) > 32 Then Exit For
        Next
    
    sAns = Mid(sAns, lCtr)
    lLen = Len(sAns)
    
    'Rtrim
        If lLen > 0 Then
            For lCtr = lLen To 1 Step -1
                sChar = Mid(sAns, lCtr, 1)
                If Asc(sChar) > 32 Then Exit For
            Next
        End If
        sAns = Left$(sAns, lCtr)
    End If
    
    TrimWithoutPrejudice = sAns

End Function

