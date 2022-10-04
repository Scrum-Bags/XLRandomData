Attribute VB_Name = "M_SB_DataGeneration"

'Parameterizes random integer generation for other functions
Function RandomInteger(min, max)
    Randomize
    RandomInteger = Int((max - min + 1) * Rnd(Time())) + min
End Function

'Randomly returns one item from the input array; useful for other functions
Function RandArrayItem(arr)
    RandArrayItem = arr(RandomInteger(LBound(arr), UBound(arr)))
End Function

' Generate a random uppercase alphabet character
Function RandomUpperAlpha()
    RandomUpperAlpha = RandArrayItem(Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"))
End Function

' Generate a random lowercase alphabet character
Function RandomLowerAlpha()
    RandomLowerAlpha = RandArrayItem(Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"))
End Function

' Generate a random numeral
Function RandomNumeral()
    RandomNumeral = RandArrayItem(Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9"))
End Function

' Generate random symbol
Function RandomSymbol()
    RandomSymbol = RandArrayItem(Array("$", "@", "_"))
End Function

' Generate a random character with optional selectors
Function RandomCharacter(Optional ByVal includeLowerAlphas As Boolean = True, Optional ByVal includeUpperAlphas As Boolean = False, Optional ByVal includeNumerals As Boolean = False, Optional ByVal includeSymbols As Boolean = False)
    Dim options As Object
    Set options = CreateObject("System.Collections.ArrayList")
    If includeLowerAlphas Then
        options.Add RandomLowerAlpha()
    End If
    If includeUpperAlphas Then
        options.Add RandomUpperAlpha()
    End If
    If includeNumerals Then
        options.Add RandomNumeral()
    End If
    If includeSymbols Then
        options.Add RandomSymbol()
    End If
    RandomCharacter = RandArrayItem(options.ToArray)
End Function

' Generate a randomized string
Function RandomString(Optional ByVal length As Integer = 10, Optional ByVal includeLowerAlphas As Boolean = True, Optional ByVal includeUpperAlphas As Boolean = False, Optional ByVal includeNumerals As Boolean = False, Optional ByVal includeSymbols As Boolean = False)
    Dim outString
    outString = ""
    For i = 1 To length Step 1
        outString = outString & RandomCharacter(includeLowerAlphas, includeUpperAlphas, includeNumerals, includeSymbols)
    Next
    RandomString = outString
End Function

' Randomly selects a user's input first name
Function GetRandomFirstName()
    GetRandomFirstName = RandArrayItem(Array("Aaron", "Barrett", "Cathy", "Dylan", "Edgar", "Fulton", "Gary", "Hector", "Isabelle", "Jeremy", "Kacey", "Lucy", "Marcus", "Noelle", "Orianna", "Peter", "Quandale", "Rochelle", "Stacy", "Tucker", "Uther", "Vanessa", "Walter", "Xavier", "Yorick", "Zane"))
End Function

' Randomly selects a user's input last name
Function GetRandomLastName()
    GetRandomLastName = RandArrayItem(Array("Aaronson", "Brown", "Christopher", "Dingle", "Eames", "Flowers", "Grissom", "Howards", "Irons", "James", "Kraus", "Lars", "McPherson", "Neilson", "Orville", "Parker", "Quincy", "Rathers", "Singer", "Todd", "Ulric", "Veers", "Walden", "Xerxes", "Yoko", "Zachary"))
End Function

' Returns a randomly selected email domain
Function RandomEmailDomain()
    RandomEmailDomain = RandArrayItem(Array("gmail.com", "hotmail.com", "tutanota.com", "yahoo.com", "comcast.net", "protonmail.com", "aol.com", "ymail.com", "outlook.com", "verizon.net", "att.net", "mail.com"))
End Function

'Generates a random valid email address
Function GetRandomValidEmailAddress()
    Dim EmailAddress
    EmailNameLength = RandomInteger(6, 24)
    EmailAddress = RandomString(EmailNameLength, True, True, True, False)
    EmailAddress = EmailAddress & "@" & RandomEmailDomain()
    GetRandomValidEmailAddress = EmailAddress
End Function

'Generates a random INvalid email address
Function GetRandomInvalidEmailAddress()
    Dim EmailAddress
    EmailNameLength = RandomInteger(6, 24)
    EmailAddress = RandomString(EmailNameLength, True, True, True, False)
    InvalidityReason = RandomInteger(1, 3)
    Select Case InvalidityReason
        Case 1
            'no '@'
            EmailAddress = EmailAddress & RandomEmailDomain()
        Case 2
            'no site
            EmailAddress = EmailAddress & "@"
        Case 3
            'no name in front
            EmailAddress = "@" & RandomEmailDomain()
        Case Else
    End Select
    GetRandomInvalidEmailAddress = EmailAddress
End Function

'Randomly selects one of Aline's gender options
Function GetRandomAlineGender()
    GetRandomGender = RandArrayItem(Array("Male", "Female", "Other", "Prefer not to say..."))
End Function

'Randomly generates a level of income with an accompanying schedule
Function GetRandomPayAndSchedule()
    AnnualValue = RandomInteger(80, 350) * 1000
    PaySchedule = RandomInteger(1, 5)
    Dim DollarValue
    Dim ScheduleDescription
    Select Case PaySchedule
        Case 2
            DollarValue = AnnualValue / 12
            ScheduleDescription = "Monthly"
        Case 3
            DollarValue = AnnualValue / 25
            ScheduleDescription = "Bi-Weekly"
        Case 4
            DollarValue = AnnualValue / 50
            ScheduleDescription = "Weekly"
        Case 5
            DollarValue = AnnualValue / 2000
            ScheduleDescription = "Hourly"
        Case Else
            DollarValue = AnnualValue
            ScheduleDescription = "Annually"
    End Select
    GetRandomPayAndSchedule = Array(DollarValue, ScheduleDescription)
End Function

' Randomly generates an income amount
Function GetRandomIncomeAmount()
    GetRandomIncomeAmount = RandomInteger(50, 350) * 1000
End Function

' Randomly generates a pay schedule
Function GetRandomPaySchedule()
    GetRandomPaySchedule = RandArrayItem(Array("Monthly", "Bi-Weekly", "Weekly", "Hourly", "Annually"))
End Function

'Randomly generates a date with formatting options
Function GetRandomDate(ByVal yearLowerBound As Integer, ByVal yearUpperBound As Integer, Optional ByVal formatStr As String = "MM-DD-YYYY")
    RandomYear = CStr(RandomInteger(yearLowerBound, yearUpperBound))
    RandomMonth = RandArrayItem(Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"))
    Dim DayCount
    Select Case RandomMonth
        Case "01", "03", "05", "07", "08", "10", "12"
            DayCount = 31
        Case "04", "06", "09", "11"
            DayCount = 30
        Case Else
            If RandomYear Mod 4 = 0 Then
                DayCount = 29
            Else
                DayCount = 28
            End If
    End Select
    RandomDay = CStr(RandomInteger(1, DayCount))
    If Len(RandomDay) < 2 Then
        RandomDay = "0" & RandomDay
    End If
    outString = formatStr
    outString = Replace(outString, "YYYY", RandomYear)
    outString = Replace(outString, "DD", RandomDay)
    outString = Replace(outString, "MM", RandomMonth)
    GetRandomDate = outString
End Function

' Combines other functions into a complete random address
Function GetRandomAddress()
    GetRandomAddress = GetRandomAddressNumber() & " " & GetRandomStreet()
End Function

'Generates randomly a street address unit number
Function GetRandomAddressNumber()
    GetRandomAddressNumber = CStr(RandomInteger(1, 999))
End Function

'Uses PRNG to randomly select a street name
Function GetRandomStreet()
    streetName = RandArrayItem(Array("1st", "2nd", "3rd", "4th", "5th", "6th", "7th", "8th", "Oak", "Maple", "Pine", "Elm", "Church", "Walnut", "Center", "High", "Park", "Cedar", "North", "South", "East", "West", "Sunset", "River", "Chestnut", "Ridge", "Mill", "Cherry", "Lakeview", "Spring", "Pearl"))
    StreetTitle = RandArrayItem(Array("St", "Rd", "Ave"))
    GetRandomStreet = streetName & " " & StreetTitle
End Function

'Uses PRNG to randomly select a city name
Function GetRandomCity()
    GetRandomCity = RandArrayItem(Array("Franklin", "Clinton", "Arlington", "Madison", "Washington", "Centerville", "Lebanon", "Georgetown", "Springfield", "Chester", "Fairview", "Greenville", "Bristol", "Dayton", "Dover", "Salem", "Winchester", "Oakland", "Milton", "Newport", "Ashland", "Bloomington", "Riverside", "Manchester", "Oxford", "Burlington", "Jackson", "Milford", "Clayton", "Kingston", "Auburn", "Lexington"))
End Function

'Uses PRNG to randomly select a U.S. State
Function GetRandomState()
    GetRandomState = RandArrayItem(Array("Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado", "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming"))
End Function

'Uses PRNG to randomly select a U.S. State Mailing Code
Function GetRandomStateCode()
    GetRandomStateCode = RandArrayItem(Array("AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY"))
End Function

'Uses PRNG to generate a random valid ZIP code
Function GetRandomValidZip()
    GetRandomValidZip = CStr(RandomInteger(10000, 99999))
End Function

'Uses PRNG to generate a random INvalid ZIP code
Function GetRandomInvalidZip()
    GetRandomInvalidZip = CStr(RandomInteger(1, 9999))
End Function

' Formats a sequence of 10 digits as a phone number with a specified format
' To use a custom format for the phone number,
' pass as an argument a new string that contains
' AAA where you want the number's area code
' BBB where you want the central office code
' CCCC where you want the line number
' e.g. "(AAA) BBB-CCCC"
Function FormatPhoneNumberString(ByVal PhoneNumberString As String, Optional ByVal numberFormat As String = "AAA-BBB-CCCC")
    Dim outString
    outString = numberFormat
    outString = Replace(outString, "AAA", mid(PhoneNumberString, 1, 3))
    outString = Replace(outString, "BBB", mid(PhoneNumberString, 4, 3))
    FormatPhoneNumberString = Replace(outString, "CCCC", mid(PhoneNumberString, 7, 4))
End Function

' Uses PRNG to create a phone number with formatting options
Function GetRandomValidPhoneNumber(Optional ByVal numberFormat As String = "AAA-BBB-CCCC")
    Dim PhoneNumberString
    PhoneNumberString = numberFormat
    AreaCodeLastTwo = "11"
    While AreaCodeLastTwo = "11"
        AreaCodeLastTwo = CStr(RandomInteger(0, 19))
    Wend
    If Len(AreaCodeLastTwo) < 2 Then
        AreaCodeLastTwo = "0" & AreaCodeLastTwo
    End If
    AreaCode = CStr(RandomInteger(2, 9)) & AreaCodeLastTwo
    CentralOfficeCode = CStr(RandomInteger(2, 9)) & CStr(RandomInteger(20, 99))
    LineNumber = RandomString(4, False, False, True, False)
    GetRandomValidPhoneNumber = FormatPhoneNumberString(AreaCode & CentralOfficeCode & LineNumber, numberFormat)
End Function

'Uses PRNG to create an invalid phone number
Function GetRandomInvalidPhoneNumber(Optional ByVal numberFormat As String = "AAA-BBB-CCCC")
    Dim PhoneNumberString
    PhoneNumberString = GetRandomValidPhoneNumber("AAABBBCCCC")
    InvalidityReason = RandomInteger(1, 2)
    Select Case InvalidityReason
        Case 1
            'incorrect length
            PhoneNumberString = mid(PhoneNumberString, 1, 9 - RandomInteger(0, 3))
        Case 2
            'bad character
            PhoneNumberString = Replace(PhoneNumberString, mid(PhoneNumberString, RandomInteger(1, Len(PhoneNumberString)), 1), RandomCharacter())
        Case Else
    End Select
    GetRandomInvalidPhoneNumber = FormatPhoneNumberString(PhoneNumberString, numberFormat)
End Function

' Generates a boolean with the passed probability / 1.0
Function GetRandomBool(probability As Double) As Boolean
    Roll = Rnd()
    Dim Same As Boolean
    Same = False
    If Roll < probability Then
        Same = True
    End If
    GetRandomBool = Same
End Function

' Wrap a boolean value as a true/false string
Function TFString(inputBool As Boolean, Optional ByVal capitalize As Boolean = False) As String
    Dim outString
    If inputBool Then
        If capitalize Then
            outString = "True"
        Else
            outString = "true"
        End If
    Else
        If capitalize Then
            outString = "False"
        Else
            outString = "false"
        End If
    End If
    TFString = outString
End Function

' Generate a SSN !NUMBER! & have another function for zero-padding?
'Uses PRNG to create a random validly formatted SSN
Function GetRandomValidSSN()
    ' ssn = RandomInteger(1, 999999999)
    ssn = ""
    For Iterator = 1 To 9 Step 1
        ssn = ssn & RandomNumeral()
    Next
    GetRandomValidSSN = """" & CStr(ssn) & """"
End Function

'Uses PRNG to create a random INvalidly formatted SSN
Function GetRandomInvalidSSN()
    ssn = ""
    Count = RandArrayItem(Array(6, 7, 8))
    For Iterator = 1 To Count Step 1
        ssn = ssn & RandomNumeral()
    Next
    GetRandomInvalidSSN = """" & CStr(ssn) & """"
End Function

'Uses PRNG to create a random validly formatted Driver's License ID
Function GetRandomValidDriversID()
    Dim IdString
    IdString = RandomCharacter(True, False, False, False)
    For Iterator = 1 To 11 Step 1
        IdString = IdString & RandomInteger(0, 9)
    Next
    GetRandomValidDriversID = IdString
End Function

'Uses PRNG to create a random INvalidly formatted Driver's License ID (too short)
Function GetRandomInvalidDriversID()
    Dim IdString
    IdString = RandomCharacter()
    GetRandomInvalidDriversID = IdString
End Function

'Uses PRNG to create a random username
Function GetRandomUserName(ByVal minLength As Integer, ByVal maxLength As Integer, ByVal includeLowerAlphas As Boolean, ByVal includeUpperAlphas As Boolean, ByVal includeNumerals As Boolean, ByVal includeSymbols As Boolean)
    Dim UserName
    GetRandomUserName = CStr(RandomString(RandomInteger(minLength, maxLength), includeLowerAlphas, includeUpperAlphas, includeNumerals, includeSymbols))
End Function

'Uses PRNG to create a random valid passphrase
Function GetRandomValidPassPhrase(ByVal minLength As Integer, ByVal maxLength As Integer, ByVal includeLowerAlphas As Boolean, ByVal includeUpperAlphas As Boolean, ByVal includeNumerals As Boolean, ByVal includeSymbols As Boolean)
    Dim added
    added = 0
    If includeLowerAlphas Then
        lower = RandomLowerAlpha()
        added = added + 1
    Else
        lower = ""
    End If
    If includeUpperAlphas Then
        upper = RandomUpperAlpha()
        added = added + 1
    Else
        upper = ""
    End If
    If includeNumerals Then
        numchar = RandomNumeral()
        added = added + 1
    Else
        numchar = ""
    End If
    If includeSymbols Then
        symchar = RandomSymbol()
        added = added + 1
    Else
        symchar = ""
    End If
    Dim PassPhrase
    PhraseLength = RandomInteger(minLength, maxLength)
    PassPhrase = lower & upper & numchar & symchar
    For Iterator = added - 1 To PhraseLength Step 1
        PassPhrase = PassPhrase & RandomCharacter(includeLowerAlphas, includeUpperAlphas, includeNumerals, includeSymbols)
    Next
    GetRandomValidPassPhrase = CStr(PassPhrase)
End Function


' Begin Tom H. Data Generation Functions


'Valid: no length limit, no special characters besides ` \ - _ ^ [ ]
'Used to generate first and last name, each its own function call
'Algo: Generate random number for length of string, then create array of length
'   each element in array gets randomly generated
Function genName(testStr, maxLen)
    Dim length
    Randomize
    length = RandomInteger(3, maxLen)
    ReDim charArr(length) As String
    
    'Generate positive data
    Dim output
    output = genLowerCaseStr(length)
    'Capitalize first letter
    output = UCase(Left(output, 1)) & mid(output, 2)
    genName = output
    
    'If set to negative, positive data will then be randomly corrupted
    If testStr <> "pos" Then
        Dim index, pos
        index = RandomInteger(1, 3)
        Select Case index
            'Case 1 - add random special character between ASCII 39-47 to random position
            Case 1:
                pos = RandomInteger(1, length)
                output = mid(output, 1, pos - 1) & chr(RandomInteger(33, 44)) & mid(output, pos, length)
            'Case 2 - add random number to random position
            Case 2:
                pos = RandomInteger(1, length)
                output = mid(output, 1, pos - 1) & chr(RandomInteger(48, 57)) & mid(output, pos, length)
            'Case 3 - add random special character between ASCII 58-64 to random position
            Case 3:
                pos = RandomInteger(1, length)
                output = mid(output, 1, pos - 1) & chr(RandomInteger(58, 64)) & mid(output, pos, length)
            
        End Select
    End If
    
    genName = output
End Function

'Returns a random string of lowercase characters with length "length"
Function genLowerCaseStr(length)
    ReDim charArr(length - 1) As String
    Dim i
    For i = 0 To length - 1
        'Number generated as [(max - min + 1)*Rnd] + 1
        charArr(i) = chr(Int(Rnd * 26) + 97)
    Next
    Dim output
    output = Join(charArr, "")
    genLowerCaseStr = output
End Function


'Valid: Of form xxx@xxx.com, can't include \ ( )
'Must have one and only one @
'May have unlimited .'s but none can be consecutive
Function genEmail(testStr, maxLen)
    Dim beforeAt, afterAt, beforeLen, afterLen, output
    beforeLen = Int(Rnd * maxLen) + 1
    afterLen = Int(Rnd * (maxLen / 2)) + 1
    beforeAt = genLowerCaseStr(beforeLen)
    afterAt = genLowerCaseStr(afterLen)
    output = beforeAt & "@" & afterAt & "." & genLowerCaseStr(3)
    
    'Negative case
    If testStr <> "pos" Then
        Dim index, pos, length, atPos
        length = beforeLen + afterLen + 1 'count of character before@, after@ and including @
        index = RandomInteger(1, 4)
        Select Case index
            'Case 1 - Add a second @ symbol
            Case 1:
                pos = RandomInteger(1, length)
                output = mid(output, 1, pos - 1) & "@" & mid(output, pos, length)
            'Case 2 - Add consecutive periods
            Case 2:
                pos = RandomInteger(1, length)
                output = mid(output, 1, pos - 1) & ".." & mid(output, pos, length)
            'Case 3 - Remove @ symbol
            Case 3:
                atPos = InStr(1, output, "@")
                output = mid(output, 1, atPos - 1) & mid(output, atPos + 1, length)
            'Case 4 - Remove .
            Case 4:
                atPos = InStr(1, output, ".")
                output = mid(output, 1, atPos - 1) & mid(output, atPos + 1, length)
        End Select
    End If
    
    genEmail = output
End Function

'Valid: 9 numbers, no other characters allowed, dashes don't need to be typed
Function genSocial(testStr)
    Dim firstSoc, midSoc, lastSoc, output
    
    firstSoc = CStr(RandomInteger(100, 999))
    midSoc = CStr(RandomInteger(10, 99))
    lastSoc = CStr(RandomInteger(1000, 9999))
    output = firstSoc & "-" & midSoc & "-" & lastSoc
    
    'Negative case only invalid input is less than full SS
    If testStr <> "pos" Then
        output = CStr(RandomInteger(1, 99999999))
    End If
    
    genSocial = output
End Function

'Valid: L###-####-#### or L########### where L is any capital letter
'Capital letter must be followed by 11 numbers
Function genDrivers(testStr)
    Dim letter, first, mid, last, output
    
    letter = chr(RandomInteger(65, 90))
    first = CStr(RandomInteger(100, 999))
    mid = CStr(RandomInteger(1000, 9999))
    last = CStr(RandomInteger(1000, 9999))
    
    output = letter & first & "-" & mid & "-" & last
    
    If testStr <> "pos" Then
        output = output & chr(RandomInteger(33, 47))
    End If
    
    genDrivers = output
End Function

Function genAddress(testStr)
    Dim output, streetNum, streetDir, streetDirIdx, streetName, streetNameLen, streetType, streetTypeIdx, streetTypeCollection
    streetNum = CStr(RandomInteger(1, 9999)) 'generate street number
    
    streetDirIdx = RandomInteger(1, 4) 'generate num between 1-4 to pick direction
    Select Case streetDirIdx
        Case 1:
            streetDir = "N"
        Case 2:
            streetDir = "E"
        Case 3:
            streetDir = "S"
        Case 4:
            streetDir = "W"
    End Select
    
    streetNameLen = RandomInteger(2, 15)
    streetName = genLowerCaseStr(streetNameLen)
    streetName = UCase(Left(streetName, 1)) & mid(streetName, 2) 'capitalize
    
    streetTypeCollection = Array("Ave", "Bch", "Blvd", "Cswy", "Ct", "Dr", "Gdns", "Holw", "Ln", "Mdw", "Orch", "Plz", "Rd", "Vis", "Way")
    streetTypeIdx = RandomInteger(1, 15)
    streetType = streetTypeCollection(streetTypeIdx - 1)
    
    output = streetNum & " " & streetDir & " " & streetName & " " & streetType
    
    If testStr <> "pos" Then
        Dim ranNum
        ranNum = RandomInteger(1, 2)
        Select Case ranNum
            Case 1: 'add special characters between ascii 33 and 45
                output = output & chr(RandomInteger(33, 45))
            Case 2: 'add special characters between ascii 58 and 64
                output = output & chr(RandomInteger(58, 64))
            Case 2: 'set address to random number
                output = CStr(RandomInteger(1, 10000))
        End Select
    End If
    
    genAddress = output
End Function

'Valid: TODO
Function genCity(testStr)
    Dim output, nameLen
    nameLen = RandomInteger(1, 25)
    output = genLowerCaseStr(nameLen)
    output = UCase(Left(output, 1)) & mid(output, 2)
    
    If testStr <> "pos" Then
        Dim ranNum
        ranNum = RandomInteger(1, 2)
        Select Case ranNum
            Case 1: 'add random number
                output = output & CStr(RandomInteger(1, 1000))
            Case 2: 'add special characters
                output = output & chr(RandomInteger(33, 47))
        End Select
    End If
    
    genCity = output
End Function
