VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "JSON Class."

'------------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2025 Vincent van Geerestein
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'------------------------------------------------------------------------------

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: JSON Module
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2026.11.09
'
' JSON implements a sophisticated json generator and parser capable of handling
' multi-dimensional arrays and serializing a VB Collection including its keys.
'
' The JSON Class is predeclared which creates a global default instance of the
' object with the same name making the object immediately accessible.
'
' Methods
' Stringify(var)        Serializes a VB data structure to a json string
' Parse(json)           Deserializes a json string to a VB data structure
' Prettify(json)        Prettifies a json string for visual purposes
' Simplify(json)        Simplifies a json string by removing whitespace
'
' These methods follow the json syntax (https://www.json.org/json-en.html)
' For original reference see http://www.ietf.org/rfc/rfc4627.txt
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Some different choices have been make compared to other VB json converters.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Here both a VB Collection and a VB Dictionary are converted to a json-object.
' A json-object with a unique set of keys is parsed to a VB Dictionary whereas a
' json-object with multiple empty keys but a further unique set of keys is
' parsed to a VB Collection.
'
' The For Each loop is used to serialize any multi-dimensional array. The loop
' follows the same column-major order in which the array elements are stored in
' memory. The parsing of a multi-dimensional json-array requires some intricate
' bookkeeping and memory manipulation.
'
' The parsing of a json-value as a number results in a VB Double, a VB Long,
' a VB LongLong or a VB Decimal depending on "the mimimal fit". A compiler
' directive can be set for handling Excel large numbers as strings.
'
' Source                Destination
' VB Dictionary         json-object (all keys are set and unique)
' VB Collection         json-object (set keys are unique)
' VB Array              json-array (any multi-dimensional array)
' VB Date               string with UTC (both date and time set) or VB format
' VB String             json-string
' VB Number             json-number
' VB Null               json-literal - null
' VB True               json-literal - true
' VB False              json-literal - false
' VB Empty              json-literal - null
' VB Nothing            will raise an error

' Source                Destination
' json-object           VB Dictionary (all keys need to be set and be unique)
' json-object           VB Collection (all set keys need to be unique)
' json-array            VB Array (any multi-dimensional array)
' json-string           VB String or VB Date (if recognized as such)
' json-number           VB Double, VB Long, VB LongLong (x64) or VB Decimal
' null                  VB Null
' false                 VB False
' true                  VB True
'
' Comparison of JsonConverter (https://github.com/VBA-tools/VBA-JSON) and JSON.
'
'                                   JsonConverter.ConvertToJson JSON.Stringify
' VB Array                          Stops after 2 dimensions    json-array
' VB Collection                     json-array                  json-object
' VB Collection Key(s)              Not accessible              String (encoded)
' VB Dictionary                     json-object                 json-object
' VB Dictionary Keys                String (not encoded)        String (encoded)
' Handles a LongLong type           no                          yes
'
'                                   JsonConverter.ParseJson     JSON.Parse
' json-array                        VB Collection               VB Array
' json-object (unique keys)         VB Dictionary               VB Dictionary
' json-object (empty keys)          Stops after 1st empty key   VB Collection
' "minimal type" number conversion  no                          yes
' large number parsed to            String                      Decimal
'
' JSON is only implemented for Windows. JsonConverter works also for Mac.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Compiler directives
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Use API or the faster Variant ByRef approach for getting Collection Keys.
#Const API = False

' Convert VB Collection Keys and include in json-object or ignore keys.
#Const jsonConvertCollectionKeys = True

' Convert a large number (> 15 digits) to a VB string and vice versa. Meant for
' Excel: Excel stores a number as a Double with 15 significant digits and has
' no LongLong, so every json-number with more than 15 digits becomes a string,
' also one that fits a LongLong. Thousands separators are removed from a
' large-number string when it is serialized.
#Const jsonConvertExcelLargeNumber = False

' Serialize the keys with quotes (default).
#Const jsonSerializeQuotedKeys = True

' Deserialize keys without quotes in addition to quoted keys.
#Const jsonDeserializeUnquotedKeys = False

' Convert VB Date to UTC String and vice versa, only if date and time are set.
#Const jsonConvertUTC = True

' Convert a json-string recognised as a date by VBA.IsDate to a VB Date.
' Disable to prevent locale-dependent silent conversion of strings like "12/01/2025".
#Const jsonConvertDate = False

' String deserialization behaviour depending on the two directives above:
'
' jsonConvertDate  jsonConvertUTC  Behaviour
' False            False           All json-strings are kept as a VB String.
' False            True            Only an exact UTC format (####-##-##T##:##:##Z)
'                                  is converted to a VB Date. All other json-strings
'                                  are kept as a VB String.
' True             False           Any json-string recognised by VBA.IsDate is
'                                  converted to a VB Date. All others as a VB String.
' True             True            Any json-string recognised by VBA.IsDate is
'                                  converted to a VB Date. If not, an exact UTC format
'                                  is converted to a VB Date. All others as a VB String.

' Stringify raises an error for VB data without a json form (default), or puts
' the error message in the json in place of the value and prints it to the
' Immediate window (the json is then invalid).
#Const jsonRaiseDataErrors = True

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private API declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/nf-wdm-rtlmovememory
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal NBytes As Long _
)

' https://docs.microsoft.com/en-us/windows/win32/api/timezoneapi/nf-timezoneapi-gettimezoneinformation
Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32.dll" ( _
    ByRef lpTimeZoneInformation As TimeZoneInformation _
) As Long

' https://learn.microsoft.com/en-us/windows/win32/api/timezoneapi/nf-timezoneapi-gettimezoneinformationforyear
Private Declare PtrSafe Function GetTimeZoneInformationForYear Lib "kernel32.dll" ( _
    ByVal wYear As Integer, _
    ByVal pdtzi As LongPtr, _
    ByRef lpTimeZoneInformation As TimeZoneInformation _
) As Long

' https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearraycreate
Private Declare PtrSafe Function SafeArrayCreate Lib "oleaut32.dll" ( _
    ByVal vt As Integer, _
    ByVal cDims As Integer, _
    ByVal rgsabound As LongPtr _
) As LongPtr

' VarPtrArray returns the address of an array variable (not for variant/array).
Private Declare PtrSafe Function VarPtrArray Lib "VBE7" Alias "VarPtr" ( _
    var() As Any _
) As LongPtr

' https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-sysreallocstring
Private Declare PtrSafe Function SysReAllocString Lib "oleaut32.dll" ( _
    ByVal pBSTR As LongPtr, _
    Optional ByVal pszStrPtr As LongPtr _
) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' The variable subtype constant for a LongPtr (not VB defined).
#If Win64 Then
Private Const vbLongPtr As Integer = VBA.vbLongLong
#Else
Private Const vbLongPtr As Integer = VBA.vbLong
#End If

' Selected memory sizes of intrinsic data types.
' https://docs.microsoft.com/en-us/office/vba/Language/Reference/User-Interface-Help/data-type-summary
Private Enum VARSIZE
    vbSizeInteger = 2
#If Win64 Then
    vbSizeLongPtr = 8
#Else
    vbSizeLongPtr = 4
#End If
End Enum

' Selected memory offsets for a VB Collection and a VB Array.
Private Enum MEMOFFSET
#If Win64 Then
    vbOffsetCollectionHead = 40         ' *head (LongPtr)
    vbOffsetCollectionTail = 48         ' *tail (LongPtr)
    vbOffsetCollectionKey = 24          ' *key  (LongPtr)
    vbOffsetCollectionPrev = 32         ' *prev (LongPtr)
    vbOffsetCollectionNext = 40         ' *next (LongPtr)
    vbOffsetArraypvData = 16
#Else
    vbOffsetCollectionHead = 24
    vbOffsetCollectionTail = 28
    vbOffsetCollectionKey = 16
    vbOffsetCollectionPrev = 20
    vbOffsetCollectionNext = 24
    vbOffsetArraypvData = 12
#End If
End Enum

' Selected VB errors.
Private Enum VBERROR
' https://onlinelibrary.wiley.com/doi/pdf/10.1002/9781118257616.app3
    vbErrorInvalidProcedureCall = 5
    vbErrorOutOfMemory = 7
    vbErrorTypeMismatch = 13
    vbErrorInternalError = 51
    vbErrorKeyAlreadyAssociated = 457
End Enum

' Json parsing errors.
Private Enum JSONERROR
    jsErrorInvalidData = VBA.vbObjectError + 1024
    jsErrorInvalidObject
    jsErrorInvalidType
    jsErrorInvalidJson
    jsErrorNameSeparatorExpected
    jsErrorValueSeparatorExpected
    jsErrorStringExpected
    jsErrorOpenString
    jsErrorOpenValue
    jsErrorOpenObject
    jsErrorOpenArray
    jsErrorDuplicateKey
    jsErrorInvalidValue
    jsErrorInvalidLiteral
    jsErrorInvalidNumber
    jsErrorInvalidCharacter
    jsErrorExtraCharacter
    jsErrorEmptyDimension
End Enum

' StringBuffer is used to prevent performance degradation due to a high volume string allocations.
Private Type StringBuilder
    Buffer As String
    Count As Long
    HeadSpace As Long
End Type

' https://docs.microsoft.com/en-us/windows/win32/api/minwinbase/ns-minwinbase-systemtime
Private Type SystemTime
    Year As Integer
    Month As Integer
    DayOfWeek As Integer
    Day As Integer
    Hour As Integer
    Minute As Integer
    Second As Integer
    MilliSeconds As Integer
End Type

' https://docs.microsoft.com/en-us/windows/win32/api/timezoneapi/ns-timezoneapi-time_zone_information
Private Type TimeZoneInformation
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As SystemTime
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As SystemTime
    DaylightBias As Long
End Type

' https://learn.microsoft.com/en-us/windows/win32/api/timezoneapi/nf-timezoneapi-gettimezoneinformation
Private Enum TIMEZONESTATUS
    TIME_ZONE_ID_INVALID = -1
    TIME_ZONE_ID_UNKNOWN = 0
    TIME_ZONE_ID_STANDARD = 1
    TIME_ZONE_ID_DAYLIGHT = 2
End Enum

' Variant ByRef construct for memory access by address.
#If API = False Then
Private Const VT_BYREF As Integer = &H4000
Private Type CONSTRUCT
    vt As Variant
    ref As Variant
End Type
Private VarByRef As CONSTRUCT
#End If

' The locale decimal point (Empty for not set, True for "." or False for ",").
Private DecimalPoint As Variant

' A Dictionary kept alive for the lifetime of the class: while no Dictionary
' exists, New Dictionary takes about 470 µs instead of 5 µs (measured), which
' would be paid for every json-object or matrix parsed.
Private KeepDictionaryAlive As Dictionary

' Lookup tables for Encode, see InitializeEncodeTables.
Private EncodeEscapes(0 To 127) As Byte
Private HexDigits(0 To 15) As Byte

' Lookup tables for Decode, see InitializeDecodeTables.
Private DecodeEscapes(0 To 127) As Byte
Private HexValues(0 To 127) As Byte


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()

    DecimalPoint = LocaleDecimalPoint
    Set KeepDictionaryAlive = New Dictionary
    InitializeEncodeTables
    InitializeDecodeTables

#If API = False Then
    InitializeVarByRef
#End If

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Stringify(ByVal var As Variant) As String
Attribute Stringify.VB_Description = "Returns the json for a VB data structure."
' Stringify is the front-end routine to start the serialization process.
' VB data without a json form raises an error, see DataError.

    Dim builder As StringBuilder
    If VBA.IsObject(var) Or VBA.IsArray(var) Then
        Serialize var, builder
    Else
        DataError builder, jsErrorInvalidData, VBA.TypeName(var)
    End If
    Stringify = Flush(builder)

End Function


Public Function Parse(ByVal JSON As String) As Variant
Attribute Parse.VB_Description = "Deserializes a json string to a VB data structure."
' Parse is the front-end for the deserialization process. The parser skips the
' whitespace between the tokens itself, so the positions in a parse error are
' those of the json as given. Any json-value is allowed at the top level.

    On Error GoTo ErrorHandler

    ' Initiate the deserialization process.
    Dim index As Long: index = SkipWhitespace(JSON, 1)
    Dim IsJsonObject As Boolean: IsJsonObject = (Mid$(JSON, index, 1) = "{")
    If IsJsonObject Then
        Set Parse = Deserialize(JSON, index)
    Else
        Parse = Deserialize(JSON, index)
    End If

    ' Only whitespace may follow the json-value.
    index = SkipWhitespace(JSON, index)
    If index <= Len(JSON) Then
        RaiseParseError jsErrorExtraCharacter, JSON, index
    End If
    Exit Function

ErrorHandler:
    ' Tidy up to prevent cascading errors: Nothing for a json-object, so that Set works.
    If IsJsonObject Then
        Set Parse = Nothing
    Else
        Parse = Empty
    End If

End Function


Public Function Prettify( _
    ByRef JSON As String, _
    Optional ByVal indentation As String = VBA.vbTab, _
    Optional ByVal space As String = " ", _
    Optional ByVal compact As Boolean = True _
) As String
Attribute Prettify.VB_Description = "Prettifies a json string for visualisation purposes."
' Compact = True outputs array values of the highest dimension on one line.

    On Error GoTo ErrorHandler

    Select Case Left$(JSON, 1) & Right$(JSON, 1)
    Case "[]", "{}"
        Dim builder As StringBuilder
        builder.HeadSpace = 2
        Append builder, vbNullString
    Case Else
        Prettify = ErrorDescription(jsErrorInvalidJson)
        Exit Function
    End Select

    Dim i As Long, Depth As Long, MaxDepth As Boolean
    For i = 1 To Len(JSON)
        Dim Char As String: Char = Mid$(JSON, i, 1)
        Select Case Char
        Case """"
            ' Copy the json-string as it is, escapes included; MatchString finds its end.
            Dim Start As Long: Start = i
            MatchString JSON, i
            Append builder, Mid$(JSON, Start, i - Start)
            i = i - 1
        Case "{"
            builder.Count = builder.Count + 1
            Mid$(builder.Buffer, builder.Count) = Char
            If MaxDepth = False Then
                Depth = Depth + 1
                Append builder, LineBreak(Depth, indentation)
            End If
        Case "}"
            If MaxDepth = False Then
                Depth = Depth - 1
                Append builder, LineBreak(Depth, indentation)
            End If
            builder.Count = builder.Count + 1
            Mid$(builder.Buffer, builder.Count) = Char
        Case "["
            If compact Then
                ' Determine whether this is the maximum array depth.
                Dim match As Long: match = InStr(i + 1, JSON, "[")
                If match = 0 Then
                    MaxDepth = True
                Else
                    MaxDepth = match > InStr(i + 1, JSON, "]")
                End If
            End If
            builder.Count = builder.Count + 1
            Mid$(builder.Buffer, builder.Count) = Char
            If MaxDepth = False Then
                Depth = Depth + 1
                Append builder, LineBreak(Depth, indentation)
            End If
        Case "]"
            If MaxDepth Then
                MaxDepth = False
            Else
                Depth = Depth - 1
                Append builder, LineBreak(Depth, indentation)
            End If
            builder.Count = builder.Count + 1
            Mid$(builder.Buffer, builder.Count) = Char
        Case ","
            builder.Count = builder.Count + 1
            Mid$(builder.Buffer, builder.Count) = Char
            If MaxDepth Then
                Append builder, space
            Else
                Append builder, LineBreak(Depth, indentation)
            End If
        Case ":"
            Append builder, Char & space
        Case Else
            ' json-numbers and json-literals end up here.
            Append builder, Char
        End Select
    Next
    Prettify = Flush(builder)
    Exit Function

ErrorHandler:
    ' Parsing the json input will pinpoint the error.
    Debug.Print Me.Parse(JSON)

End Function


Private Function LineBreak(ByVal level As Long, ByRef indentation As String) As String
Attribute LineBreak.VB_Description = "Returns a line break and the indentation for a level."
' The indentation may be any string: it is repeated level times.

    LineBreak = VBA.vbCrLf & VBA.Replace(space$(level), " ", indentation)

End Function


Public Function Simplify(ByVal JSON As String) As String
Attribute Simplify.VB_Description = "Simplifies a json string by removing whitespace."
' ByVal is used explicly to create a local copy of json.

    ' WhiteSpace match pattern used with the Like operator.
    Const WhiteSpace As String = "[" & " " & VBA.vbTab & VBA.vbLf & VBA.vbCr & "]"

    Dim Last As Long
    Dim Start As Long: Start = 1
    Dim i As Long, CharWithinString As Boolean
    For i = 1 To Len(JSON)
        Dim Char As String: Char = Mid$(JSON, i, 1)
        If Char = """" Then
            CharWithinString = Not CharWithinString
        ElseIf CharWithinString Then
            ' Skip any escaped character in a string in order to skip \".
            If Char = "\" Then
                i = i + 1
            End If
        ElseIf Char Like WhiteSpace Then
            ' Append the buffered json sub-string.
            Mid$(JSON, Last + 1) = Mid$(JSON, Start, i - Start)
            Last = Last + i - Start
            Start = i + 1
        End If
    Next

    If Last = 0 Then
        ' No whitespace has been detected.
        Simplify = JSON
    Else
        ' Append the remaining string which has no whitespace.
        Mid$(JSON, Last + 1) = Mid$(JSON, Start)
        Last = Last + Len(JSON) - Start + 1
        Simplify = Mid$(JSON, 1, Last)
    End If

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Serialize(ByRef var As Variant, ByRef builder As StringBuilder)
Attribute Serialize.VB_Description = "Appends the json for a VB Data structure to a string buffer."
' This main decision tree is called recursively in the serialization process.
' A string or a date needs no headspace of its own: Append always leaves 2
' characters free, for the closing """" and the next structural character.

    ' At this point there is always at least 1 headspace left!
    Select Case VBA.VarType(var)
    Case VBA.vbString
        builder.Count = builder.Count + 1
        Mid(builder.Buffer, builder.Count) = """"
#If jsonConvertExcelLargeNumber Then
        If IsExcelLargeNumberString(var) Then
            ' Convert an Excel Large Number String to a json-string.
            If DecimalPoint Then
                ' Remove any thousand separators.
                Append builder, VBA.Replace(var, ",", vbNullString)
            Else
                ' Remove any thousand separators and update a digital separator.
                Append builder, VBA.Replace(VBA.Replace(var, ".", vbNullString), ",", ".", 1, 1)
            End If
        Else
            ' Convert a VB String to a json-string.
            Append builder, Encode(CStr(var))
        End If
#Else
        ' Convert a VB String to a json-string.
        Append builder, Encode(CStr(var))
#End If
        builder.Count = builder.Count + 1
        Mid(builder.Buffer, builder.Count) = """"
    Case VBA.vbInteger, VBA.vbLong, VBA.vbByte, vbLongPtr
        ' Convert a VB number to a json-number.
        Append builder, CStr(var)
    Case VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
        ' Convert a VB number to a json-number.
        If DecimalPoint Then
            Append builder, CStr(var)
        Else
            Append builder, VBA.Replace(var, ",", ".", 1, 1)
        End If
    Case VBA.vbObject
        ' TypeOf rather than TypeName: TypeName of a Scripting.Dictionary takes
        ' about 500 µs per call, TypeOf takes no measurable time.
        If TypeOf var Is Dictionary Then
            ' Serialize a VB Dictionary to a json-object.
            If var.Count = 0 Then
                Append builder, "{}"
            Else
                SerializeDictionary var, builder
            End If
        ElseIf TypeOf var Is Collection Then
            ' Serialize a VB Collection to a json-object.
            If var.Count = 0 Then
                Append builder, "{}"
            Else
                SerializeCollection var, builder
            End If
        Else
            ' Other VB objects can't be serialized.
            DataError builder, jsErrorInvalidObject, VBA.TypeName(var)
        End If
    Case Is >= VBA.vbArray
        ' Serialize the VB Array to a json-array.
        Dim NDims As Long: NDims = ArrDims(var)
        Select Case NDims
        Case 1
            SerializeArray var, builder
        Case 0
            Append builder, "[]"
        Case Else
            SerializeMatrix var, builder, NDims
        End Select
    Case VBA.vbBoolean
        If var Then
            Append builder, "true"
        Else
            Append builder, "false"
        End If
    Case VBA.vbNull, VBA.vbEmpty
        Append builder, "null"
    Case VBA.vbDate
        builder.Count = builder.Count + 1
        Mid(builder.Buffer, builder.Count) = """"
#If jsonConvertUTC Then
        If var <> VBA.CLng(var) And var >= 1 Then
            ' Convert a VB Date to a UTC format if both the date and the time are set.
            Append builder, ConvertLocalTimeToUTC(var)
        Else
            ' The variant/date is implicitly converted to a string.
            Append builder, CStr(var)
        End If
#Else
        ' The variant/date is implicitly converted to a string.
        Append builder, CStr(var)
#End If
        builder.Count = builder.Count + 1
        Mid(builder.Buffer, builder.Count) = """"
    Case Else
        ' Other VB types can't be serialized: vbUserDefinedType, vbError and vbDataObject.
        DataError builder, jsErrorInvalidType, VBA.TypeName(var)
    End Select

End Sub


Private Sub SerializeDictionary( _
    ByRef var As Variant, _
    ByRef builder As StringBuilder _
)
Attribute SerializeDictionary.VB_Description = "Appends the json for a VB Dictionary to a string buffer."
' A Dictionary key is usually an integer or a string, but it can be anything
' except an array. The keys are unique and one key could be an empty string.
' Here the keys are converted to strings, which will raise an error if this
' is not possible, e.g. when the key is an object.

    ' Headspace for adding structural characters.
    Const HeadSpace As Long = 2
    Adjust builder, HeadSpace

    Dim Keys() As Variant: Keys = var.Keys

    ' Append the json-object to the string buffer.
    builder.Count = builder.Count + 1
    Mid$(builder.Buffer, builder.Count) = "{"
    Dim Item As Variant, i As Long
    For Each Item In var.Items
#If jsonSerializeQuotedKeys Then
        builder.Count = builder.Count + 1
        Mid(builder.Buffer, builder.Count) = """"
        Append builder, Encode(CStr(Keys(i)))
        Mid(builder.Buffer, builder.Count + 1) = """:"
        builder.Count = builder.Count + 2
#Else
        AppendUnquotedKey builder, CStr(Keys(i))
        builder.Count = builder.Count + 1
        Mid$(builder.Buffer, builder.Count) = ":"
#End If
        Serialize Item, builder
        builder.Count = builder.Count + 1
        Mid$(builder.Buffer, builder.Count) = ","
        i = i + 1
    Next
    Mid$(builder.Buffer, builder.Count) = "}"

    ' Release the reserved headspace.
    Adjust builder, -HeadSpace

End Sub


Private Sub SerializeCollection( _
    ByRef var As Variant, _
    ByRef builder As StringBuilder _
)
Attribute SerializeCollection.VB_Description = "Appends the json for a VB Collection to a string buffer."
' A Collection key is always a string. For a Collection not all keys need to be
' set. For json any unset key is converted to an empty string.

    ' Headspace for adding structural characters, 2 extra for an empty key.
    Const HeadSpace As Long = 4
    Adjust builder, HeadSpace

#If jsonConvertCollectionKeys Then
#If API Then
    Dim Keys() As String: Keys = CollKeys(var)
#Else
    Dim Keys() As String: Keys = CollKeys(var, VarByRef.vt)
#End If
#End If

    ' Append the json-object to the string buffer.
    builder.Count = builder.Count + 1
    Mid$(builder.Buffer, builder.Count) = "{"
    Dim Item As Variant, i As Long
    For Each Item In var
#If jsonSerializeQuotedKeys Then
        builder.Count = builder.Count + 1
        Mid(builder.Buffer, builder.Count) = """"
#End If
#If jsonSerializeQuotedKeys Then
#If jsonConvertCollectionKeys Then
        If Len(Keys(i)) <> 0 Then
            Append builder, Encode(Keys(i))
        End If
#End If
#ElseIf jsonConvertCollectionKeys Then
        AppendUnquotedKey builder, Keys(i)
#Else
        AppendUnquotedKey builder, vbNullString
#End If
#If jsonSerializeQuotedKeys Then
        Mid(builder.Buffer, builder.Count + 1) = """:"
        builder.Count = builder.Count + 2
#Else
        builder.Count = builder.Count + 1
        Mid(builder.Buffer, builder.Count) = ":"
#End If
        Serialize Item, builder
        builder.Count = builder.Count + 1
        Mid$(builder.Buffer, builder.Count) = ","
        i = i + 1
    Next
    Mid$(builder.Buffer, builder.Count) = "}"

    ' Release the reserved headspace.
    Adjust builder, -HeadSpace

End Sub


Private Sub SerializeArray( _
    ByRef var As Variant, _
    ByRef builder As StringBuilder _
)
Attribute SerializeArray.VB_Description = "Appends the json for a VB Array to a string buffer."

    ' Headspace for adding structural characters.
    Const HeadSpace As Long = 2
    Adjust builder, HeadSpace

    ' Append the json-array to the string buffer.
    builder.Count = builder.Count + 1
    Mid$(builder.Buffer, builder.Count) = "["
    Dim i As Long
    For i = LBound(var) To UBound(var)
        Serialize var(i), builder
        builder.Count = builder.Count + 1
        Mid$(builder.Buffer, builder.Count) = ","
    Next
    Mid$(builder.Buffer, builder.Count) = "]"

    ' Release the reserved headspace.
    Adjust builder, -HeadSpace

End Sub


Private Sub SerializeMatrix( _
    ByRef var As Variant, _
    ByRef builder As StringBuilder, _
    ByVal NDims As Long _
)
Attribute SerializeMatrix.VB_Description = "Appends the json for a multi-dimensional VB Array to string buffer."
' The For Each loop enumerates the elements of any multi-dimensional array in a
' column-major order which corresponds to the way in which the array elements
' are physically stored in memory. Although the enumeration of the elements is
' thus straightforward, the generation of the appropriate structural characters
' requires some thinkering. In the serialization process a string of structural
' characters is added after each element and its specific form depends on the
' number of changing dimensions moving from the current to the next element. For
' no change of dimension the structural character is obviously just ",". For 1
' change of dimension the string of structural characters is "],[", and for 2
' changes "]],[[", etc. This number of changing dimensions can be obtained by
' comparing the moduli of the sequence index with the strides of the array.
' https://en.wikipedia.org/wiki/Row-_and_column-major_order for strides.

    ' An array with an empty dimension has no json form: the shape would be lost.
    ' ReDim can't make such an array, but the Windows API or a COM component can.
    Dim n As Long
    For n = 1 To NDims
        If UBound(var, n) < LBound(var, n) Then
            DataError builder, jsErrorEmptyDimension, "dimension " & n
            Exit Sub
        End If
    Next

    ' Headspace for adding structural characters.
    Dim HeadSpace As Long: HeadSpace = 2 * NDims
    Adjust builder, HeadSpace

    ' The delimiters depend on the number of array dimensions.
    Static Delimiters(1 To 32) As String ' Max dimensions in VB = 32.
    Static Initialized As Long
    If NDims > Initialized Then
        If Initialized = 0 Then
            Delimiters(1) = ","
            Initialized = 1
        End If
        For n = Initialized + 1 To NDims
            Delimiters(n) = "]" & Delimiters(n - 1) & "["
        Next
        Initialized = NDims
    End If

    ' The strides depend on the actual array dimensions.
    Dim Strides() As Long: ReDim Strides(0 To NDims)
    Strides(0) = 1
    For n = 1 To NDims
        Strides(n) = Strides(n - 1) * (UBound(var, n) - LBound(var, n) + 1)
    Next

    ' Append the json-array to the string buffer.
    Mid$(builder.Buffer, builder.Count + 1) = String$(NDims, "[")
    builder.Count = builder.Count + NDims
    Dim Item As Variant, i As Long
    For Each Item In var
        i = i + 1
        For n = 1 To NDims - 1
            If (i Mod Strides(n)) <> 0 Then
                Exit For
            End If
        Next
        Serialize Item, builder
        Mid$(builder.Buffer, builder.Count + 1) = Delimiters(n)
        builder.Count = builder.Count + Len(Delimiters(n))
    Next
    builder.Count = builder.Count - NDims + 1
    Mid$(builder.Buffer, builder.Count) = "]"

    ' Release the reserved headspace.
    Adjust builder, -HeadSpace

End Sub


#If jsonDeserializeUnquotedKeys Or Not jsonSerializeQuotedKeys Then
Private Function IsSimpleKey(ByRef key As String) As Boolean
Attribute IsSimpleKey.VB_Description = "Determines whether a key is a simple name that needs no quotes."
' A simple name starts with a letter, "_" or "$", followed by letters, digits,
' "_" or "$", as in JavaScript and JSON5. Like compares binary in this class.

    IsSimpleKey = key Like "[A-Za-z_$]*" And Not key Like "?*[!A-Za-z0-9_$]*"

End Function
#End If


#If Not jsonSerializeQuotedKeys Then
Private Sub AppendUnquotedKey(ByRef builder As StringBuilder, ByRef key As String)
Attribute AppendUnquotedKey.VB_Description = "Appends a key without quotes, or with quotes when it is not a simple name."
' An empty key or a key with other characters is quoted anyway, so that the
' json can be read back.

    If IsSimpleKey(key) Then
        Append builder, key
    Else
        Append builder, """" & Encode(key) & """"
    End If

End Sub
#End If


Private Sub InitializeEncodeTables()
Attribute InitializeEncodeTables.VB_Description = "Fills the lookup tables for Encode."

    ' The character after "\" for the characters with a short escape.
    EncodeEscapes(8) = 98       ' VBA.vbBack -> "\b"
    EncodeEscapes(9) = 116      ' VBA.vbTab -> "\t"
    EncodeEscapes(10) = 110     ' VBA.vbLf -> "\n"
    EncodeEscapes(12) = 102     ' VBA.vbFormFeed -> "\f"
    EncodeEscapes(13) = 114     ' VBA.vbCr -> "\r"
    EncodeEscapes(34) = 34      ' """"  -> "\"""
' RFC 8259 allows "\/" but does not require escaping "/", see Encode.
'   EncodeEscapes(47) = 47      ' "/" -> "\/"
    EncodeEscapes(92) = 92      ' "\" -> "\\"

    ' The character codes of the hexadecimal digits 0-9 and A-F.
    Dim i As Long
    For i = 0 To 15
        HexDigits(i) = VBA.Asc(Mid$("0123456789ABCDEF", i + 1, 1))
    Next

End Sub


Private Function Encode(ByRef str As String) As String
Attribute Encode.VB_Description = "Encodes a Unicode/Ascii string to a printable Ascii string."
' Encodes a Unicode/Ascii string to a printable Ascii string.
'
' Most strings need no encoding at all. The Like test below detects that in a
' fraction of a microsecond and the string is returned as it is.
'
' The class has no Option Compare statement, so Like compares binary and the
' ranges in the pattern are character codes: printable Ascii except """" (34)
' and "\" (92). "]" (93) can't be in a character list; it takes the slow path,
' which leaves it as it is. "/" is not escaped: RFC 8259 allows "\/" but does
' not require it, so "/" is in the pattern (in the range #-[) and EncodeEscapes
' (47) is not set. To write "\/", for example for json inside an HTML <script>
' element, set EncodeEscapes(47) = 47 and also take "/" out of the pattern:
' "*[! -!#-.0-[^-~]*". A "\" is always written as "\\", which is required.
'
' The slow path reads the characters from a byte array and writes the result
' to a byte array, so no strings are created until the end. A VB string has 2
' bytes a character; the result is printable Ascii, so only the low byte of a
' character is written and the high byte stays 0.

    ' Any character that might need encoding.
    Const MightNeedEncoding As String = "*[! -!#-[^-~]*"
    If Not str Like MightNeedEncoding Then
        Encode = str
        Exit Function
    End If

    ' VB allows a quick copy of a string into a byte array.
    Dim Chars() As Byte: Chars = str
    ' The maximum size when all characters would be \u encoded.
    Dim Encoded() As Byte: ReDim Encoded(0 To 12 * Len(str) - 1)
    Dim index As Long

    Dim i As Long, Code As Long, Escaped As Long
    For i = 0 To UBound(Chars) Step 2
        ' The character code as an unsigned number (0 to 65535).
        Code = Chars(i) Or Chars(i + 1) * &H100&
        If Code < 127 Then
            Escaped = EncodeEscapes(Code)
        Else
            Escaped = 0
        End If
        If Escaped <> 0 Then
            ' Special Ascii characters.
            Encoded(index) = 92                 ' "\"
            Encoded(index + 2) = Escaped
            index = index + 4
        ElseIf Code < 32 Or Code >= 127 Then
            ' Non-printable Ascii, extended Ascii and unicode characters.
            Encoded(index) = 92                 ' "\"
            Encoded(index + 2) = 117            ' "u"
            Encoded(index + 4) = HexDigits(Code \ &H1000&)
            Encoded(index + 6) = HexDigits((Code \ &H100&) And &HF&)
            Encoded(index + 8) = HexDigits((Code \ &H10&) And &HF&)
            Encoded(index + 10) = HexDigits(Code And &HF&)
            index = index + 12
        Else
            ' Printable Ascii characters.
            Encoded(index) = Code
            index = index + 2
        End If
    Next

    ReDim Preserve Encoded(0 To index - 1)
    Encode = Encoded

End Function


Private Function Deserialize( _
    ByRef JSON As String, _
    ByRef index As Long _
) As Variant
Attribute Deserialize.VB_Description = "Deserializes a json string to a VB data stucture."
' This main decision tree is called recursively in the deserialization process.
'
' value is a Variant, not a String: a filled String local costs about 0.65 µs
' extra when the procedure exits, a Variant hardly anything. Deserialize is
' called for every json-value.

    On Error GoTo ErrorHandler

    Dim value As Variant
    If Mid$(JSON, index, 1) <= " " Then index = SkipWhitespace(JSON, index)
    Select Case Mid$(JSON, index, 1)
    Case """"
        ' MatchString advances the index to the next string position.
        value = MatchString(JSON, index)
#If jsonConvertDate Then
        If VBA.IsDate(value) Then
            ' Convert json-string containing a date to a VB Date.
            Deserialize = VBA.CDate(value)
#If jsonConvertUTC Then
        ElseIf value Like "####-##-##T##:##:##Z" Then
            ' Convert json-string containing a UTC format to a VB Date.
            Deserialize = ConvertUTCToLocalTime(value)
#End If
        Else
            ' Keep json-string as VB String.
            Deserialize = value
        End If
#Else
#If jsonConvertUTC Then
        If value Like "####-##-##T##:##:##Z" Then
            ' Convert json-string containing a UTC format to a VB Date.
            Deserialize = ConvertUTCToLocalTime(value)
        Else
            ' Keep json-string as VB String.
            Deserialize = value
        End If
#Else
        ' Keep json-string as VB String.
        Deserialize = value
#End If
#End If
    Case "{"
        Dim Ahead As Long: Ahead = SkipWhitespace(JSON, index + 1)
        If Mid$(JSON, Ahead, 1) = "}" Then
            ' Deserialize empty json-object to an empty VB Dictionary.
            Set Deserialize = New Dictionary
            index = Ahead + 1
        Else
            ' Deserialize json-object to VB Dictionary or VB Collection.
            Set Deserialize = DeserializeDictionary(JSON, index)
        End If
    Case "["
        ' Count the number of consecutive '[' characters at the Index position.
        Dim NDim As Long
        Ahead = index
        Do While Mid$(JSON, Ahead, 1) = "["
            NDim = NDim + 1
            Ahead = SkipWhitespace(JSON, Ahead + 1)
        Loop
        ' If the inner dimension is an empty array it counts as an element.
        If Mid$(JSON, Ahead, 1) = "]" Then
            NDim = NDim - 1
        End If
        Select Case NDim
        Case 0
            ' Deserialize empty json-array to an empty VB Variant() array.
            Deserialize = VBA.Array()
            index = Ahead + 1
        Case 1
            ' Deserialize json-array to a VB Variant() array.
            Deserialize = DeserializeArray(JSON, index)
        Case Else
            ' Deserialize n-dimensional json-array to a VB Variant() array.
            Deserialize = DeserializeMatrix(JSON, index, NDim)
        End Select
    Case "-", "0" To "9"
        ' MatchValue advances the index to the next string position.
        value = MatchValue(JSON, index)
        If VBA.IsNumeric(value) = False Then
            ' Value is not recognized as a valid json-number.
            index = index - Len(value) ' Roll-back.
            RaiseParseError jsErrorInvalidNumber, JSON, index, value
        ElseIf InStr(value, "e") <> 0 Or InStr(value, "E") <> 0 Then
            ' Convert json-number with exponent to a VB Double.
            Deserialize = VBA.Val(value)
#If jsonConvertExcelLargeNumber Then
        ElseIf IsExcelLargeNumberString(value) Then
           ' Keep the json-number as a VB String.
            If DecimalPoint Then
                Deserialize = value
            Else
                Deserialize = VBA.Replace(value, ".", ",", 1, 1)
            End If
#End If
        ElseIf InStr(value, ".") = 0 Then
            On Error Resume Next
            ' Convert json-number to a VB Long.
            Deserialize = VBA.CLng(value)
            If Err.Number = 0 Then Exit Function
#If Win64 Then
            ' Convert json-number to a VB LongLong; clear the overflow of CLng first.
            Err.Clear
            Deserialize = VBA.CLngLng(value)
            If Err.Number = 0 Then Exit Function
#End If
            On Error GoTo ErrorHandler
            ' Convert json-number to a VB Decimal.
            Deserialize = VBA.CDec(value)
        ElseIf Len(value) = Len(VBA.CStr(Val(value))) Then
            ' Convert json-number to VB Double.
            Deserialize = VBA.Val(value)
        ElseIf DecimalPoint Then
            ' Convert json-number to a VB Decimal.
            Deserialize = VBA.CDec(value)
        Else
            ' Convert json-number to a VB Decimal.
            Deserialize = VBA.CDec(VBA.Replace(value, ".", ",", 1, 1))
        End If
    Case Else
        ' MatchValue advances the index to the next string position.
        value = MatchValue(JSON, index)
        ' Convert json-literal to its VB equivalent.
        Select Case value
        Case "true"
            Deserialize = True
        Case "false"
            Deserialize = False
        Case "null"
            Deserialize = Null
        Case Else
            ' Value is not recognized as a valid json-literal.
            index = index - Len(value) ' Roll-back.
            RaiseParseError jsErrorInvalidLiteral, JSON, index, value
        End Select
    End Select
    Exit Function

ErrorHandler:
    If Err.Number = vbErrorKeyAlreadyAssociated Then
        ' Multiple empty keys have been detected in parsing the json-object.
        Set Deserialize = DeserializeCollection(JSON, index)
    Else
        Err.Raise Err.Number
    End If

End Function


Private Function DeserializeDictionary( _
    ByRef JSON As String, _
    ByRef index As Long _
) As Variant
Attribute DeserializeDictionary.VB_Description = "Deserializes a json-object to a VB Dictionary."
' Json doesn't require a unique set of keys and allows multiple keys to be empty
' strings. A VB Dictionary doesn't allow duplicate keys, but one can be an empty
' string. If duplicate empty keys are detected, an attempt will be made to parse
' the json-object to a Collection which does allow for multiple empty keys.

    On Error GoTo ErrorHandler

    Dim Items As Dictionary: Set Items = New Dictionary

    Dim i As Long, key As String
    For i = index + 1 To Len(JSON)
        ' Expect key.
        If Mid$(JSON, i, 1) <= " " Then i = SkipWhitespace(JSON, i)
        If Mid$(JSON, i, 1) = """" Then
            key = MatchString(JSON, i)
        Else
#If jsonDeserializeUnquotedKeys Then
            key = MatchValue(JSON, i)
            If Not IsSimpleKey(key) Then RaiseParseError jsErrorStringExpected, JSON, i - Len(key), key
#Else
            RaiseParseError jsErrorStringExpected, JSON, i
#End If
        End If

        ' Expect name separator ":".
        If Mid$(JSON, i, 1) <= " " Then i = SkipWhitespace(JSON, i)
        If Mid$(JSON, i, 1) = ":" Then
            i = i + 1
        Else
            RaiseParseError jsErrorNameSeparatorExpected, JSON, i
        End If

        ' Add key : (parsed) value pair to the Dictionary.
        Items.Add key, Deserialize(JSON, i)

        ' Expect a value separator "," or object-end "}".
        If Mid$(JSON, i, 1) <= " " Then i = SkipWhitespace(JSON, i)
        Select Case Mid$(JSON, i, 1)
        Case "}"
            Exit For
        Case vbNullString
            RaiseParseError jsErrorOpenObject, JSON, i
        Case Is <> ","
            RaiseParseError jsErrorValueSeparatorExpected, JSON, i
        End Select
    Next
    ' The loop also ends after a "," at the end of the json.
    If Mid$(JSON, i, 1) <> "}" Then RaiseParseError jsErrorOpenObject, JSON, i
    Set DeserializeDictionary = Items
    index = i + 1
    Exit Function

ErrorHandler:
    If Err.Number = vbErrorKeyAlreadyAssociated And key <> VBA.vbNullString Then
        RaiseParseError jsErrorDuplicateKey, JSON, i - Len(key), key
    Else
        Err.Raise Err.Number
    End If

End Function


Private Function DeserializeCollection( _
    ByRef JSON As String, _
    ByRef index As Long _
) As Variant
Attribute DeserializeCollection.VB_Description = "Deserializes a json-object to a VB Collection."
' Json doesn't require a unique set of keys and allows multiple keys to be empty
' strings. A VB Collection doesn't allow duplicate keys, but keys can be omitted
' as left as empty strings. If duplicate keys are detected the parsing fails.

    On Error GoTo ErrorHandler

    Dim Items As Collection: Set Items = New Collection

    Dim i As Long, key As String
    For i = index + 1 To Len(JSON)
        ' Expect key.
        If Mid$(JSON, i, 1) <= " " Then i = SkipWhitespace(JSON, i)
        If Mid$(JSON, i, 1) = """" Then
            key = MatchString(JSON, i)
        Else
#If jsonDeserializeUnquotedKeys Then
            key = MatchValue(JSON, i)
            If Not IsSimpleKey(key) Then RaiseParseError jsErrorStringExpected, JSON, i - Len(key), key
#Else
            RaiseParseError jsErrorStringExpected, JSON, i
#End If
        End If

        ' Expect name separator ":".
        If Mid$(JSON, i, 1) <= " " Then i = SkipWhitespace(JSON, i)
        If Mid$(JSON, i, 1) = ":" Then
            i = i + 1
        Else
            RaiseParseError jsErrorNameSeparatorExpected, JSON, i
        End If

        ' Add key : value pair to VB the Collection.
        If Len(key) = 0 Then
            Items.Add Deserialize(JSON, i)
        Else
            Items.Add Deserialize(JSON, i), key
        End If

        ' Expect value separator "," or object-end "}".
        If Mid$(JSON, i, 1) <= " " Then i = SkipWhitespace(JSON, i)
        Select Case Mid$(JSON, i, 1)
        Case "}"
            Exit For
        Case vbNullString
            RaiseParseError jsErrorOpenObject, JSON, i + 1
        Case Is <> ","
            RaiseParseError jsErrorValueSeparatorExpected, JSON, i
        End Select
    Next
    ' The loop also ends after a "," at the end of the json.
    If Mid$(JSON, i, 1) <> "}" Then RaiseParseError jsErrorOpenObject, JSON, i
    Set DeserializeCollection = Items
    index = i + 1
    Exit Function

ErrorHandler:
    If Err.Number = vbErrorKeyAlreadyAssociated Then
        RaiseParseError jsErrorDuplicateKey, JSON, i - Len(key), key
    Else
        Err.Raise Err.Number
    End If

End Function


Private Function DeserializeArray( _
    ByRef JSON As String, _
    ByRef index As Long _
) As Variant
Attribute DeserializeArray.VB_Description = "Deserializes a json-array to a VB Array."

    On Error GoTo ErrorHandler

    ' Items is a buffer array to store the parsed array elements.
    Const InitialLength As Long = 100, GrowthFactor As Double = 2
    Dim Items() As Variant: ReDim Items(0 To InitialLength - 1)

    Dim i As Long, Last As Long
    For i = index + 1 To Len(JSON)
        If Last > UBound(Items) Then
            ' Increase the length of the array whilst preserving its content.
            ReDim Preserve Items(0 To GrowthFactor * UBound(Items) + 1)
        End If

        ' Add the parsed value to the next array element.
        If Mid$(JSON, i, 1) <= " " Then i = SkipWhitespace(JSON, i)
        Select Case Mid$(JSON, i, 1)
        Case "{"
            Set Items(Last) = Deserialize(JSON, i)
        Case Else
            Items(Last) = Deserialize(JSON, i)
        End Select

        ' Expect a value separator "," or array-end "]".
        If Mid$(JSON, i, 1) <= " " Then i = SkipWhitespace(JSON, i)
        Select Case Mid$(JSON, i, 1)
        Case "]"
            Exit For
        Case vbNullString
            RaiseParseError jsErrorOpenArray, JSON, i + 1
        Case Is <> ","
            RaiseParseError jsErrorValueSeparatorExpected, JSON, i
        End Select
        Last = Last + 1
    Next
    ' The loop also ends after a "," at the end of the json, or for "[" alone.
    If Mid$(JSON, i, 1) <> "]" Then RaiseParseError jsErrorOpenArray, JSON, i

    ' Resize array buffer to actual size and return.
    ReDim Preserve Items(0 To Last)
    DeserializeArray = Items
    index = i + 1
    Exit Function

ErrorHandler:
    Err.Raise Err.Number

End Function


Private Function DeserializeMatrix( _
    ByRef JSON As String, _
    ByRef index As Long, _
    ByVal NDim As Long _
) As Variant
Attribute DeserializeMatrix.VB_Description = "Deserializes a n-dimensional json-array to a VB Array."
' The parsing of a n-dimensional json-array requires quite a bit of work. The
' array dimensions need to be established whilst the json-array may also contain
' syntax errors or perhaps a particular dimension is in fact a jagged array. The
' process requires book keeping to match the structural characters and to record
' the lengths of all sub-arrays in order to detect a jagged array situation.
'
' The array elements are initially stored in a 1-dimensional array. At the
' end of the procedure the data of this one dimensional array is converted to
' the correct n-dimensional array.

    On Error GoTo ErrorHandler

    ' Items is a buffer array to store the parsed array elements.
    Const InitialLength As Long = 1000, GrowthFactor As Double = 2
    Dim Items() As Variant: ReDim Items(0 To InitialLength - 1)

    ' A stack is used to match the "[" and "]" structural characters.
    Dim Stack() As Long

    ' Lengths stores all arrays lengths at a certain depth. Typed As Dictionary:
    ' late-bound calls on a Dictionary in a Variant cost about 1 ms per matrix.
    Dim Lengths() As Dictionary

    Dim Last As Long, Depth As Long, Mixed As Boolean, Again As Boolean
    Do
        Last = 0 ' Important to set to 0 for jagged arrays!
        Depth = 0
        Mixed = False

        ReDim Stack(1 To NDim)

        ' Initialize the bookkeeping of the sub-array lengths.
        ReDim Lengths(1 To NDim)
        Dim n As Long
        For n = 1 To NDim
            Set Lengths(n) = New Dictionary
        Next

        ' The recursion depth is limited by NDim for handling jagged arrays.
        Dim i As Long
        For i = index To Len(JSON)
            ' Parse any "[" structural characters.
            If Mid$(JSON, i, 1) <= " " Then i = SkipWhitespace(JSON, i)
            Do While Mid$(JSON, i, 1) = "[" And Depth < NDim
                If Mid$(JSON, SkipWhitespace(JSON, i + 1), 1) = "]" Then
                    Exit Do
                End If
                Depth = Depth + 1
                Stack(Depth) = i
                ' Start counting the elements in this sub-array.
                Lengths(Depth).Add i, 1
                i = SkipWhitespace(JSON, i + 1)
            Loop

            ' A value (or an empty json-array) above the deepest level, as in [[1,2],3]:
            ' the json-array is an array of arrays and values. Parse it again with the
            ' dimensions down to this level; Deserialize parses the deeper arrays.
            If Depth < NDim Then
                NDim = Depth
                Mixed = True
                Exit For
            End If

            ' Check the array length and extend it if necessary.
            If Last > UBound(Items) Then
                ' Double the length of the array whilst preserving its content.
                ReDim Preserve Items(0 To GrowthFactor * UBound(Items) + 1)
            End If

            ' Add the parsed value to the next array element in Items.
            Select Case Mid$(JSON, i, 1)
            Case "{"
                Set Items(Last) = Deserialize(JSON, i)
            Case Else
                Items(Last) = Deserialize(JSON, i)
            End Select

            ' Parse any "]" structural characters.
            If Mid$(JSON, i, 1) <= " " Then i = SkipWhitespace(JSON, i)
            Do While Mid$(JSON, i, 1) = "]"
                Depth = Depth - 1
                i = i + 1
                If Depth = 0 Then
                    Exit For
                End If
                If Mid$(JSON, i, 1) <= " " Then i = SkipWhitespace(JSON, i)
            Loop

            ' Expect a value separator ",".
            Select Case Mid$(JSON, i, 1)
            Case ","
                Dim Top As Long: Top = Stack(Depth) ' Peek.
                ' Count the element for this sub-array.
                Lengths(Depth).Item(Top) = Lengths(Depth).Item(Top) + 1
            Case vbNullString
                RaiseParseError jsErrorOpenArray, JSON, i + 1
            Case Else
                RaiseParseError jsErrorValueSeparatorExpected, JSON, i
            End Select
            Last = Last + 1
        Next
        ' The loop also ends after a "," at the end of the json.
        If Depth <> 0 And Not Mixed Then RaiseParseError jsErrorOpenArray, JSON, i

        ' Parse again for a mixed or a jagged json-array (IsJaggedArray lowers NDim).
        If Mixed Then
            Again = True
        Else
            Again = IsJaggedArray(Lengths, NDim)
        End If
    Loop While Again

    ' Resize array buffer to actual size.
    ReDim Preserve Items(0 To Last)

    ' Redimension the 1-dimensional array to the proper n-dimensional array and return.
    If NDim = 1 Then
        ' Only happens if the json-array was jagged.
        DeserializeMatrix = Items
    Else
        DeserializeMatrix = RedimensionArray(Items, Lengths, NDim)
    End If
    index = i
    Exit Function

ErrorHandler:
    Err.Raise Err.Number

End Function


Private Function IsJaggedArray(ByRef Lengths() As Dictionary, ByRef NDim As Long) As Boolean
Attribute IsJaggedArray.VB_Description = "Determines whether the n-dimensional json-array is jagged or not."
' For an array all array lengths at a certain depth need to be identical.

    Dim Length As Long, value As Variant
    Dim n As Long
    For n = 2 To NDim
        With Lengths(n)
            Length = .Items()(0)
            For Each value In .Items
                If value <> Length Then
                    NDim = NDim - 1
                    IsJaggedArray = True
                    Exit Function
                End If
            Next
        End With
    Next

End Function


Private Function RedimensionArray( _
    ByRef Items() As Variant, _
    ByRef Lengths() As Dictionary, _
    ByVal NDim As Long _
) As Variant()
Attribute RedimensionArray.VB_Description = "Redimensions an one-dimensional array to a multi-dimensional array."
' The redimensioning is accomplished by swapping SafeArray data pointers.

    ' Create a n-dimensional SafeArray with the proper dimensions and boundaries.
    Dim Count As Long: Count = 1
    Dim rgsabound() As Long: ReDim rgsabound(0 To 2 * NDim - 1)
    Dim n As Long
    For n = 0 To 2 * NDim - 1 Step 2
        rgsabound(n) = Lengths(NDim - n \ 2).Items()(0)
        Count = Count * rgsabound(n)
    Next
    ' Safety check to prevent system crash.
    If Count <> (UBound(Items) - LBound(Items) + 1) Then Err.Raise vbErrorInternalError

    ' Create the actual n-dimensional SafeArray.
    Dim psa As LongPtr: psa = SafeArrayCreate(VBA.vbVariant, NDim, ByVal VarPtr(rgsabound(0)))
    If psa = 0 Then Err.Raise vbErrorOutOfMemory

    Dim Matrix() As Variant, DataPtr As LongPtr
#If API Then
    ' Assign the created n-dimensional SafeArray to the Matrix variable.
    CopyMemory ByVal VarPtrArray(Matrix), psa, vbSizeLongPtr

    ' Save the array data pointer for the Matrix array.
    CopyMemory DataPtr, ByVal psa + vbOffsetArraypvData, vbSizeLongPtr

    ' Copy the array data pointer for the Items array to the Matrix SafeArray.
    CopyMemory ByVal psa + vbOffsetArraypvData, VarPtr(Items(0)), vbSizeLongPtr

    ' Get the SafeArray pointer for the Items array by derefencing the Items variable.
    CopyMemory psa, ByVal VarPtrArray(Items), vbSizeLongPtr

    ' Copy the array data pointer for the Matrix array to the Items SafeArray.
    CopyMemory ByVal psa + vbOffsetArraypvData, DataPtr, vbSizeLongPtr
#Else
    ' The same pointer copies with the Variant ByRef construct instead of the API.
    CopyPtrByRef VarPtrArray(Matrix), psa, VarByRef.vt, VarByRef.ref
    DataPtr = PeekPtrByRef(psa + vbOffsetArraypvData, VarByRef.vt, VarByRef.ref)
    CopyPtrByRef psa + vbOffsetArraypvData, VarPtr(Items(0)), VarByRef.vt, VarByRef.ref
    psa = PeekPtrByRef(VarPtrArray(Items), VarByRef.vt, VarByRef.ref)
    CopyPtrByRef psa + vbOffsetArraypvData, DataPtr, VarByRef.vt, VarByRef.ref
#End If

    ' Return the n-dimensional Matrix.
    RedimensionArray = Matrix

End Function


Private Sub InitializeDecodeTables()
Attribute InitializeDecodeTables.VB_Description = "Fills the lookup tables for Decode."

    ' The character for the character after "\" of a short escape.
    DecodeEscapes(34) = 34      ' "\"""  -> """"
    DecodeEscapes(47) = 47      ' "\/" -> "/"
    DecodeEscapes(92) = 92      ' "\\" -> "\"
    DecodeEscapes(98) = 8       ' "\b" -> VBA.vbBack
    DecodeEscapes(102) = 12     ' "\f" -> VBA.vbFormFeed
    DecodeEscapes(110) = 10     ' "\n" -> VBA.vbLf
    DecodeEscapes(114) = 13     ' "\r" -> VBA.vbCr
    DecodeEscapes(116) = 9      ' "\t" -> VBA.vbTab

    ' The value of the hexadecimal digits 0-9, A-F and a-f; 255 for the others.
    Dim i As Long
    For i = 0 To 127
        HexValues(i) = 255
    Next
    For i = 0 To 9
        HexValues(48 + i) = i
    Next
    For i = 10 To 15
        HexValues(55 + i) = i       ' "A" to "F"
        HexValues(87 + i) = i       ' "a" to "f"
    Next

End Sub


Private Function Decode(ByRef JSON As String) As String
Attribute Decode.VB_Description = "Decodes a printable Ascii string to a Unicode string."
' Decodes the escapes of a json-string.
'
' A string without a "\" is returned as it is. Otherwise InStr jumps from "\" to
' "\" and every run of characters in between is copied to the buffer at once;
' the result is never longer than the json-string. The escapes are decoded with
' lookup tables, reading the characters from a byte array (2 bytes a character).
' An invalid escape raises a parse error: an unknown character after "\", a "\"
' at the end, or a "\u" without 4 hexadecimal digits. Both "\/" and "/" decode
' to "/". A json-string with "\/" therefore keeps its value, but Encode writes
' it back as "/", the shortest form that is allowed.

    Dim i As Long: i = InStr(JSON, "\")
    If i = 0 Then
        Decode = JSON
        Exit Function
    End If

    ' VB allows a quick copy of a string into a byte array.
    Dim Chars() As Byte: Chars = JSON
    ' The result is never longer than the json-string.
    Dim Buffer As String: Buffer = space$(Len(JSON))
    Dim index As Long

    Dim Start As Long: Start = 1
    Dim j As Long, Code As Long, Digit As Long
    Do
        ' Copy the run of characters before the "\".
        If i > Start Then
            Mid$(Buffer, index + 1) = Mid$(JSON, Start, i - Start)
            index = index + i - Start
        End If
        If i = Len(JSON) Then RaiseParseError jsErrorInvalidCharacter, JSON, i, "\"
        ' The character after the "\" starts at byte 2 * i.
        Code = Chars(2 * i) Or Chars(2 * i + 1) * &H100&
        If Code = 117 Then
            ' "\u" and 4 hexadecimal digits.
            If i + 5 > Len(JSON) Then RaiseParseError jsErrorInvalidCharacter, JSON, i, Mid$(JSON, i)
            Code = 0
            For j = 2 * i + 2 To 2 * i + 8 Step 2
                If Chars(j + 1) = 0 And Chars(j) < 128 Then Digit = HexValues(Chars(j)) Else Digit = 255
                If Digit = 255 Then RaiseParseError jsErrorInvalidCharacter, JSON, j \ 2 + 1, Mid$(JSON, i, 6)
                Code = Code * 16 + Digit
            Next
            Start = i + 6
        Else
            ' A short escape like "\n".
            If Code < 128 Then Code = DecodeEscapes(Code) Else Code = 0
            If Code = 0 Then RaiseParseError jsErrorInvalidCharacter, JSON, i + 1, Mid$(JSON, i, 2)
            Start = i + 2
        End If
        index = index + 1
        Mid$(Buffer, index) = VBA.ChrW$(Code)
        i = InStr(Start, JSON, "\")
    Loop Until i = 0

    ' Copy the last run of characters, if any.
    If Start <= Len(JSON) Then
        Mid$(Buffer, index + 1) = Mid$(JSON, Start)
        index = index + Len(JSON) - Start + 1
    End If
    Decode = Left$(Buffer, index)

End Function


Private Function MatchString(ByRef JSON As String, ByRef index As Long) As String
Attribute MatchString.VB_Description = "Matches a string at the current position and advances the index."
' Mid$(json, Index, 1) = """"
'
' InStr jumps from """" to """"; a """" is escaped when an odd number of "\"
' precedes it. Only the characters just before a """" are looked at.
'
' InStr, not VBA.InStr: the qualified VBA.InStr copies the whole string on
' every call (330 µs for a json of 1 MB), the unqualified InStr does not.

    Dim i As Long: i = index
    Do
        ' Find the next """".
        i = InStr(i + 1, JSON, """")
        If i = 0 Then
            ' The end of the json is reached without finding a closing structural character.
            RaiseParseError jsErrorOpenString, JSON, index
        End If
        ' Count the "\" just before it.
        Dim j As Long: j = i - 1
        Do While j > index
            If Mid$(JSON, j, 1) <> "\" Then Exit Do
            j = j - 1
        Loop
    Loop While (i - 1 - j) Mod 2 = 1

    MatchString = Decode(Mid$(JSON, index + 1, i - index - 1))
    index = i + 1

End Function


Private Function SkipWhitespace(ByRef JSON As String, ByVal index As Long) As Long
Attribute SkipWhitespace.VB_Description = "Returns the position of the first character at or after a position that is not whitespace."
' Json whitespace is a space, tab, line feed or carriage return. Beyond the end
' of the json the position is returned as it is. The callers first test
' Mid$(JSON, index, 1) <= " ", so the call is only made for a character that
' might be whitespace.

    Do While index <= Len(JSON)
        Select Case VBA.AscW(Mid$(JSON, index, 1))
        Case 32, 9, 10, 13
            index = index + 1
        Case Else
            Exit Do
        End Select
    Loop
    SkipWhitespace = index

End Function


Private Function MatchValue(ByRef JSON As String, ByRef index As Long) As String
Attribute MatchValue.VB_Description = "Matches a value at the current position and advances the index."

    ' A structural character "{}[],:", whitespace or the end of the json marks the
    ' end of the value.
    Dim i As Long
    For i = index To Len(JSON)
        Select Case VBA.AscW(Mid$(JSON, i, 1))
        Case 123, 125, 91, 93, 44, 58, 32, 9, 10, 13
            Exit For
        End Select
    Next
    MatchValue = Mid$(JSON, index, i - index)
    index = i

End Function


Private Function IsExcelLargeNumberString(ByRef var As Variant) As Boolean
Attribute IsExcelLargeNumberString.VB_Description = "Determines whether a string is a large Excel number or not."
' When you type a number that contains more than 15 digits in a cell in Excel,
' any digits past the 15th place are changed to zeros. In order to prevent data
' loss, a large number with more than 15 digits is conserved as a string.
' https://learn.microsoft.com/en-us/office/troubleshoot/excel/last-digits-changed-to-zeros

    Const ExcelLargeNumber As Long = 15
    If Len(var) <= ExcelLargeNumber Then Exit Function
    If VBA.VarType(var) <> VBA.vbString Then Exit Function

    Dim Count As Long, FractionDetected As Boolean, ThousandSeparator As Long
    Dim i As Long
    For i = 1 To Len(var)
        Select Case VBA.Asc(Mid$(var, i, 1))
        Case 48 To 57 ' "0" To "9"
            Count = Count + 1
        Case 45 ' "-"
            If i <> 1 Then Exit Function
        Case 46 ' "."
            ' Assumed to be thousand separator if DecimalPoint = False
            If DecimalPoint Then
                ' Only one decimal separator is allowed.
                If FractionDetected Then Exit Function
                FractionDetected = True
            ElseIf ThousandSeparator = 0 Then
                ThousandSeparator = i
            ElseIf i = ThousandSeparator + 4 Then
                ThousandSeparator = i
            Else
                ' Not a thousand separator.
                Exit Function
            End If
        Case 44 ' ","
            ' Assumed to be thousand separator if DecimalPoint = True
            If DecimalPoint = False Then
                ' Only one decimal separator is allowed.
                If FractionDetected Then Exit Function
                FractionDetected = True
            ElseIf ThousandSeparator = 0 Then
                ThousandSeparator = i
            ElseIf i = ThousandSeparator + 4 Then
                ThousandSeparator = i
            Else
                ' Not a thousand separator.
                Exit Function
            End If
        End Select
    Next

    IsExcelLargeNumberString = Count > ExcelLargeNumber

End Function


Private Sub RaiseParseError( _
    ByVal error As JSONERROR, _
    ByRef JSON As String, _
    ByVal index As Long, _
    Optional ByVal Info As String _
)
Attribute RaiseParseError.VB_Description = "Raises an error and prints information about the parse error."

    ' The width of the printed line.
    Const Width As Long = 80

    Dim Buffer As StringBuilder

    Append Buffer, ErrorDescription(error)
    If Info = VBA.vbNullString Then
        Append Buffer, VBA.vbCrLf
    Else
        Shorten Buffer
        Append Buffer, " (" & Info & ")." & VBA.vbCrLf
    End If

    ' Create an locator string to indicate the problem character.
    Dim Length As Long: Length = VBA.IIf(index > Len(JSON), index, Len(JSON))
    Dim locator As String: locator = space$(Length)
    Mid$(locator, index, 1) = "^"

    ' Align information string with json over multiple lines if necessary.
    Dim i As Long
    For i = 1 To Length Step Width
        Append Buffer, Mid$(JSON, i, Width) & VBA.vbCrLf
        Append Buffer, Mid$(locator, i, Width) & VBA.vbCrLf
    Next

    Debug.Print Flush(Buffer)

    ' Raise the error to be passed on and handled by the caller routine.
    Err.Raise VBA.vbObjectError Or error, Description:=ErrorDescription(error)

End Sub


Private Sub DataError( _
    ByRef builder As StringBuilder, _
    ByVal error As JSONERROR, _
    Optional ByVal Info As String _
)
Attribute DataError.VB_Description = "Raises an error for VB data that can't be serialized, or puts the message in the json."
' The information, e.g. the VB type name, is added to the description. See the
' directive jsonRaiseDataErrors.

    Dim Message As String: Message = ErrorDescription(error)
    If Len(Info) <> 0 Then Message = Left$(Message, Len(Message) - 1) & " (" & Info & ")."
#If jsonRaiseDataErrors Then
    Err.Raise VBA.vbObjectError Or error, Description:=Message
#Else
    Append builder, Message
    Debug.Print Message
#End If

End Sub


Private Function ErrorDescription( _
    ByVal error As JSONERROR _
) As String
Attribute ErrorDescription.VB_Description = "Returns the description for an error."

    Select Case error
    Case jsErrorInvalidData
        ErrorDescription = "Invalid data: array or object expected."
    Case jsErrorInvalidObject
        ErrorDescription = "Invalid data: object can't be serialized."
    Case jsErrorInvalidType
        ErrorDescription = "Invalid data: type can't be serialized."
    Case jsErrorInvalidJson
        ErrorDescription = "Invalid json syntax."
    Case jsErrorNameSeparatorExpected
        ErrorDescription = "Invalid json: name separator expected."
    Case jsErrorValueSeparatorExpected
        ErrorDescription = "Invalid json: value separator expected."
    Case jsErrorStringExpected
        ErrorDescription = "Invalid json: string expected."
    Case jsErrorOpenString
        ErrorDescription = "Invalid json: string can't be closed (EOF)."
    Case jsErrorOpenValue
        ErrorDescription = "Invalid json: value can't be closed (EOF)."
    Case jsErrorOpenObject
        ErrorDescription = "Invalid json: object can't be closed (EOF)."
    Case jsErrorOpenArray
        ErrorDescription = "Invalid json: array can't be closed (EOF)."
    Case jsErrorDuplicateKey
        ErrorDescription = "Invalid json: duplicate key."
    Case jsErrorInvalidLiteral
        ErrorDescription = "Invalid json: invalid literal."
    Case jsErrorInvalidNumber
        ErrorDescription = "Invalid json: invalid number."
    Case jsErrorInvalidCharacter
        ErrorDescription = "Invalid json: invalid character."
    Case jsErrorExtraCharacter
        ErrorDescription = "Invalid json: extra character(s)."
    Case jsErrorEmptyDimension
        ErrorDescription = "Invalid data: array with an empty dimension."
    End Select

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods - StringBuilder
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' These methods are a standard module version of a basic StringBuilder Class. A
' buffer is used to prevent possible performance degradation due to high volume
' string allocations which may happen during the serialization process.

Private Sub Append(ByRef builder As StringBuilder, ByRef str As String)
Attribute Append.VB_Description = "Appends a string to the buffered string."
' The choice was made to grow the buffer exponentially rather than linearly.
' Besides the headspace, 2 characters are always left free: for the closing
' """" of a string or a date and the structural character after it, so that
' Serialize needs no Adjust for every string.

    ' The characters always left free after the appended string.
    Const Spare As Long = 2

    ' Check the size of the buffer and extend it if necessary.
    Do While Len(builder.Buffer) <= (builder.Count + builder.HeadSpace + Len(str) + Spare)
        If Len(builder.Buffer) = 0 Then
            Const InitialSize As Long = 1024
            builder.Buffer = space$(InitialSize)
        Else
            Const GrowthFactor As Double = 2
            builder.Buffer = builder.Buffer & space$((GrowthFactor - 1) * Len(builder.Buffer))
        End If
    Loop

    ' Copy the argument to the buffer.
    Mid$(builder.Buffer, builder.Count + 1) = str
    builder.Count = builder.Count + Len(str)

End Sub


Private Sub Adjust(ByRef builder As StringBuilder, ByVal amount As Long)
Attribute Adjust.VB_Description = "Adjust the head space of the buffered string."

    builder.HeadSpace = builder.HeadSpace + amount

    If amount > 0 Then
        If Len(builder.Buffer) <= (builder.Count + builder.HeadSpace) Then
            Append builder, vbNullString
        End If
    End If

End Sub


Private Sub Shorten(ByRef builder As StringBuilder, Optional ByVal amount As Long = 1)
Attribute Shorten.VB_Description = "Shortens the buffered string by an amount of characters."

    If amount < builder.Count Then
        builder.Count = builder.Count - amount
    Else
        builder.Count = 0
    End If

End Sub


Private Function Flush(ByRef builder As StringBuilder) As String
Attribute Flush.VB_Description = "Flushes the buffered string."
' The string buffer keeps its current size.

    Flush = Left$(builder.Buffer, builder.Count)
    builder.Count = 0
    builder.HeadSpace = 0

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods - Date conversion
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function ConvertLocalTimeToGMT(Optional ByVal LocalTime As Variant) As Date
Attribute ConvertLocalTimeToGMT.VB_Description = "Converts a local time to the corresponding GMT."

    Dim TZInfo As TimeZoneInformation, TZStatus As TIMEZONESTATUS
    TZStatus = GetTimeZoneInformation(TZInfo)

    If VBA.IsMissing(LocalTime) Then
        LocalTime = VBA.Now
    ElseIf VBA.IsDate(LocalTime) Then
        LocalTime = VBA.CDate(LocalTime)
        ' Determine the time zone status which may be different than for Today.
        TZStatus = DST(LocalTime)
    Else
        Err.Raise vbErrorTypeMismatch, , "Date is invalid"
    End If

    With TZInfo
        Dim GMT As Date
        Select Case TZStatus
        Case TIME_ZONE_ID_DAYLIGHT
            GMT = LocalTime + VBA.TimeSerial(0, .Bias + .DaylightBias, 0)
        Case TIME_ZONE_ID_STANDARD
            GMT = LocalTime + VBA.TimeSerial(0, .Bias, 0)
        Case Else
            GMT = LocalTime
        End Select
    End With

    ConvertLocalTimeToGMT = GMT

End Function


Private Function ConvertGMTToLocalTime(ByVal GMT As Date) As Date
Attribute ConvertGMTToLocalTime.VB_Description = "Converts a GMT to the corresponding local time."

    Dim TZInfo As TimeZoneInformation, TZStatus As TIMEZONESTATUS
    TZStatus = GetTimeZoneInformation(TZInfo)

    With TZInfo
        Dim LocalTime As Date
        LocalTime = GMT - VBA.TimeSerial(0, .Bias + .DaylightBias, 0)
        If DST(LocalTime) <> TIME_ZONE_ID_DAYLIGHT Then
            LocalTime = GMT - VBA.TimeSerial(0, .Bias, 0)
            If DST(LocalTime) <> TIME_ZONE_ID_STANDARD Then
                LocalTime = GMT
            End If
        End If
    End With

    ConvertGMTToLocalTime = LocalTime

End Function


Private Function ConvertLocalTimeToUTC( _
    ByVal LocalTime As Variant, _
    Optional ByVal Delimiter As Boolean = True _
) As String
Attribute ConvertLocalTimeToUTC.VB_Description = "Converts a local serial time to the corresponding UTC."
' For formats see https://en.wikipedia.org/wiki/ISO_8601

    Dim ISOFormat As String
    If VBA.CLng(VBA.Int(LocalTime)) = 0 Then
        If Delimiter Then
            ISOFormat = "hh:nn:ss\Z"
        Else
            ISOFormat = "\Thhnnss\Z"
        End If
    Else
        If Delimiter Then
            ISOFormat = "yyyy-mm-dd\Thh:nn:ss\Z"
        Else
            ISOFormat = "yyyymmdd\Thhnnss\Z"
        End If
    End If

    ConvertLocalTimeToUTC = Format$(ConvertLocalTimeToGMT(LocalTime), ISOFormat)

End Function


Private Function ConvertUTCToLocalTime(ByVal UTC As String) As Date
Attribute ConvertUTCToLocalTime.VB_Description = "Converts a UTC to the corresponding local serial time."
' For formats see https://en.wikipedia.org/wiki/ISO_8601

    Dim LocalTime As Date
    Select Case True
    Case (Right$(UTC, 1) <> "Z")
        Exit Function
    Case UTC Like "####-##-##T##:##:##Z"
        LocalTime = VBA.DateSerial(Mid$(UTC, 1, 4), Mid$(UTC, 6, 2), Mid$(UTC, 9, 2)) + _
                   VBA.TimeSerial(Mid$(UTC, 12, 2), Mid$(UTC, 15, 2), Mid$(UTC, 18, 2))
    Case UTC Like "########T######Z"
        LocalTime = VBA.DateSerial(Mid$(UTC, 1, 4), Mid$(UTC, 5, 2), Mid$(UTC, 7, 2)) + _
                   VBA.TimeSerial(Mid$(UTC, 10, 2), Mid$(UTC, 12, 2), Mid$(UTC, 14, 2))
    Case UTC Like "T######Z"
        LocalTime = VBA.Date + VBA.TimeSerial(Mid$(UTC, 2, 2), Mid$(UTC, 4, 2), Mid$(UTC, 6, 2))
    Case UTC Like "##:##:##Z"
        LocalTime = VBA.Date + VBA.TimeSerial(Mid$(UTC, 1, 2), Mid$(UTC, 4, 2), Mid$(UTC, 7, 2))
    Case Else
        Exit Function
    End Select

    ConvertUTCToLocalTime = ConvertGMTToLocalTime(LocalTime)

End Function


Private Function ConvertSystemTimeToSerialTime(ByRef SystemTime As SystemTime) As Date
Attribute ConvertSystemTimeToSerialTime.VB_Description = "Converts a system time to the corresponding serial time."
' Serial time is the time as implemented with the VB Date type.

    With SystemTime
        ConvertSystemTimeToSerialTime = _
            VBA.DateSerial(.Year, .Month, .Day) + _
            VBA.TimeSerial(.Hour, .Minute, .Second)
    End With

End Function


Private Function ConvertDateToSystemTime(ByVal serialTime As Date) As SystemTime
Attribute ConvertDateToSystemTime.VB_Description = "Converts a serial time to the corresponding system time."

    With ConvertDateToSystemTime
        .Year = VBA.Year(serialTime)
        .Month = VBA.Month(serialTime)
        .Day = VBA.Day(serialTime)
        .Hour = VBA.Hour(serialTime)
        .Minute = VBA.Minute(serialTime)
        .Second = VBA.Second(serialTime)
        ' vbSunday = 1 whereas for the Windows system time Monday = 1.
        .DayOfWeek = VBA.WeekDay(serialTime) - 1
        .MilliSeconds = 0
    End With

End Function


Private Function DST(ByVal serialDate As Date) As TIMEZONESTATUS
Attribute DST.VB_Description = "Determines whether a date falls within the daylight saving period or not."

    Dim TZInfo As TimeZoneInformation
    If VBA.CLng(VBA.Int(serialDate)) = 0 Then
        DST = GetTimeZoneInformation(TZInfo)
        Exit Function
    End If

    Dim Year As Long: Year = VBA.Year(serialDate)
    If GetTimeZoneInformationForYear(Year, 0, TZInfo) = 0 Then
        Exit Function
    End If

    ' TZInfo.DaylightDate.Month = 0 means the timezone does not observe DST.
    If TZInfo.DaylightDate.Month = 0 Then
        DST = TIME_ZONE_ID_STANDARD
        Exit Function
    End If

    Dim DaylightDate As Date, StandardDate As Date
    With TZInfo.DaylightDate
        DaylightDate = LastDayOfWeekOfMonth(.DayOfWeek + 1, VBA.DateSerial(Year, .Month, 1))
    End With
    With TZInfo.StandardDate
        StandardDate = LastDayOfWeekOfMonth(.DayOfWeek + 1, VBA.DateSerial(Year, .Month, 1))
    End With

    If serialDate >= DaylightDate And serialDate <= StandardDate Then
        DST = TIME_ZONE_ID_DAYLIGHT
    Else
        DST = TIME_ZONE_ID_STANDARD
    End If

End Function


Private Function LastDayOfWeekOfMonth( _
    ByVal DayOfWeek As VbDayOfWeek, _
    ByVal serialTime As Date _
) As Date
Attribute LastDayOfWeekOfMonth.VB_Description = "Returns the date of the last weekday in the month specified by a date."

    If DayOfWeek < 1 Or DayOfWeek > 7 Then Err.Raise vbErrorInvalidProcedureCall, , "Day is invalid"

    ' Set DateValue to VBA.Date if it doesn't contain a date.
    If VBA.CLng(VBA.Int(serialTime)) = 0 Then serialTime = VBA.Date

    Dim LastDay As Date: LastDay = VBA.DateSerial(VBA.Year(serialTime), VBA.Month(serialTime) + 1, 0)
    LastDayOfWeekOfMonth = LastDay - (7 + VBA.WeekDay(LastDay) - DayOfWeek) Mod 7

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods - Utility
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function LocaleDecimalPoint() As Boolean
Attribute LocaleDecimalPoint.VB_Description = "Determine the locale for the decimal point."

    ' Quick way to establish the locale for the decimal point without API.
    Select Case Mid$(VBA.CStr(1.1), 2, 1)
    Case "."
        ' The json decimal separator is always a point.
        LocaleDecimalPoint = True
    Case ","
        ' The output VB decimal separator can be a comma.
        LocaleDecimalPoint = False
    Case Else
        ' Any other decimal separator is not allowed.
        Err.Raise vbErrorInvalidProcedureCall, , "Locale decimal separator is invalid"
    End Select

End Function


Private Function ArrDims(ByRef arr As Variant) As Long
Attribute ArrDims.VB_Description = "Returns the number of dimensions of an array."
' Returns 0 for an empty array.
' LBound raises on out-of-range n — used as the loop exit condition.

    If Not VBA.IsArray(arr) Then Err.Raise vbErrorTypeMismatch

    On Error GoTo ForcedError
    Dim n As Long: n = 1
    Do While LBound(arr, n) <= UBound(arr, n)
        n = n + 1
    Loop

ForcedError:
    ArrDims = n - 1
    Err.Clear

End Function


#If API Then
Private Function CollKeys(ByVal coll As Collection) As String()
' Returns an empty string array for an empty Collection.

    ' The null pointer constant (not VB defined).
    Const vbNullPtr As LongPtr = 0

    If coll.Count = 0 Then Exit Function

    ' Traverse the double linked list and copy all the key strings.
    Dim Keys() As String: ReDim Keys(0 To coll.Count - 1)
    Dim NodePtr As LongPtr, KeyPtr As LongPtr
    NodePtr = ObjPtr(coll) + vbOffsetCollectionHead
    Dim i As Long
    For i = 0 To coll.Count - 1
        CopyMemory NodePtr, ByVal NodePtr, vbSizeLongPtr
        CopyMemory KeyPtr, ByVal NodePtr + vbOffsetCollectionKey, vbSizeLongPtr
        If KeyPtr <> vbNullPtr Then
            SysReAllocString VarPtr(Keys(i)), KeyPtr
        End If
        NodePtr = NodePtr + vbOffsetCollectionNext
    Next

    CollKeys = Keys

End Function
#Else
Private Sub InitializeVarByRef()
Attribute InitializeVarByRef.VB_Description = "Initializes the Variant ByRef construct."

    VarByRef.vt = VarPtr(VarByRef.ref)
    CopyMemory VarByRef.vt, VBA.vbInteger Or VT_BYREF, vbSizeInteger

End Sub


Private Sub CopyPtrByRef( _
    ByVal address As LongPtr, _
    ByVal value As LongPtr, _
    ByRef vt As Variant, _
    ByRef ref As Variant _
)
Attribute CopyPtrByRef.VB_Description = "Copies a pointer to a memory address using the Variant ByRef construct."

#If Win64 Then
    ' A LongLong by reference is not supported (error 458). The 8 bytes of the
    ' LongLong are read as a Currency by changing the type only, and that
    ' Currency is written by reference: the bits are copied, nothing is calculated.
    VarByRef.ref = value
    vt = VBA.vbCurrency
    Dim Bits As Currency: Bits = ref

    VarByRef.ref = address
    vt = VBA.vbCurrency Or VT_BYREF
    ref = Bits
#Else
    VarByRef.ref = address
    vt = VBA.vbLong Or VT_BYREF
    ref = value
#End If

End Sub


Private Function PeekPtrByRef( _
    ByVal address As LongPtr, _
    ByRef vt As Variant, _
    ByRef ref As Variant _
) As LongPtr
Attribute PeekPtrByRef.VB_Description = "Returns the pointer at a memory address using the Variant ByRef construct."

    VarByRef.ref = address
    vt = vbLongPtr Or VT_BYREF
    PeekPtrByRef = ref

End Function


Private Function CollKeys(ByVal coll As Collection, ByRef vt As Variant) As String()
Attribute CollKeys.VB_Description = "Returns an array containing the keys of all the items in a VB Collection."
' Returns an empty string array for an empty Collection.

    If coll.Count = 0 Then Exit Function

    ' Traverse the double linked list and copy all the key strings.
    Dim Keys() As String: ReDim Keys(0 To coll.Count - 1)
    Dim NodePtr As LongPtr
    NodePtr = ObjPtr(coll) + vbOffsetCollectionHead - vbOffsetCollectionNext
    Dim i As Long
    For i = 0 To coll.Count - 1
        VarByRef.ref = NodePtr + vbOffsetCollectionNext
        vt = vbLongPtr Or VT_BYREF
        NodePtr = VarByRef.ref
        VarByRef.ref = NodePtr + vbOffsetCollectionKey
        vt = VBA.vbString Or VT_BYREF
        Keys(i) = VarByRef.ref
    Next

    CollKeys = Keys

End Function
#End If
