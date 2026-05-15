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
' Version: 2025.09.09
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

' Convert a large number (> 15 digits) to a VB string and vice versa.
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


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()

    DecimalPoint = LocaleDecimalPoint

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

    If VBA.IsObject(var) Or VBA.IsArray(var) Then
        Dim builder As StringBuilder
        Serialize var, builder
        Stringify = Flush(builder)
    Else
        Stringify = ErrorDescription(jsErrorInvalidData)
    End If

End Function


Public Function Parse(ByVal JSON As String) As Variant
Attribute Parse.VB_Description = "Deserializes a json string to a VB data structure."
' Parse is the front-end for the deserialization process.

    On Error GoTo ErrorHandler

    ' Remove the whitespace in the local ByVal copy of Json.
    JSON = Simplify(JSON)

    ' Initiate the deserialization process.
    Dim index As Long: index = 1
    If Left$(JSON, 1) & Right$(JSON, index) = "{}" Then
        Set Parse = Deserialize(JSON, index)
    Else
        Parse = Deserialize(JSON, index)
    End If

    If index <= Len(JSON) Then
        RaiseParseError jsErrorExtraCharacter, JSON, index
    End If
    Exit Function

ErrorHandler:
    ' Tidy up to prevent cascading errors.
    If Left$(JSON, 1) & Right$(JSON, index) = "{}" Then
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
            builder.Count = builder.Count + 1
            Mid$(builder.Buffer, builder.Count) = """"
            Append builder, MatchString(JSON, i)
            builder.Count = builder.Count + 1
            Mid$(builder.Buffer, builder.Count) = """"
            i = i - 1
        Case "{"
            builder.Count = builder.Count + 1
            Mid$(builder.Buffer, builder.Count) = Char
            If MaxDepth = False Then
                Depth = Depth + 1
                builder.Count = builder.Count + 1
                Mid$(builder.Buffer, builder.Count) = VBA.vbCrLf
                Append builder, String$(Depth, indentation)
            End If
        Case "}"
            If MaxDepth = False Then
                Depth = Depth - 1
                builder.Count = builder.Count + 1
                Mid$(builder.Buffer, builder.Count) = VBA.vbCrLf
                Append builder, String$(Depth, indentation)
            End If
            builder.Count = builder.Count + 1
            Mid$(builder.Buffer, builder.Count) = Char
        Case "["
            If compact Then
                ' Determine whether this is the maximum array depth.
                Dim match As Long: match = VBA.InStr(i + 1, JSON, "[")
                If match = 0 Then
                    MaxDepth = True
                Else
                    MaxDepth = match > VBA.InStr(i + 1, JSON, "]")
                End If
            End If
            builder.Count = builder.Count + 1
            Mid$(builder.Buffer, builder.Count) = Char
            If MaxDepth = False Then
                Depth = Depth + 1
                builder.Count = builder.Count + 1
                Mid$(builder.Buffer, builder.Count) = VBA.vbCrLf
                Append builder, String$(Depth, indentation)
            End If
        Case "]"
            If MaxDepth Then
                MaxDepth = False
            Else
                Depth = Depth - 1
                builder.Count = builder.Count + 1
                Mid$(builder.Buffer, builder.Count) = VBA.vbCrLf
                Append builder, String$(Depth, indentation)
            End If
            builder.Count = builder.Count + 1
            Mid$(builder.Buffer, builder.Count) = Char
        Case ","
            builder.Count = builder.Count + 1
            Mid$(builder.Buffer, builder.Count) = Char
            If MaxDepth Then
                builder.Count = builder.Count + 1
                Mid$(builder.Buffer, builder.Count) = space
            Else
                builder.Count = builder.Count + 1
                Mid$(builder.Buffer, builder.Count) = VBA.vbCrLf
                Append builder, String$(Depth, indentation)
            End If
        Case ":"
            Mid$(builder.Buffer, builder.Count + 1) = Char & space
            builder.Count = builder.Count + 2
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
' This main decision tree is called recursively in the serialization process.

    ' Headspace for adding structural characters.
    Const HeadSpace As Long = 2

    ' At this point there is always at least 1 headspace left!
    Select Case VBA.VarType(var)
    Case VBA.vbString
        builder.Count = builder.Count + 1
        Mid(builder.Buffer, builder.Count) = """"
        ' Headspace for adding structural characters.
        Adjust builder, HeadSpace
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
        Adjust builder, -HeadSpace
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
        Select Case VBA.TypeName(var)
        Case "Dictionary"
            ' Serialize a VB Dictionary to a json-object.
            If var.Count = 0 Then
                Append builder, "{}"
            Else
                SerializeDictionary var, builder
            End If
        Case "Collection"
            ' Serialize a VB Collection to a json-object.
            If var.Count = 0 Then
                Append builder, "{}"
            Else
                SerializeCollection var, builder
            End If
        Case Else
            ' Serialize other VB objects.
            Append builder, "{" & VBA.TypeName(var) & "}"
            Debug.Print ErrorDescription(jsErrorInvalidObject) & ": " & VBA.TypeName(var)
        End Select
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
        ' Headspace for adding structural characters.
        Adjust builder, HeadSpace
        ' Next Append will check for the extra headspace.
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
        Adjust builder, -HeadSpace
    Case Else
        ' Serialization for other VB types.
        ' Not implemented for vbUserDefinedType, vbError, vbDataObject and vbEmpty.
        Append builder, VBA.TypeName(var)
        Debug.Print ErrorDescription(jsErrorInvalidType) & ": " & VBA.TypeName(var)
    End Select

End Sub


Private Sub SerializeDictionary( _
    ByRef var As Variant, _
    ByRef builder As StringBuilder _
)
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
        Append builder, Encode(CStr(Keys(i)))
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
#If jsonConvertCollectionKeys Then
        If Len(Keys(i)) <> 0 Then
            Append builder, Encode(Keys(i))
        End If
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
        Dim n As Long
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


Private Function Encode(ByRef str As String) As String
' Encodes a Unicode/Ascii string to a printable Ascii string.
'
' Buffering printable ascii characters achieves a significant performance
' improvement over the straightforward one-by-one character approach.
'
' Best case timer ratio's for 100% printable ascii characters
' n         1       10      100    1000
' r      0.98      3.5     16.1    28.4
'
' Worst case timer ratio's for 100% non-printable ascii characters
' n         1       10      100    1000
' r      0.77     0.90     0.95    0.97
'
' n = length of input string
' r = ratio of timer(one-by-one) / timer(improved)

    ' Lookup table for character encoding.
    Static Table(0 To 127) As Byte
    If Table(8) = 0 Then
        Table(8) = 98         ' VBA.vbBack -> "\b"
        Table(9) = 116        ' VBA.vbTab -> "\t"
        Table(10) = 110       ' VBA.vbLf -> "\n"
        Table(12) = 102       ' VBA.vbFormFeed -> "\f"
        Table(13) = 114       ' VBA.vbCr -> "\r"
        Table(34) = 34        ' """"  -> "\"""
'       Table(47) = 47        ' "/" -> "\/" ' RFC 4627 permits but does not require escaping "/"
        Table(92) = 92        ' "\" -> "\\"
    End If

    ' The maximum buffer size when all characters would be \u encoded.
    Dim Buffer As String: Buffer = Space$(6 * Len(str))
    Dim index As Long

    Dim Start As Long: Start = 1
    Dim i As Long, Code As Long
    For i = 1 To Len(str)
        Code = VBA.AscW(Mid$(str, i, 1))
        If Code < 0 Then
            ' Extended Ascii and unicode characters.
            ' AscW returns as a signed integer which needs to be converted to a unsigned number.
            Mid$(Buffer, index + 1) = Mid$(str, Start, i - Start)
            index = index + i - Start + 6
            Mid$(Buffer, index - 5) = "\u"
            Mid$(Buffer, index - 3) = Right$("0000" & Hex$(Code + 65536), 4)
            Start = i + 1
        ElseIf Code >= 127 Then
            ' Extended Ascii and unicode characters.
            Mid$(Buffer, index + 1) = Mid$(str, Start, i - Start)
            index = index + i - Start + 6
            Mid$(Buffer, index - 5) = "\u"
            Mid$(Buffer, index - 3) = Right$("0000" & Hex$(Code), 4)
            Start = i + 1
        ElseIf Table(Code) <> 0 Then
            ' Special Ascii characters.
            Mid$(Buffer, index + 1) = Mid$(str, Start, i - Start)
            index = index + i - Start + 2
            Mid$(Buffer, index - 1) = "\"
            Mid$(Buffer, index) = Chr$(Table(Code))
            Start = i + 1
        ElseIf Code < 32 Then
            ' Non-printable Ascii characters.
            Mid$(Buffer, index + 1) = Mid$(str, Start, i - Start)
            index = index + i - Start + 6
            Mid$(Buffer, index - 5) = "\u"
            Mid$(Buffer, index - 3) = Right$("0000" & Hex$(Code), 4)
            Start = i + 1
        End If
    Next

    If Start = 1 Then
        Encode = str
    Else
        ' Append the last buffered part of the string.
        Mid$(Buffer, index + 1) = Mid$(str, Start)
        index = index + Len(str) - Start + 1
        Encode = Mid$(Buffer, 1, index)
    End If

End Function


Private Function Deserialize( _
    ByRef JSON As String, _
    ByRef index As Long _
) As Variant
' This main decision tree is called recursively in the deserialization process.

    On Error GoTo ErrorHandler

    Dim value As String
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
        If Mid$(JSON, index + 1, 1) = "}" Then
            ' Deserialize empty json-object to an empty VB Dictionary.
            Set Deserialize = New Dictionary
            index = index + 2
        Else
            ' Deserialize json-object to VB Dictionary or VB Collection.
            Set Deserialize = DeserializeDictionary(JSON, index)
        End If
    Case "["
        ' Count the number of consecutive '[' characters at the Index position.
        Dim NDim As Long
        Do While Mid$(JSON, index + NDim, 1) = "["
            NDim = NDim + 1
        Loop
        ' If the inner dimension is an empty array it counts as an element.
        If Mid$(JSON, index + NDim, 1) = "]" Then
            NDim = NDim - 1
        End If
        Select Case NDim
        Case 0
            ' Deserialize empty json-array to an empty VB Variant() array.
            Deserialize = VBA.Array()
            index = index + 2
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
        ElseIf VBA.InStr(value, "e") <> 0 Or VBA.InStr(value, "E") <> 0 Then
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
        ElseIf VBA.InStr(value, ".") = 0 Then
            On Error Resume Next
            ' Convert json-number to a VB Long.
            Deserialize = VBA.CLng(value)
            If Err.Number = 0 Then Exit Function
#If Win64 Then
            ' Convert json-number to a VB LongLong.
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
' Json doesn't require a unique set of keys and allows multiple keys to be empty
' strings. A VB Dictionary doesn't allow duplicate keys, but one can be an empty
' string. If duplicate empty keys are detected, an attempt will be made to parse
' the json-object to a Collection which does allow for multiple empty keys.

    On Error GoTo ErrorHandler

    Dim Items As Dictionary: Set Items = New Dictionary

    Dim i As Long, key As String
    For i = index + 1 To Len(JSON)
        ' Expect key.
        If Mid$(JSON, i, 1) = """" Then
            key = MatchString(JSON, i)
        Else
#If jsonDeserializeUnquotedKeys Then
            key = MatchValue(JSON, i)
#Else
            RaiseParseError jsErrorStringExpected, JSON, i
#End If
        End If

        ' Expect name separator ":".
        If Mid$(JSON, i, 1) = ":" Then
            i = i + 1
        Else
            RaiseParseError jsErrorNameSeparatorExpected, JSON, i
        End If

        ' Add key : (parsed) value pair to the Dictionary.
        Items.Add key, Deserialize(JSON, i)

        ' Expect a value separator "," or object-end "}".
        Select Case Mid$(JSON, i, 1)
        Case "}"
            Exit For
        Case vbNullString
            RaiseParseError jsErrorOpenObject, JSON, i
        Case Is <> ","
            RaiseParseError jsErrorValueSeparatorExpected, JSON, i
        End Select
    Next
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
' Json doesn't require a unique set of keys and allows multiple keys to be empty
' strings. A VB Collection doesn't allow duplicate keys, but keys can be omitted
' as left as empty strings. If duplicate keys are detected the parsing fails.

    On Error GoTo ErrorHandler

    Dim Items As Collection: Set Items = New Collection

    Dim i As Long, key As String
    For i = index + 1 To Len(JSON)
        ' Expect key.
        If Mid$(JSON, i, 1) = """" Then
            key = MatchString(JSON, i)
        Else
#If jsonDeserializeUnquotedKeys Then
            key = MatchValue(JSON, i)
#Else
            RaiseParseError jsErrorStringExpected, JSON, i
#End If
        End If

        ' Expect name separator ":".
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
        Select Case Mid$(JSON, i, 1)
        Case "}"
            Exit For
        Case vbNullString
            RaiseParseError jsErrorOpenObject, JSON, i + 1
        Case Is <> ","
            RaiseParseError jsErrorValueSeparatorExpected, JSON, i
        End Select
    Next
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
        Select Case Mid$(JSON, i, 1)
        Case "{"
            Set Items(Last) = Deserialize(JSON, i)
        Case Else
            Items(Last) = Deserialize(JSON, i)
        End Select

        ' Expect a value separator "," or array-end "]".
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

    ' Lengths stores all arrays lengths at a certain depth.
    Dim Lengths() As Variant

    Dim Last As Long
    Do
        Last = 0 ' Important to set to 0 for jagged arrays!

        ReDim Stack(1 To NDim)

        ' Initialize the bookkeeping of the sub-array lengths.
        ReDim Lengths(1 To NDim)
        Dim n As Long
        For n = 1 To NDim
            Set Lengths(n) = New Dictionary
        Next

        ' The recursion depth is limited by NDim for handling jagged arrays.
        Dim Depth As Long
        Dim i As Long
        For i = index To Len(JSON)
            ' Parse any "[" structural characters.
            Do While Mid$(JSON, i, 1) = "[" And Depth < NDim
                If Mid$(JSON, i + 1, 1) = "]" Then
                    Exit Do
                End If
                Depth = Depth + 1
                Stack(Depth) = i
                ' Start counting the elements in this sub-array.
                Lengths(Depth).Add i, 1
                i = i + 1
            Loop

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
            Do While Mid$(JSON, i, 1) = "]"
                Depth = Depth - 1
                i = i + 1
                If Depth = 0 Then
                    Exit For
                End If
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
    Loop While IsJaggedArray(Lengths, NDim)

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


Private Function IsJaggedArray(ByRef Lengths() As Variant, ByRef NDim As Long) As Boolean
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
    ByRef Lengths() As Variant, _
    ByVal NDim As Long _
) As Variant()
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

    ' Assign the created n-dimensional SafeArray to the Matrix variable.
    Dim Matrix() As Variant: CopyMemory ByVal VarPtrArray(Matrix), psa, vbSizeLongPtr

    ' Save the array data pointer for the Matrix array.
    Dim DataPtr As LongPtr: CopyMemory DataPtr, ByVal psa + vbOffsetArraypvData, vbSizeLongPtr

    ' Copy the array data pointer for the Items array to the Matrix SafeArray.
    CopyMemory ByVal psa + vbOffsetArraypvData, VarPtr(Items(0)), vbSizeLongPtr

    ' Get the SafeArray pointer for the Items array by derefencing the Items variable.
    CopyMemory psa, ByVal VarPtrArray(Items), vbSizeLongPtr

    ' Copy the array data pointer for the Matrix array to the Items SafeArray.
    CopyMemory ByVal psa + vbOffsetArraypvData, DataPtr, vbSizeLongPtr

    ' Return the n-dimensional Matrix.
    RedimensionArray = Matrix

End Function


Private Function Decode(ByRef JSON As String) As String

    ' Lookup table for character decoding.
    Static Table(0 To 127) As Byte
    If Table(34) = 0 Then
        Table(34) = 34        ' """"
        Table(47) = 47        ' "\/" -> "/"
        Table(92) = 92        ' "\"
        Table(98) = 8         ' "b"
        Table(102) = 12       ' "f"
        Table(110) = 10       ' "n"
        Table(114) = 13       ' "r"
        Table(116) = 9        ' "t"
    End If

    Dim Buffer As StringBuilder
    Dim Start As Long: Start = 1
    Dim i As Long, Code As Long
    Do
        ' Find next escaped character or unicode.
        i = VBA.InStr(Start, JSON, "\")
        If i = 0 Then Exit Do
        ' Append the buffered string.
        Append Buffer, Mid$(JSON, Start, i - Start)
        Code = VBA.Asc(Mid$(JSON, i + 1, 1))
        If Code = 117 Then ' "u"
            Append Buffer, ChrW$("&H" & Mid$(JSON, i + 2, 4))
            i = i + 4
        ElseIf Table(Code) <> 0 Then
            Append Buffer, Chr$(Table(Code))
        Else
            RaiseParseError jsErrorInvalidCharacter, JSON, i + 1
        End If
        Start = i + 2
    Loop

    If Start = 1 Then
        ' Just append the whole string as all characters were uncoded.
        Decode = JSON
    Else
        ' Append the last buffered part of the string and flush the buffer.
        Append Buffer, Mid$(JSON, Start)
        Decode = Flush(Buffer)
    End If

End Function


Private Function MatchString(ByRef JSON As String, ByRef index As Long) As String
' Mid$(json, Index, 1) = """"

    ' Finding the next unescaped """" marks the end of the string.
    Dim i As Long
    For i = index + 1 To Len(JSON)
        Select Case Mid$(JSON, i, 1)
        Case """"
            MatchString = Decode(Mid$(JSON, index + 1, i - index - 1))
            index = i + 1
            Exit Function
        Case "\" ' Escape
            i = i + 1
        End Select
    Next

    ' The end of the json is reached without finding a closing structural character.
    RaiseParseError jsErrorOpenString, JSON, index

End Function


Private Function MatchValue(ByRef JSON As String, ByRef index As Long) As String

    Const StructuralCharacters As String = "{}[],:"

    ' Finding a structural character marks the end of the value.
    Dim i As Long
    For i = index To Len(JSON)
        If VBA.InStr(StructuralCharacters, Mid$(JSON, i, 1)) <> 0 Then
            MatchValue = Mid$(JSON, index, i - index)
            index = i
            Exit Function
        End If
    Next

    ' The end of the json is reached without finding the closing structural character.
    RaiseParseError jsErrorOpenValue, JSON, index

End Function


Private Function IsExcelLargeNumberString(ByRef var As Variant) As Boolean
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
    Dim locator As String: locator = Space$(Length)
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


Private Function ErrorDescription( _
    ByVal error As JSONERROR _
) As String

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
    End Select

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods - StringBuilder
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' These methods are a standard module version of a basic StringBuilder Class. A
' buffer is used to prevent possible performance degradation due to high volume
' string allocations which may happen during the serialization process.

Private Sub Append(ByRef builder As StringBuilder, ByRef str As String)
' The choice was made to grow the buffer exponentially rather than linearly.

    ' Check the size of the buffer and extend it if necessary.
    Do While Len(builder.Buffer) <= (builder.Count + builder.HeadSpace + Len(str))
        If Len(builder.Buffer) = 0 Then
            Const InitialSize As Long = 1024
            builder.Buffer = Space$(InitialSize)
        Else
            Const GrowthFactor As Double = 2
            builder.Buffer = builder.Buffer & Space$((GrowthFactor - 1) * Len(builder.Buffer))
        End If
    Loop

    ' Copy the argument to the buffer.
    Mid$(builder.Buffer, builder.Count + 1) = str
    builder.Count = builder.Count + Len(str)

End Sub


Private Sub Adjust(ByRef builder As StringBuilder, ByVal amount As Long)

    builder.HeadSpace = builder.HeadSpace + amount

    If amount > 0 Then
        If Len(builder.Buffer) <= (builder.Count + builder.HeadSpace) Then
            Append builder, vbNullString
        End If
    End If

End Sub


Private Sub Shorten(ByRef builder As StringBuilder, Optional ByVal amount As Long = 1)

    If amount < builder.Count Then
        builder.Count = builder.Count - amount
    Else
        builder.Count = 0
    End If

End Sub


Private Function Flush(ByRef builder As StringBuilder) As String
' The string buffer keeps its current size.

    Flush = Left$(builder.Buffer, builder.Count)
    builder.Count = 0
    builder.HeadSpace = 0

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods - Date conversion
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function ConvertLocalTimeToGMT(Optional ByVal LocalTime As Variant) As Date

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
' Serial time is the time as implemented with the VB Date type.

    With SystemTime
        ConvertSystemTimeToSerialTime = _
            VBA.DateSerial(.Year, .Month, .Day) + _
            VBA.TimeSerial(.Hour, .Minute, .Second)
    End With

End Function


Private Function ConvertDateToSystemTime(ByVal serialTime As Date) As SystemTime

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

    VarByRef.vt = VarPtr(VarByRef.ref)
    CopyMemory VarByRef.vt, VBA.vbInteger Or VT_BYREF, vbSizeInteger

End Sub


Private Function CollKeys(ByVal coll As Collection, ByRef vt As Variant) As String()
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
