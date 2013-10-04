VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Validator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'Designed By engineersmnky 10/4/2013
'Validator acts as a validation engine for VBA Applications
'Validator's formatting was inspired by Rails validations and uses as similar a foramtting style as possible
'It must be used in conjunction with ValidatorErrors which will hold the name and error messgae for each object that fails validation
' Validator includes multiple methods for retreiving these error including
' errors - Which returns the full collection of Errors
' error_messages - which returns only the messages for objects that failed validation
' error_keys - returns only the name for objects that failed validation
' uniq_keys - returns a unique collection of errors based on name so only the first failed validation for a named object will be included in this collection
'Usage:
'       Dim v As New Validator
'       With v
'            .validates 123, "Number", numericalityOf:=True, only_integer:=True, greater_than_or_equal_to:=11
'            .validates "String", "String", stringnessOf:=True, length:=6, contains:="ing", begins_with:="S"
'            .validates "12345","Number2",numericalityOf:=True, greater_than:=12344, force:=vbInteger, stringnessOf:=True, min_length:=3, begins_with:="1"
'       End With
'
'       If v.is_valid Then
'           Do Something
'       Else
'           Do something Else
'       End If
' Alternate Usage:
'       Dim v As New Validator
'       v.value = 123
'       v.name = "Number"
'       ^^^ is the same as v.validates 123,"Number"
'
'       v.numericality only_integer:=True, greater_than_or_equal_to:=11
'       v.stringness length:=6, contains:="ing", begins_with:="S"
'       v.dateness 
' Each validates call returns a Boolean so you can use this in conditionals as well.
'       If v.stringness(length:=2) then
'           Msgbox "String is 2 Characters Long."
'       Else
'            Msgbox "String is not 2 Characters Long."
'       End If
'
' Custom Validation:
'   If you find the builtin methods are not enough to handle your validation it comes with a .custom_validation function which allows you to 
'       specify a Boolean statement with an optional name and error_message
'


Private pname As String
Private test_val As Variant
Private error_list As New Collection
Private true_val As Variant


'returns the current test name
Public Property Get name() As String
    name = pname
End Property

'name allows you to change the name without changing the test value
Public Property Let name(value As String)
    pname = value
End Property

'returns the current true value
Public Property Get value() As Variant
    value = true_val
End Property

'returns the current test value may be different data type based on test specifications
Public Property Get test_value() As Variant
    test_value = test_val
End Property

'value allows you to change the value without changing the name
Public Property Let value(value As Variant)
    true_val = val
    test_val = true_val
End Property

'errors() returns the full error collection with field and error
Public Property Get errors() As Collection
    Set errors = error_list
End Property


'error_messages returns just the errors for the error collection
Public Function error_messages() As Collection
    Dim messages As New Collection
    Dim e As ValidatorError
    For Each e In error_list
        messages.Add e.error_message
    Next e
    Set error_messages = messages
End Function

'error_keys() returns just the keys for error collection
Public Function error_keys() As Collection
    Dim keys As New Collection
    Dim e As ValidatorError
    For Each e In error_list
        keys.Add e.field
    Next e
    Set error_keys = keys
End Function

'returns only unique names that failed validation good for flagging fields after validation failure
Public Function uniq_keys() As Collection
On Error Resume Next
    Dim unique_errors As New Collection
    Dim e As ValidatorError
    For Each e In error_list
        If Not (in_collection(unique_errors, e.field)) Then
            unique_errors.Add e.field, e.field
        End If
    Next e
    Set uniq_keys = unique_errors
End Function


'valid returns the validtiy of the complete validation
Public Property Get is_valid() As Boolean
    is_valid = Not (error_list.Count > 0)
End Property

' validates allows you to set a values and an optional name for each validation
' Options are:
'       Shared:
'            allow_null (Boolean) - returns true for validation even if value is null * Does not apply to presence
'            other_than (Variant) - returns true if the value is something other than what is specified * Shared By numericality and dateness
'       presenceOf: 
'           returns true is object has a value
'       numericalityOf options:(Default=False)
'           only_integer (Boolean) - checks to see if val is a vbInteger
'           allow_null (Boolean) - this is the only option that does not require a number as it will return its own value if val is Null
'           is_equal_to (Variant) - checks to see if test_val = is_equal_to
'           greater_than (Variant) - will check to see if val is greater than this value and return (takes presidence over _or_equal_to)
'           greater_than_or_equal_to (Variant) - same a greater than but with an equality check
'           less_than (Variant) - will check to see if a value is less than this value and return (takes presidence over _or_equal_to)
'           less_than_or_equal_to (Variant) - same as less than with an equality check
'           odd (Boolean) - checks to see if val is odd if True
'           even (Boolean) - checks to see if a val is even if True
'           is_type (VBA.vbVarType) - checks to see if val is of a specific data_type
'           force(VBA.VbVarType) - attempts to force test_val to a specified data-type prior to testing 
'                     * This value can be retrieved afterwards with .test_value (unless used with stringness strict:=False(Default))
'
'       stringnessOf options: (Default=False)
'           allow_blank (Boolean) - this is similar to allow_null only it will fail if the sting is null but pass if it is an empty string ""
'           length(Integer) - check if the string is a specified length
'           min_length(Integer) - check if the length of a string is greater than or equal to min length
'           max_length(Integer) - check is a string is short than or equal to max length
'           begins_with (String) - Check is a string begins with a specified string
'           ends_with (String) - Check is a string ends with a specified string (can be used with case_sensitive)
'           contains (String) - Check is a string contains a specified string (can be used with case_sensitive)
'           matches (String) - Check is a string matches a given regex pattern (case sensitive has no effect on this method)
'           case_sensitive (Boolean) - to be used in conjunction with begins_with, ends_with, and contains
'           strict (Boolean) - if strict it validates test_val as is otherwise it attempts to make it a string before testing  
'                   *This value can be retrieved afterwards with .test_value (overrides numericality force)
''
'       datenessOf options:(Default=False)
'           allow_null (Boolean) - this is the only option that does not require a date as it will return its own value if test_val is Null
'           after (Date) - will check to see if test_val is after this value and return (takes presidence over on_or_)
'           on_or_after (Date) - same a after but with an equality check
'           before(Date) - will check to see if a value is before this value and return (takes presidence over on_or_)
'           on_or_before (Date) - same as before with an equality check
'
' example: validates 123.45, "Number", [Options]
Public Function validates(value As Variant, Optional name As String, _
    Optional presenceOf As Boolean = False, _
    Optional numericalityOf As Boolean = False, _
    Optional only_integer As Boolean = False, Optional allow_null As Boolean = False, _
    Optional is_equal_to As Variant, Optional greater_than As Variant, _
    Optional greater_than_or_equal_to As Variant, Optional less_than As Variant, _
    Optional less_than_or_equal_to As Variant, Optional other_than As Variant, _
    Optional odd As Boolean = False, Optional even As Boolean = False, _
    Optional is_type As VBA.VbVarType, Optional force As VBA.VbVarType, _
    Optional datenessOf As Boolean = False, _
    Optional after As Date, Optional on_or_after As Date, _
    Optional before As Date, Optional on_or_before As Date, _
    Optional stringnessOf As Boolean = False, _
    Optional allow_blank As Boolean = False, Optional length As Integer = -1, _
    Optional min_length As Integer = -1, Optional max_length As Integer = -1, _
    Optional begins_with As String, Optional ends_with As String, _
    Optional contains As String, Optional matches As String, _
    Optional case_sensitive As Boolean = False, _
    Optional strict As Boolean = False _
    ) As Boolean
    
    Dim result As Boolean
    
    true_val = value
    test_val = true_val
    
    pname = name
    result = True
    If presenceOf Then
        result = result And presence
    End If
    If numericalityOf Then
        result = result And numericality(only_integer, allow_null, is_equal_to, greater_than, greater_than_or_equal_to, _
                                         less_than, less_than_or_equal_to, other_than, odd, even, is_type, force)
    End If
    If stringnessOf Then
        result = result And stringness(allow_null, allow_blank, length, min_length, max_length, begins_with, ends_with, contains, matches, case_sensitive, strict)
    End If
    If datenessOf Then
        result = result And dateness(allow_null, after, on_or_after, before, on_or_before, CDate(other_than))
    End If

    validates = result
End Function

Public Function custom_validation(valid As Boolean, Optional name As String, Optional message As String) As Boolean
    pname = name
    append_error valid, message
    custom_validation = valid
End Function


'[BEGIN]============================================numericality====================================================[BEGIN]
'Check a number field for validation purposes numericality returns true based on the supplied parameters
'val is a variant and will accept any object type using Error Handling to catch any issues and return False
'if val is a legitimate number the following parameters are available
'only_integer (Boolean) - checks to see if val is a vbInteger
'allow_null (Boolean) - this is the only option that does not require a number as it will return its own value if val is Null
'is_equal_to (Variant) - checks to see if test_val = is_equal_to
'greater_than (Variant) - will check to see if val is greater than this value and return (takes presidence over _or_equal_to)
'greater_than_or_equal_to (Variant) - same a greater than but with an equality check
'less_than (Variant) - will check to see if a value is less than this value and return (takes presidence over _or_equal_to)
'less_than_or_equal_to (Variant) - same as less than with an equality check
'other_than (Variant) - checks to see if val is anything other than this value
'odd (Boolean) - checks to see if val is odd if True
'even (Boolean) - checks to see if a val is even if True
'is_type (VBA.vbVarType) - checks to see if val is of a specific data_type
'force(VBA.VbVarType) - attempts to force test_val to a specified data-type prior to testing

'Options can be used in conjunction with each other
'example: numericality only_integer:=True,greater_than:=11,less_than:=22
'this method relies on the follwoing helper methods:
'   value_supplied(val As Variant) As Boolean
'   check_only_integer(val As Variant) As Boolean
'   check_is_type(val As Variant, vb_type As VBA.VbVarType) As Boolean
'   check_greater_than(val As Variant, greater As Variant, Optional equality As Boolean = False) As Boolean
'   check_less_than(val As Variant, lesser As Variant, Optional equality As Boolean = False) As Boolean
'   check_other_than(val As Variant, other As Variant) As Boolean
'   check_even_odd(val As Variant, Optional even As Boolean = True) As Boolean
'   check_numbers_only(val As Variant) As Boolean


Public Function numericality( _
    Optional only_integer As Boolean = False, _
    Optional allow_null As Boolean = False, _
    Optional is_equal_to As Variant, _
    Optional greater_than As Variant, _
    Optional greater_than_or_equal_to As Variant, _
    Optional less_than As Variant, _
    Optional less_than_or_equal_to As Variant, _
    Optional other_than As Variant, _
    Optional odd As Boolean = False, _
    Optional even As Boolean = False, _
    Optional is_type As VBA.VbVarType, _
    Optional force As VBA.VbVarType _
    ) As Boolean
    
On Error GoTo Failure
    Dim result As Boolean

    If IsNull(test_val) Or VarType(test_val) = vbNull Then
        result = append_error(allow_null, "cannot be null.")
    ElseIf Not (check_numbers_only()) Then
        result = append_error(check_numbers_only(), "can only conatin numbers.")
    Else
        result = True
        If force <> vbEmpty Then
            force_type (force)
        End If
        If only_integer Then
            result = result And append_error(check_only_integer(), "must be an integer.")
        End If
        If value_supplied(is_equal_to) Then
            result = append_error(test_val = is_equal_to, "must be equal to " & is_equal_to & ".")
        Else
            If value_supplied(greater_than) Then
                result = result And append_error(check_greater_than(greater_than), "must be greater than " & greater_than & ".")
            ElseIf value_supplied(greater_than_or_equal_to) Then
                result = result And append_error(check_greater_than(greater_than_or_equal_to, True), "must be greater than or equal to " & greater_than_or_equal_to & ".")
            End If
            If value_supplied(less_than) Then
                result = result And append_error(check_less_than(less_than), "must be less than " & less_than & ".")
            ElseIf value_supplied(less_than_or_equal_to) Then
                result = result And append_error(check_less_than(less_than_or_equal_to, True), "must be less than or equal to " & less_than_or_equal_to & ".")
            End If
        End If
        If value_supplied(other_than) Then
            result = result And append_error(check_other_than(other_than), "must be something other than " & other_than & ".")
        End If
        If odd Then
            result = result And append_error(check_even_odd(False), "must be odd.")
        End If
        If even Then
            result = result And append_error(check_even_odd(True), "must be even.")
        End If
        If is_type <> vbEmpty Then
            result = result And append_error(check_is_type(is_type), "must be a " & is_type & " data type.")
        End If
    End If
    
    numericality = result
Exit_numericality:
    valid = numericality
    Exit Function
Failure:
    numericality = False
    GoTo Exit_numericality
End Function


'check_only_integer - checks to see if the value is vbInteger or vbLong data type
Private Function check_only_integer() As Boolean
    check_only_integer = check_is_type(vbInteger) Or check_is_type(vbLong)
End Function
'check_is_type(vb_type VBA.VbVarType) - checks to see if the value is of the same data type as vb_type
Private Function check_is_type(vb_type As VBA.VbVarType) As Boolean
    check_is_type = VarType(test_val) = vb_type
End Function
'Force test_val to a specific data type before validating
Private Sub force_type(conv As VBA.VbVarType)
    Select Case conv
        Case vbInteger
            test_val = CInt(test_val)
        Case vbLong
            test_val = CLng(test_val)
        Case vbSingle
            test_val = CSng(test_val)
        Case vbDouble
            test_val = CDbl(test_val)
        Case vbCurrency
            test_val = CCur(test_val)
        Case vbDecimal
            test_val = CDec(test_val)
    End Select
        
End Sub


'check_even_odd(Optional even As Boolean = True) - checks to see if a value is even or odd based on even Boolean
'if even = False then odd
Private Function check_even_odd(Optional even As Boolean = True) As Boolean
    If even Then
        check_even_odd = test_val Mod 2 = 0
    Else
        check_even_odd = test_val Mod 2 <> 0
    End If
End Function

'check_numbers_only - this is a fix for vba's IsNumeric function which acknowledges "123FD23" as being numeric
'this checks to make sure a number only contains digits 0-9, a period(.), or comma(,)
Private Function check_numbers_only() As Boolean
        With CreateObject("vbscript.Regexp")
            .pattern = "/[^0-9.,]/"
            check_numbers_only = Not (.test(test_val))
        End With
End Function
'[END]===========================================numericality=====================================================[END]


'[BEGIN]=============================================presence====================================================[BEGIN]
'presenceOf - checks to make sure the validation element is not null and have a length of greater than 0

Public Function presence() As Boolean
    presence = append_error(Len(IIf(IsNull(test_val) Or test_val = vbNull, "", test_val)) > 0, "must have a value.")
End Function
'[END]=============================================presence====================================================[END]



'[BEGIN]==============================================dateness=====================================================[BEGIN]
'Check a date field for validation purposes dateness returns true based on the supplied parameters
'test_val is a variant and will accept any object type using Error Handling to catch any issues and return False
'if test_val is a legitimate date the following parameters are available
'allow_null (Boolean) - this is the only option that does not require a date as it will return its own value if test_val is Null
'after (Date) - will check to see if test_val is after this value and return (takes presidence over on_or_)
'on_or_after (Date) - same a after but with an equality check
'before(Date) - will check to see if a value is before this value and return (takes presidence over on_or_)
'on_or_before (Date) - same as before with an equality check
'other_than (Date) - checks to see if test_val is anything other than this value

'Options can be used in conjunction with each other
'example: dateness before:=#2/20/2013#,on_or_after:=#1/1/2013
'this method relies on the follwoing helper methods:
'   value_supplied(val As Variant) As Boolean
'   check_greater_than(val As Variant, greater As Variant, Optional equality As Boolean = False) As Boolean
'   check_less_than(val As Variant, lesser As Variant, Optional equality As Boolean = False) As Boolean
'   check_other_than(val As Variant, other As Variant) As Boolean

Function dateness( _
Optional allow_null As Boolean = False, _
Optional after As Date, _
Optional on_or_after As Date, _
Optional before As Date, _
Optional on_or_before As Date, _
Optional other_than As Date _
) As Boolean
    'Check if val is a Date additional parameters for a range it must fall into
On Error GoTo Failure
    Dim result As Boolean
    If IsNull(test_val) Or test_val = vbNull Then
        result = append_error(allow_null, "cannot be null.")
    ElseIf Not (IsDate(test_val)) Then
        result = append_error(check_numbers_only(), "must be a Date.")
    Else
        result = True
        If date_supplied(after) Then
            result = result And append_error(check_greater_than(after), "must be after " & after & ".")
        ElseIf date_supplied(on_or_after) Then
            result = result And append_error(check_greater_than(on_or_after, True), "must be after or on " & on_or_after & ".")
        End If
        If date_supplied(before) Then
            result = result And append_error(check_less_than(before), "must be before " & before & ".")
        ElseIf date_supplied(on_or_before) Then
            result = result And append_error(check_less_than(on_or_before, True), "must be before or on " & on_or_before & ".")
        End If
        If date_supplied(other_than) Then
            result = result And append_error(check_other_than(other_than), "must be something other than " & other_than & ".")
        End If
    End If
    
    dateness = result
Exit_dateness:
    Exit Function
Failure:
    dateness = False
    GoTo Exit_dateness
End Function
'[END]============================================dateness=======================================================[END]

'[BEGIN]=========================================stringness======================================================[BEGIN]
'Check a string for validation purposes stringness returns true based on the supplied parameters
'test_val is a variant and will accept any object type using Error Handling to catch any issues and return False
'if test_val is a legitimate string the following parameters are available
'allow_null (Boolean) - this is the only option that does not require a string as it will return its own value if test_val is Null
'allow_blank (Boolean) - this is similar to allow_null only it will fail if the sting is null but pass if it is an empty string ""
'length(Integer) - check if the string is a specified length
'min_length(Integer) - check if the length of a string is greater than or equal to min length
'max_length(Integer) - check is a string is short than or equal to max length
'begins_with (String) - Check is a string begins with a specified string
'ends_with (String) - Check is a string ends with a specified string (can be used with case_sensitive)
'contains (String) - Check is a string contains a specified string (can be used with case_sensitive)
'matches (String) - Check is a string matches a given regex pattern (case sensitive has no effect on this method)
'case_sensitive (Boolean) - to be used in conjunction with begins_with, ends_with, and contains
'strict (Boolean) - if strict it validates test_val as is otherwise it attempts to make it a string before testing


'Options can be used in conjunction with each other
'example: stringness allow_blank:=True,min_length:=0,max_length:=23,contains:="word"
'this method relies on the following helper methods:
'   value_supplied(val As Variant) As Boolean
'   string_supplied(val As String) AS Boolean
'   check_length(length As Integer, Optional min As Boolean = False, Optional max As Boolean = False) As Boolean
'   check_pattern(pattern As String) As Boolean

Public Function stringness( _
Optional allow_null As Boolean = False, _
Optional allow_blank As Boolean = False, _
Optional length As Integer = -1, _
Optional min_length As Integer = -1, _
Optional max_length As Integer = -1, _
Optional begins_with As String, _
Optional ends_with As String, _
Optional contains As String, _
Optional matches As String, _
Optional case_sensitive As Boolean = False, _
Optional strict As Boolean = False _
) As Boolean
On Error GoTo Failure
    Dim result As Boolean
    Dim match_case As String
    
    If Not (strict) Then
        test_val = CStr(true_val)
    End If
    
    If case_sensitive Then
        match_case = "(Case Sensitive)"
    Else
        match_case = "(Non-Case Sensitive)"
    End If
    If IsNull(test_val) Or test_val = vbNull Then
        result = append_error(allow_null, "cannot be null.")
    ElseIf test_val = "" Then
        result = append_error(allow_blank, "cannot be blank")
    ElseIf VarType(test_val) <> vbString And strict Then
        result = append_error(False, "must be a string.")
    Else
        result = True
        If length >= 0 Then
            result = result And append_error(check_length(length), "must be " & length & " character(s) long.")
        ElseIf min_length >= 0 Or max_length > 0 Then
            If min_length >= 0 Then
                 result = result And append_error(check_length(length, min:=True), "must be atleast " & min_length & " character(s) long.")
            End If
            If max_length > 0 Then
                 result = result And append_error(check_length(length, max:=True), "must be no more than " & max_length & " character(s) long.")
            End If
        End If
        If string_supplied(begins_with) Then
            result = result And append_error(check_sub_string(begins_with, begins:=True, case_sensitive:=case_sensitive), "must begin with " & begins_with & "." & match_case)
        End If
        If string_supplied(ends_with) Then
            result = result And append_error(check_sub_string(ends_with, ends:=True, case_sensitive:=case_sensitive), "must end with " & ends_with & "." & match_case)
        End If
        If string_supplied(contains) Then
            result = result And append_error(check_sub_string(contains, contains:=True, case_sensitive:=case_sensitive), "must contain " & contains & "." & match_case)
        End If
        If string_supplied(matches) Then
            result = result And append_error(check_pattern(matches), "must match " & matches & " pattern.")
        End If
        
    End If
    stringness = result
Exit_Stringness:
    Exit Function
Failure:
    stringness = False
    GoTo Exit_Stringness
End Function

'Checks to see if a String was supplied since default value is empty String
Private Function string_supplied(val As String) As Boolean
    string_supplied = val <> ""
End Function

'Checks the length of a string based on a specified length, min length or max length
Private Function check_length(length As Integer, Optional min As Boolean = False, Optional max As Boolean = False) As Boolean
    If min Then
        check_length = Len(test_val) >= length
    ElseIf max Then
        check_length = Len(test_val) <= length
    Else
        check_length = Len(test_val) = length
    End If
End Function

'Checks a string for a substring based on begins, ends, or contains and case sensitivity
Private Function check_sub_string(sub_string As String, _
Optional begins As Boolean = False, _
Optional ends As Boolean = False, _
Optional contains As Boolean = False, _
Optional case_sensitive As Boolean = False) As Boolean

    Dim t_string As String
    If Not (case_sensitive) Then
        t_string = UCase(test_val)
        sub_string = UCase(sub_string)
    Else
        t_string = test_val
    End If
    If begins Then
        check_sub_string = StrComp(Left(t_string, Len(sub_string)), sub_string, vbBinaryCompare) = 0
    ElseIf ends Then
        check_sub_string = StrComp(Right(t_string, Len(sub_string)), sub_string, vbBinaryCompare) = 0
    ElseIf contains Then
        check_sub_string = InStr(1, t_string, sub_string, vbBinaryCompare) > 0
    End If

End Function
'checks if a string matches a given pattern
Private Function check_pattern(pattern As String) As Boolean
        With CreateObject("vbscript.Regexp")
            .pattern = pattern
            check_pattern = .test(test_val)
        End With
End Function
'[END]==============================================================stringness==============================================================[END]


'===================================================Shared Private Methods for Validation Functions============================================
'append_error builds a collection of errors with a message and a field key if one is set
Private Function append_error(valid As Boolean, error As String) As Boolean
    If Not (valid) Then
        Dim new_error As New ValidatorError
        new_error.field = pname
        new_error.error = error
        error_list.Add new_error
    End If
    append_error = valid
End Function
'value_supplied checks to see if the variable was set in the call
Private Function value_supplied(val As Variant) As Boolean
    value_supplied = Not IsMissing(val)
End Function
Private Function date_supplied(val As Date) As Boolean
    date_supplied = val <> 0
End Function
'check_greater_than(greater As Variant, Optional equality As Boolean = False) - checks to see if the value is greater than the value specified by greater
'if equality is true then it adds a check for equality
Private Function check_greater_than(greater As Variant, Optional equality As Boolean = False) As Boolean
    If equality Then
        check_greater_than = test_val >= greater
    Else
        check_greater_than = test_val > greater
    End If
End Function

'check_less_than(lesser As Variant, Optional equality As Boolean = False) - checks to see if the value is less than the value specified by lesser
'if equality is true then it adds a check for equality
Private Function check_less_than(lesser As Variant, Optional equality As Boolean = False) As Boolean
    If equality Then
        check_less_than = test_val <= lesser
    Else
        check_less_than = test_val < lesser
    End If
End Function

'check_other_than(other As Variant) - checks to see if the value is anything other than the value specified by other
Private Function check_other_than(other As Variant) As Boolean
    check_other_than = test_val <> other
End Function

'Checks if an item exists in a colleciton by catching the non existent Error
Private Function in_collection(coll As Collection, key As String) As Boolean
On Error Resume Next
    Err.Clear
    coll.Item (key)
    in_collection = Not (CLng(Err.Number) = 5)
End Function
