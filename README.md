VBA_Rails_Style
===============

Rails Style Validation for VBA

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
