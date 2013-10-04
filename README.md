VBA_Rails_Style
===============

Rails Style Validation for VBA

Designed By engineersmnky 10/4/2013

Validator acts as a validation engine for VBA Applications

Validators formatting was inspired by Rails validations and uses as similar a foramtting style as possible

It must be used in conjunction with ValidatorError which will hold the name and error messgae for each object that fails validation
 
Validator includes multiple methods for retreiving these error including
 
errors - Which returns the full collection of Errors
 
error_messages - which returns only the messages for objects that failed validation
 
error_keys - returns only the name for objects that failed validation
 
uniq_keys - returns a unique collection of errors based on name so only the first failed validation for a named object will be included in this collection

Usage:

       Dim v As New Validator
       With v
            .validates 123, "Number", numericalityOf:=True, only_integer:=True, greater_than_or_equal_to:=11
            .validates "String", "String", stringnessOf:=True, length:=6, contains:="ing", begins_with:="S"
            .validates "12345","Number2",numericalityOf:=True, greater_than:=12344, force:=vbInteger, stringnessOf:=True, min_length:=3, begins_with:="1"
       End With


       If v.is_valid Then
           Do Something
       Else
           Do something Else
       End If
       
 Alternate Usage:
 
       Dim v As New Validator
       v.value = 123
       v.name = "Number"


       ^^^ is the same as v.validates 123,"Number"

       v.numericality only_integer:=True, greater_than_or_equal_to:=11
       v.stringness length:=6, contains:="ing", begins_with:="S"
       v.dateness 
       
 Each validates call returns a Boolean so you can use this in conditionals as well.
 
 
       If v.stringness(length:=2) then
           Msgbox "String is 2 Characters Long."
       Else
            Msgbox "String is not 2 Characters Long."
       End If


 Custom Validation:
   If you find the builtin methods are not enough to handle your validation it comes with a .custom_validation function which allows you to 
       specify a Boolean statement with an optional name and error_message


 `.validates` allows you to set a values and an optional name for each validation
 Options are:
 
       Shared:
            allow_null (Boolean) - returns true for validation even if value is null * Does not apply to presence
            other_than (Variant) - returns true if the value is something other than what is specified * Shared By numericality and dateness
       presenceOf: 
           returns true is object has a value
       numericalityOf options:(Default=False)
           only_integer (Boolean) - checks to see if val is a vbInteger
           allow_null (Boolean) - this is the only option that does not require a number as it will return its own value if val is Null
           is_equal_to (Variant) - checks to see if test_val = is_equal_to
           greater_than (Variant) - will check to see if val is greater than this value and return (takes presidence over _or_equal_to)
           greater_than_or_equal_to (Variant) - same a greater than but with an equality check
           less_than (Variant) - will check to see if a value is less than this value and return (takes presidence over _or_equal_to)
           less_than_or_equal_to (Variant) - same as less than with an equality check
           odd (Boolean) - checks to see if val is odd if True
           even (Boolean) - checks to see if a val is even if True
           is_type (VBA.vbVarType) - checks to see if val is of a specific data_type
           force(VBA.VbVarType) - attempts to force test_val to a specified data-type prior to testing 
                     * This value can be retrieved afterwards with .test_value (unless used with stringness strict:=False(Default))

       stringnessOf options: (Default=False)
           allow_blank (Boolean) - this is similar to allow_null only it will fail if the sting is null but pass if it is an empty string ""
           length(Integer) - check if the string is a specified length
           min_length(Integer) - check if the length of a string is greater than or equal to min length
           max_length(Integer) - check is a string is short than or equal to max length
           begins_with (String) - Check is a string begins with a specified string
           ends_with (String) - Check is a string ends with a specified string (can be used with case_sensitive)
           contains (String) - Check is a string contains a specified string (can be used with case_sensitive)
           matches (String) - Check is a string matches a given regex pattern (case sensitive has no effect on this method)
           case_sensitive (Boolean) - to be used in conjunction with begins_with, ends_with, and contains
           strict (Boolean) - if strict it validates test_val as is otherwise it attempts to make it a string before testing  
                   *This value can be retrieved afterwards with .test_value (overrides numericality force)

       datenessOf options:(Default=False)
           allow_null (Boolean) - this is the only option that does not require a date as it will return its own value if test_val is Null
           after (Date) - will check to see if test_val is after this value and return (takes presidence over on_or_)
           on_or_after (Date) - same a after but with an equality check
           before(Date) - will check to see if a value is before this value and return (takes presidence over on_or_)
           on_or_before (Date) - same as before with an equality check

     example: validates 123.45, "Number", [Options]
