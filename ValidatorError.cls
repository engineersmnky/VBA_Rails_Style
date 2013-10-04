VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidatorError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private pfield As String
Private perror As String
'returns the field key for the error
Public Property Get field() As String
    field = pfield
End Property
'sets the field key for the error
Public Property Let field(value As String)
    pfield = value
End Property
'returns the message for the error
Public Property Get error_message() As String
    error_message = perror
End Property
'sets message for the error
Public Property Let error(value As String)
    perror = value
End Property
'returns the field and message as a full string
Public Property Get full_message() As String
    full_message = pfield & " " & perror
End Property



