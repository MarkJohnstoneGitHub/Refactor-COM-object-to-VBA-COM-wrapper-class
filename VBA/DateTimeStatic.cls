VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTimeStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("DotNet.System")
'@PredeclaredId

Option Explicit

Private Type TDateTimeStatic
    MaxValue As DateTimeWrapper
    MinValue As DateTimeWrapper
End Type

Private this As TDateTimeStatic

Private Sub Class_Initialize()
    Set this.MaxValue = New DateTimeWrapper
    Set this.MaxValue.ComObject = DateTimeWrapper.ComObject.MaxValue
    
    Set this.MinValue = New DateTimeWrapper
    Set this.MinValue.ComObject = DateTimeWrapper.ComObject.MinValue
End Sub

Public Property Get MaxValue() As DateTimeWrapper
    Set MaxValue = this.MaxValue
End Property

Public Property Get MinValue() As DateTimeWrapper
    Set MinValue = this.MinValue
End Property


