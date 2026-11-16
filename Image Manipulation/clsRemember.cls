VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRemember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Attached As Form
Public TagValue As String
Public RememberPosition As Boolean

Public Sub SaveFormState()
    Dim objElement As Control
    Dim sValue As String
    For Each objElement In GetElements()
    
        If TypeOf objElement Is TextBox Then sValue = objElement.Text
        If TypeOf objElement Is ComboBox Then sValue = objElement.Text
        If TypeOf objElement Is ListBox Then sValue = objElement.Text
        
        If TypeOf objElement Is CheckBox Then sValue = objElement.Value
        If TypeOf objElement Is OptionButton Then sValue = objElement.Value
        If TypeOf objElement Is HScrollBar Then sValue = objElement.Value
        If TypeOf objElement Is VScrollBar Then sValue = objElement.Value
        
        SaveSetting App.Title, Attached.Name, ObjToName(objElement), sValue
    Next
    
    If RememberPosition Then
        SaveSetting App.Title, Attached.Name, "Left", Attached.Left
        SaveSetting App.Title, Attached.Name, "Top", Attached.Top
        SaveSetting App.Title, Attached.Name, "Width", Attached.Width
        SaveSetting App.Title, Attached.Name, "Height", Attached.Height
    End If

End Sub

Public Sub GetFormState()
    Dim objElement As Control
    Dim sValue As String
    For Each objElement In GetElements()
        
        If TypeOf objElement Is TextBox Then sValue = objElement.Text
        If TypeOf objElement Is ComboBox Then sValue = objElement.Text
        If TypeOf objElement Is ListBox Then sValue = objElement.Text
        
        If TypeOf objElement Is CheckBox Then sValue = objElement.Value
        If TypeOf objElement Is OptionButton Then sValue = objElement.Value
        If TypeOf objElement Is HScrollBar Then sValue = objElement.Value
        If TypeOf objElement Is VScrollBar Then sValue = objElement.Value
        
        sValue = GetSetting(App.Title, Attached.Name, ObjToName(objElement), sValue)
        
        If TypeOf objElement Is TextBox Then objElement.Text = sValue
        If TypeOf objElement Is ComboBox Then objElement.Text = sValue
        If TypeOf objElement Is ListBox Then objElement.Text = sValue
        
        If TypeOf objElement Is CheckBox Then objElement.Value = sValue
        If TypeOf objElement Is OptionButton Then objElement.Value = sValue
        If TypeOf objElement Is HScrollBar Then objElement.Value = sValue
        If TypeOf objElement Is VScrollBar Then objElement.Value = sValue
    Next

    If RememberPosition Then
        Attached.Left = GetSetting(App.Title, Attached.Name, "Left", Attached.Left)
        Attached.Top = GetSetting(App.Title, Attached.Name, "Top", Attached.Top)
        Attached.Width = GetSetting(App.Title, Attached.Name, "Width", Attached.Width)
        Attached.Height = GetSetting(App.Title, Attached.Name, "Height", Attached.Height)
    End If

End Sub

Private Function GetElements() As Collection

    Dim colElements As Collection
    Dim objElement As Control
    Dim bConsider As Boolean
    Dim iTab As Integer
    Dim sName As String
    Dim iIndex As Integer
    
    Set colElements = New Collection
    
    For iTab = 0 To Attached.Controls.Count - 1
    
        Set objElement = Attached.Controls(iTab)
    
        If Left$(objElement.Tag, Len(TagValue)) = TagValue Then
            colElements.Add objElement, ObjToName(objElement)
        End If
    Next
    
    Set GetElements = colElements
    Set colElements = Nothing
End Function

Private Function ObjToName(objControl As Control) As String
    Dim sDummy As String
    On Error Resume Next
    sDummy = CStr(objControl.Index)
    If Err.Number = 343 Then
        ObjToName = objControl.Name
    Else
        ObjToName = objControl.Name & "(" & objControl.Index & ")"
    End If
End Function


Private Sub Class_Initialize()
    TagValue = "S"
    RememberPosition = True
End Sub
