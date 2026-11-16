VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRsize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#####################################################################################
'#
'#  clsRsize - Manage resize events on forms with multiple GUI components
'#  v2.00    - Now handles 50% resizes on items
'#  v4.00    - Now handles resize rules defined in controls Tag property
'#
'#  Form Global - Dim ReSize As clsRsize
'#  Form Load   - Set ReSize = New clsRsize
'#              V1/V2 compat
'#              - ReSize.HandleForm = Me
'#              - ReSize.Attach GUIOBJECT, False, False, True, True
'#           OR - Resize.VAttach GUIOBJECT, 0/50/100, 0/50/100, 0/50/100, 0/50/100
'#              V4 compat
'#              - Resize.HandleFormEx = Me  (where GUIOBJECT.Tag = resize(FFTT))
'#           OR - Resize.HandleFormEx = Me  (where GUIOBJECT.Tag = resize(nnn,nnn,nnn,nnn))
'#              - Resize.Ready = True
'#
'#              - Slider controls - hslider(picSplitter,text1;text2;text5,text3;text4)
'#  Form Resize - ReSize.ReSize
'#
'#####################################################################################

Dim Attached As Collection
Dim AttachForm As Form
Dim edge(4) As Integer
Dim iReady As Boolean

Dim Sliders As Collection
Dim mbMoving As Boolean
Dim MoveDirection As String
Dim MovingSlider As Collection

Public Function Attach(GUIobject As Object, VLeft As Variant, VRight As Variant, VTop As Variant, VBottom As Variant)
Dim AttachedObj As clsAtObj
Dim bSticky(4) As Boolean
Set AttachedObj = New clsAtObj

For i = 1 To 4
bSticky(4) = False
Next

If VLeft = True Then bSticky(1) = VLeft
If VRight = True Then bSticky(2) = VRight
If VTop = True Then bSticky(3) = VTop
If VBottom = True Then bSticky(4) = VBottom

'AttachedObj.SetGUI GUIobject, bleft, bright, btop, bbottom
AttachedObj.SetGUI GUIobject, bSticky(1), bSticky(2), bSticky(3), bSticky(4)

Attached.Add AttachedObj

End Function

Public Function VAttach(GUIobject As Object, VLeft As Variant, VRight As Variant, VTop As Variant, VBottom As Variant)
Dim AttachedObj As clsAtObj
Dim bSticky(4) As Integer
Set AttachedObj = New clsAtObj

For i = 1 To 4
bSticky(4) = -1 '-1 = false, 0=50%, 1=true
Next

bSticky(1) = VLeft
bSticky(2) = VRight
bSticky(3) = VTop
bSticky(4) = VBottom

'AttachedObj.SetGUI GUIobject, bleft, bright, btop, bbottom
AttachedObj.VSetGUI GUIobject, bSticky(1), bSticky(2), bSticky(3), bSticky(4)

Attached.Add AttachedObj

End Function

Private Sub Class_Initialize()
Set Attached = New Collection
iCounter = 0
End Sub


Public Property Get HandleForm() As Form
Set HandleForm = AttachForm

End Property

Public Property Let HandleForm(ByVal vNewValue As Form)
Set AttachForm = vNewValue
edge(1) = vNewValue.Left
edge(2) = vNewValue.Width
edge(3) = vNewValue.Top
edge(4) = vNewValue.Height
End Property

Public Property Let HandleFormEx(ByVal vNewValue As Form)

    Dim objElement As Control
    Dim sTag As String
    Dim sSliderMode As String
    Dim sSliderPic As String
    Dim sSliderTLList As String
    Dim sSliderBRList As String
    Dim colSliderDefinition As Collection
    Dim colTL As Collection
    Dim colBR As Collection
    Dim sItem As String
    Dim iPos As Integer
    
    Set AttachForm = vNewValue
    edge(1) = vNewValue.Left
    edge(2) = vNewValue.Width
    edge(3) = vNewValue.Top
    edge(4) = vNewValue.Height
    
    Set Sliders = New Collection

    For Each objElement In AttachForm.Controls
        sTag = ""
        On Error Resume Next
        sTag = objElement.Tag
        iPos = InStr(LCase(sTag), "resize(")
        If iPos > 0 Then
            sTag = UCase$(Mid$(sTag, iPos + 7))
            iPos = InStr(sTag, ")")
            If iPos > 0 Then
                sTag = Left$(sTag, iPos - 1)
            End If
            
            'boolean form - resize(TFTF)
            If Len(sTag) = 4 Then
                Me.Attach objElement, _
                          Left$(sTag, 1) = "T", Mid$(sTag, 2, 1) = "T", _
                          Mid$(sTag, 3, 1) = "T", Right$(sTag, 1) = "T"
            End If
        
            'variable form - resize(000,100,050,100)
            If Len(sTag) = 15 Then
                Me.VAttach objElement, _
                           CInt(Left$(sTag, 3)), CInt(Mid$(sTag, 5, 3)), _
                           CInt(Mid$(sTag, 9, 3)), CInt(Right$(sTag, 3))
            End If
        End If
        
        'V4.0 - Handle slider processing
        sTag = objElement.Tag
        iPos = InStr(LCase(sTag), "slider(")
        If iPos > 0 Then
            If iPos > 1 Then
                sSliderMode = UCase$(Mid$(sTag, iPos - 1, 1))
            Else
                sSliderMode = "H"
            End If
            sTag = Mid$(sTag, iPos + 7)
            iPos = InStr(sTag, ")")
            If iPos > 0 Then
                sTag = Left$(sTag, iPos - 1)
            End If
            
            iPos = InStr(sTag, ",")
            If iPos > 0 Then
                sSliderPic = Left$(sTag, iPos - 1)
                sTag = Mid$(sTag, iPos + 1)
                iPos = InStr(sTag, ",")
                If iPos > 0 Then
                    sSliderTLList = Left$(sTag, iPos - 1)
                    sSliderBRList = Mid$(sTag, iPos + 1)
                        
                    Set colSliderDefinition = New Collection
                    
                    Sliders.Add colSliderDefinition, objElement.Name
                    
                    colSliderDefinition.Add sSliderMode
                    colSliderDefinition.Add AttachForm.Controls(sSliderPic)
                    Set colTL = New Collection
                    Set colBR = New Collection
                    
                    colSliderDefinition.Add colTL
                    colSliderDefinition.Add colBR
                    
                    While sSliderTLList <> ""
                        iPos = InStr(sSliderTLList, ";")
                        If iPos = 0 Then
                            sValue = sSliderTLList
                            sSliderTLList = ""
                        Else
                            sValue = Left$(sSliderTLList, iPos - 1)
                            sSliderTLList = Mid$(sSliderTLList, iPos + 1)
                        End If
                        colTL.Add AttachForm.Controls(sValue)
                    Wend
                
                    While sSliderBRList <> ""
                        iPos = InStr(sSliderBRList, ";")
                        If iPos = 0 Then
                            sValue = sSliderBRList
                            sSliderBRList = ""
                        Else
                            sValue = Left$(sSliderBRList, iPos - 1)
                            sSliderBRList = Mid$(sSliderBRList, iPos + 1)
                        End If
                        colBR.Add AttachForm.Controls(sValue)
                    Wend
                        
                End If
            End If
        End If
    Next
    
    'hslider(picSplitter,text1;text2;text5,text3;text4)
    
    'collection(keyed by imgslider name) contains collection objects
    '  Item 1 - 'H' or 'V'
    '  Item 2 - picture control
    '  Item 3 - collection of T/L controls
    '  Item 4 - collection of B/R controls
    
End Property

' 0 = mousedown
' 1 = mousemove
' 2 = mmouseup
Public Sub Drag(Mode As Integer, SliderControl As Image, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single
    Dim sglSplitLimit As Single
    Dim objControl As Control
    
    On Error Resume Next
    
    sglSplitLimit = 500

    Select Case Mode
        Case 0
        Set MovingSlider = Sliders(SliderControl.Name)
        MoveDirection = UCase$(CStr(MovingSlider(1)))
        
        With SliderControl
            If MoveDirection = "H" Then
                MovingSlider(2).Move .Left, .Top, .Width \ 2, .Height - 20
            Else
                MovingSlider(2).Move .Left, .Top, .Width - 20, .Height / 2
            End If
        End With
        MovingSlider(2).Visible = True
        mbMoving = True

        Case 1
        If mbMoving Then
            If MoveDirection = "H" Then
                sglPos = X + SliderControl.Left
                If sglPos < sglSplitLimit Then
                    MovingSlider(2).Left = sglSplitLimit
                ElseIf sglPos > AttachForm.Width - sglSplitLimit Then
                    MovingSlider(2).Left = AttachForm.Width - sglSplitLimit
                Else
                    MovingSlider(2).Left = sglPos
                End If
            Else
                sglPos = Y + SliderControl.Top
                If sglPos < sglSplitLimit Then
                    MovingSlider(2).Top = sglSplitLimit
                ElseIf sglPos > AttachForm.Height - sglSplitLimit Then
                    MovingSlider(2).Top = AttachForm.Height - sglSplitLimit
                Else
                    MovingSlider(2).Top = sglPos
                End If
            End If
        End If

        Case 2
        'MovingSlider(2).Left has the new position, SliderControl.Left has original
        If MoveDirection = "H" Then
            sglPos = MovingSlider(2).Left - SliderControl.Left
            MovingSlider(2).Visible = False
            
            For Each objControl In MovingSlider(3)
                objControl.Width = objControl.Width + sglPos
            Next
            
            For Each objControl In MovingSlider(4)
                objControl.Move objControl.Left + sglPos, objControl.Top, objControl.Width - sglPos
            Next
            
            SliderControl.Move MovingSlider(2).Left
        Else
            sglPos = MovingSlider(2).Top - SliderControl.Top
            MovingSlider(2).Visible = False
            
            For Each objControl In MovingSlider(3)
                objControl.Height = objControl.Height + sglPos
            Next
            
            For Each objControl In MovingSlider(4)
                objControl.Move objControl.Left, objControl.Top + sglPos, objControl.Width, objControl.Height - sglPos
            Next
            
            SliderControl.Move SliderControl.Left, MovingSlider(2).Top
        End If
        
        mbMoving = False
        Set MovingSlider = Nothing
        
        ' reset new sizes
        Set Attached = New Collection
        iCounter = 0
        Me.HandleFormEx = AttachForm
    
    End Select
End Sub

Public Function ReSize()

If Not Ready Then Exit Function
If AttachForm.WindowState = 1 Then Exit Function

  Dim i As Integer
  Dim a As New clsAtObj
  Dim newsize As Integer

  AttachForm.Cls

  For i = 1 To Attached.Count
  
    Set a = Attached.Item(i)
    
    If Not a.Varies Then
    
        If a.Bottom Then
        On Error GoTo height_error
            If a.Top Then
                newsize = a.OrgHeight + AttachForm.Height - edge(4)
                a.Attached.Height = newsize
            Else
                a.Attached.Top = a.OrgTop - edge(4) + AttachForm.Height
            End If
        End If
        
        If a.Right Then
            On Error GoTo width_error
            If a.Left Then
                newsize = a.OrgWidth - edge(2) + AttachForm.Width
                a.Attached.Width = newsize
            Else
                a.Attached.Left = a.OrgLeft - edge(2) + AttachForm.Width
            End If
        End If
    
    Else
    
        If a.VBottom = 100 Then
        On Error GoTo height_error
            If a.VTop = 100 Then
                ' Stretch with top/bottom 100%
                newsize = a.OrgHeight + AttachForm.Height - edge(4)
                a.Attached.Height = newsize
            Else
                If a.VTop = 0 Then
                    ' Move downwards with bottom
                    a.Attached.Top = a.OrgTop - edge(4) + AttachForm.Height
                Else
                    ' Head downwards by amount (height increases by %)
                    'newsize = a.OrgHeight + ((AttachForm.Height - edge(4)) / 2)
                    newsize = a.OrgHeight + ((AttachForm.Height - edge(4)) * (a.VTop / 100))
                    a.Attached.Height = newsize
                    'a.Attached.Top = a.OrgTop + ((AttachForm.Height - edge(4)) / 2)
                    a.Attached.Top = a.OrgTop + ((AttachForm.Height - edge(4)) * ((100 - a.VTop) / 100))
                End If
            End If
        End If
                                
        'If a.VBottom = 50 Then
        If a.VBottom > 0 And a.VBottom < 100 Then
            On Error GoTo width_error
            'If a.VTop = 50 Then
            ''If a.VTop = (100 - a.VBottom) Then
            '    ' Slide downwards at %speed
            '    'a.Attached.Top = a.OrgTop + ((AttachForm.Height - edge(4)) / 2)
            ''    a.Attached.Top = a.OrgTop + ((AttachForm.Height - edge(4)) * (a.VTop / 100))
            ''Else
                ' Stretch downwards at half speed
                If a.VTop = 100 Then
                    'newsize = a.OrgHeight + ((AttachForm.Height - edge(4)) / 2)
                    newsize = a.OrgHeight + ((AttachForm.Height - edge(4)) * ((a.VBottom) / 100))
                    a.Attached.Height = newsize
                Else '0 - need to work out differince in size
                    'Shrinking because top is moving faster than bottom!
                    'a.Attached.Top = a.OrgTop + AttachForm.Height - edge(4)
                    a.Attached.Top = a.OrgTop + ((AttachForm.Height - edge(4)) * ((100 - a.VTop) / 100))
                    'newsize = a.OrgHeight - ((AttachForm.Height - edge(4)) / 2)
                    newsize = a.OrgHeight - ((AttachForm.Height - edge(4)) * ((100 - a.VBottom) - a.VTop) / 100)
                    a.Attached.Height = newsize
                End If
            ''End If
        End If
                                
        If a.VRight = 100 Then
            On Error GoTo width_error
            If a.VLeft = 100 Then
                newsize = a.OrgWidth - edge(2) + AttachForm.Width
                a.Attached.Width = newsize
            Else
                If a.VLeft = 0 Then
                    a.Attached.Left = a.OrgLeft - edge(2) + AttachForm.Width
                Else '50
                    'newsize = a.OrgWidth + ((AttachForm.Width - edge(2)) / 2)
                    newsize = a.OrgWidth + ((AttachForm.Width - edge(2)) * (a.VLeft / 100))
                    a.Attached.Width = newsize
                    'a.Attached.Left = a.OrgLeft + ((AttachForm.Width - edge(2)) / 2)
                    a.Attached.Left = a.OrgLeft + ((AttachForm.Width - edge(2)) * ((100 - a.VLeft) / 100))
                End If
            End If
        End If
        
        'If a.VRight = 50 Then
        If a.VRight > 0 And a.VRight < 100 Then
            On Error GoTo width_error
            'If a.VLeft = 50 Then
            '    a.Attached.Left = a.OrgLeft + ((AttachForm.Width - edge(2)) / 2)
            'Else
                If a.VLeft = 100 Then
                    'newsize = a.OrgWidth + ((AttachForm.Width - edge(2)) / 2)
                    newsize = a.OrgWidth + ((AttachForm.Width - edge(2)) * ((a.VRight) / 100))
                    a.Attached.Width = newsize
                Else '0
                    'Shrinking because left is moving faster than right!
                    'a.Attached.Left = a.OrgLeft + AttachForm.Width - edge(2)
                    a.Attached.Left = a.OrgLeft + ((AttachForm.Width - edge(2)) * ((100 - a.VLeft) / 100))
                    'newsize = a.OrgWidth - ((AttachForm.Width - edge(2)) / 2)
                    newsize = a.OrgWidth - ((AttachForm.Width - edge(2)) * ((100 - a.VRight) - a.VLeft) / 100)
                    a.Attached.Width = newsize
                End If
            'End If
        End If
        
    
    End If
    
    Set a = Nothing

  Next i
  Exit Function

width_error:
  AttachForm.Width = edge(2)
  Exit Function
height_error:
  AttachForm.Height = edge(4)
  
End Function

Public Property Get Ready() As Boolean
Ready = iReady
End Property

Public Property Let Ready(ByVal vNewValue As Boolean)
iReady = vNewValue
End Property

