VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Distort"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBR 
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   10
      Tag             =   "resize(050,050,000,100)"
      Text            =   "200"
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox txtBR 
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   9
      Tag             =   "resize(TFFT)"
      Text            =   "200"
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox txtBL 
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   8
      Tag             =   "resize(050,050,000,100)"
      Text            =   "200"
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox txtBL 
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Tag             =   "resize(TFFT)"
      Text            =   "0"
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox txtTR 
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   6
      Tag             =   "resize(050,050,100,000)"
      Text            =   "0"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox txtTR 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   5
      Tag             =   "resize(050,050,100,000)"
      Text            =   "200"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox txtTL 
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   4
      Text            =   "0"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox txtTL 
      Height          =   285
      Index           =   0
      Left            =   260
      TabIndex        =   3
      Text            =   "0"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Tag             =   "resize(FTFT)"
      Top             =   3480
      Width           =   735
   End
   Begin VB.PictureBox picPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Index           =   1
      Left            =   3840
      ScaleHeight     =   3000
      ScaleWidth      =   3000
      TabIndex        =   1
      Tag             =   "resize(050,100,100,100)"
      Top             =   360
      Width           =   3000
   End
   Begin VB.PictureBox picPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Index           =   0
      Left            =   360
      ScaleHeight     =   3000
      ScaleWidth      =   3000
      TabIndex        =   0
      Tag             =   "resize(100,050,100,100)"
      Top             =   360
      Width           =   3000
   End
   Begin VB.Label lblSize 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   3600
      Width           =   1815
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReSize As clsRsize
Dim objremember As clsRemember

Dim ImageFile As String
Dim ImageFileChanged As Boolean

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


Private Sub Form_Load()
    Set ReSize = New clsRsize
    Set objremember = New clsRemember
    
    ReSize.HandleFormEx = Me
'#              - Resize.HandleFormEx = Me  (where GUIOBJECT.Tag = resize(FFTT))
'#           OR - Resize.HandleFormEx = Me  (where GUIOBJECT.Tag = resize(nnn,nnn,nnn,nnn))
    ReSize.Ready = True
    SetCaption
    
    SetSourceDefaults

    Set objremember.Attached = Me
    objremember.GetFormState
End Sub

Sub SetCaption()
    Caption = App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision
    If ImageFile <> "" Then
        Caption = ImageFile & " - " & Caption
        If ImageFileChanged Then
            Caption = "*" & Caption
        End If
    End If
End Sub

Sub SetSourceDefaults()
    Me.txtTL(0).Text = "0"
    Me.txtTL(1).Text = "0"
    
    Me.txtTR(0).Text = Me.picPicture(0).ScaleWidth - Screen.TwipsPerPixelX
    Me.txtTR(1).Text = "0"

    Me.txtBR(0).Text = Me.picPicture(0).ScaleWidth - Screen.TwipsPerPixelX
    Me.txtBR(1).Text = Me.picPicture(0).ScaleHeight - Screen.TwipsPerPixelY

    Me.txtBL(0).Text = "0"
    Me.txtBL(1).Text = Me.picPicture(0).ScaleHeight - Screen.TwipsPerPixelY

End Sub


Private Sub cmdGo_Click()

    Me.picPicture(0).Picture = LoadPicture("C:\Users\Steve\Documents\Home\Image Manipulation\chessboard.jpg")
    RedrawImage 0

End Sub

'Private Sub Form_Load()
'    Me.lblSize.Caption = picPicture(0).Width / Screen.TwipsPerPixelX & " x " & picPicture(0).Height / Screen.TwipsPerPixelY
'End Sub

Sub Distort(TL1X As Single, TL1Y As Single, _
            BR1X As Single, BR1Y As Single, _
            TL2X As Single, TL2Y As Single, _
            TR2X As Single, TR2Y As Single, _
            BL2X As Single, BL2Y As Single, _
            BR2X As Single, BR2Y As Single)
            
    Dim W1 As Single
    Dim H1 As Single
    
    Dim W2T As Single
    Dim H2L As Single
    Dim W2B As Single
    Dim H2R As Single
        
    Dim X1 As Single
    Dim Y1 As Single
    
    Dim X2 As Single
    Dim Y2 As Single
    
    Dim DX1 As Single
    Dim DY1 As Single
    
    Dim DX2 As Single
    Dim DY2 As Single
    
    W1 = BR1X - TL1X + 1
    H1 = BR1Y - TL1Y + 1
    
    W2T = TR2X - TL2X + 1
    W2B = BR2X - BL2X + 1
    
    H2L = BL2Y - TL2Y + 1
    H2R = BR2Y - TR2Y + 1
    
    
    For Y1 = TL1Y To BR1Y
        DY1 = (Y1 / H1)
        
        'work out where MINX,Y and MAXX,Y is on distort
        'and what DX2 and DY2 step is for W1 steps
        
        
        
        For X1 = TL1X To BR1X
            DX1 = (X1 / W1)

            'DX1,DY1 is percentage across

        Next
    Next

End Sub

Private Sub txtBL_Change(Index As Integer)
    RedrawImage Index
End Sub

Private Sub txtBR_Change(Index As Integer)
    RedrawImage Index
End Sub

Private Sub txtTL_Change(Index As Integer)
    RedrawImage Index
End Sub

Private Sub txtTR_Change(Index As Integer)
    RedrawImage Index
End Sub

Private Sub RedrawImage(Index As Integer)

    Dim TopLeftX As Integer
    Dim TopRightX As Integer
    Dim BotRightX As Integer
    Dim BotLeftX As Integer
    Dim TopLeftY As Integer
    Dim TopRightY As Integer
    Dim BotRightY As Integer
    Dim BotLeftY As Integer
    
    TopLeftX = Val(txtTL(0).Text)
    TopLeftY = Val(txtTL(1).Text)
    TopRightX = Val(txtTR(0).Text)
    TopRightY = Val(txtTR(1).Text)
    BotLeftX = Val(txtBL(0).Text)
    BotLeftY = Val(txtBL(1).Text)
    BotRightX = Val(txtBR(0).Text)
    BotRightY = Val(txtBR(1).Text)
    
    ' source
    picPicture(0).Cls
    picPicture(0).Line (TopLeftX, TopLeftY)-(TopRightX, TopRightY), vbRed
    picPicture(0).Line -(BotRightX, BotRightY), vbRed
    picPicture(0).Line -(BotLeftX, BotLeftY), vbRed
    picPicture(0).Line -(TopLeftX, TopLeftY), vbRed
    
    ' destination
    picPicture(1).Cls
    picPicture(1).Line (0, 0)-(picPicture(1).ScaleWidth - Screen.TwipsPerPixelX, 0), vbGreen
    picPicture(1).Line -(picPicture(1).ScaleWidth - Screen.TwipsPerPixelX, picPicture(1).ScaleHeight - Screen.TwipsPerPixelY), vbGreen
    picPicture(1).Line -(0, picPicture(1).ScaleHeight - Screen.TwipsPerPixelY), vbGreen
    picPicture(1).Line -(0, 0), vbGreen
    

End Sub

Private Sub Form_Resize()
    ReSize.ReSize
    RedrawImage 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.WindowState = vbNormal Then
        objremember.SaveFormState
    End If
End Sub


