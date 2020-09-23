VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPaint 
   Caption         =   "^^^zErOcOoL's PaInT pRoGrAm^^^"
   ClientHeight    =   6645
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScrPicMainPic 
      Height          =   7335
      Left            =   11640
      TabIndex        =   8
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScrPicMainPic 
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   7320
      Width           =   8175
   End
   Begin MSComctlLib.StatusBar staBarPaint 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7935
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   15319
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            TextSave        =   "05/01/2002"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            TextSave        =   "10:54 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picMainPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7335
      Left            =   3480
      ScaleHeight     =   7305
      ScaleWidth      =   8145
      TabIndex        =   3
      Top             =   0
      Width           =   8175
      Begin VB.Line HorizontalMirrorLine 
         BorderWidth     =   5
         Visible         =   0   'False
         X1              =   0
         X2              =   8520
         Y1              =   4087
         Y2              =   4087
      End
      Begin VB.Line VertMirrorLine 
         BorderWidth     =   5
         Visible         =   0   'False
         X1              =   4260
         X2              =   4260
         Y1              =   -360
         Y2              =   7560
      End
   End
   Begin VB.PictureBox picToolBar 
      Height          =   8175
      Left            =   0
      ScaleHeight     =   8115
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   0
      Width           =   3375
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   2220
         Left            =   240
         TabIndex        =   6
         Top             =   2640
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   3916
         ButtonWidth     =   1799
         ButtonHeight    =   953
         AllowCustomize  =   0   'False
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Pencil"
               Key             =   "Pencil"
               Description     =   "Draw within the Box"
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eraser"
               Key             =   "Eraser"
               Description     =   "Erases the contents where ever the curser is locted."
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Paint Bucket"
               Key             =   "PaintBucket"
               Description     =   "Fills the whole background a certain color."
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Mirror"
               Key             =   "Mirror"
               Description     =   "Mirrors objects depending on what you select."
            EndProperty
         EndProperty
      End
      Begin MSComDlg.CommonDialog dlgCommon 
         Left            =   2760
         Top             =   2040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox ForeColorSample 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   600
         ScaleHeight     =   465
         ScaleWidth      =   705
         TabIndex        =   4
         Top             =   1800
         Width           =   735
      End
      Begin VB.PictureBox BackColorSample 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1200
         ScaleHeight     =   465
         ScaleWidth      =   705
         TabIndex        =   5
         Top             =   2040
         Width           =   735
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         Picture         =   "frmPaint.frx":0000
         ScaleHeight     =   1665
         ScaleWidth      =   2985
         TabIndex        =   2
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.PictureBox picBase 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   3480
      ScaleHeight     =   7575
      ScaleWidth      =   8415
      TabIndex        =   9
      Top             =   0
      Width           =   8415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewPicture 
         Caption         =   "New &Picture"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuCustomColor 
         Caption         =   "Custom Color"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuDrawingSize 
         Caption         =   "&Drawing Size"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuRainboPaint 
         Caption         =   "&Rainbo Paint"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuPicSize 
         Caption         =   "Picture Si&ze"
         Shortcut        =   ^Z
      End
   End
End
Attribute VB_Name = "frmPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
        
    frmPicSize.Show
    frmPaint.Enabled = False
    
    blnRainboPaint = False
    blnDraw = False
    
    blnVertM = False
    blnHorM = False
    blnBothM = False
    
    blnPaintBuck = False
    
    blnLeft = True
    blnRight = False
    
    blnErase = False
    
    intDrawWidth = 4
    intStyle = 0
    intPaint = 0
    
    intRed = 0
    intGreen = 0
    intBlue = 0
    
    BCRed = 255
    BCGreen = 255
    BCBlue = 255
    
    blnCav1 = False
    blnCav2 = True
    blnCav3 = False
    blnCav4 = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Dim strAns As String
    
    strAns = MsgBox("Are You Sure You Want to Quit zErOcOoL's PaInT pRoGrAm!?", vbQuestion + vbYesNo, "Quit?")
    
    If strAns = vbYes Then
        End
    Else
        Cancel = 1
    End If
    
End Sub

Private Sub HScrPicMainPic_Change()
    picMainPic.Width = picMainPic.Width + HScrPicMainPic.value
End Sub

Private Sub HScrPicMainPic_Scroll()
    picMainPic.Width = picMainPic.Width + HScrPicMainPic.value
End Sub

Private Sub mnuCustomColor_Click()
    frmPaint.Enabled = False
    frmColorButton.Show
End Sub

Private Sub mnuDrawingSize_Click()
    frmPaint.Enabled = False
    frmDrawingSize.Show
End Sub

Private Sub mnuExit_Click()
    Dim strAns As String
        
    strAns = MsgBox("Are You Sure You Want to Quit zErOcOoL's PaInT pRoGrAm!?", vbQuestion + vbYesNo, "Quit?")
    
    If strAns = vbYes Then
        End
    Else
        Exit Sub
    End If
    
End Sub

Private Sub mnuNewPicture_Click()
    frmConfirmation.Show
    frmPaint.Enabled = False
End Sub

Private Sub mnuPicSize_Click()
    frmPicSize.Show
    frmPaint.Enabled = False
End Sub

Private Sub mnuRainboPaint_Click()
    
    If mnuRainboPaint.Checked = False Then
        blnRainboPaint = True
        mnuRainboPaint.Checked = True
    Else
        blnRainboPaint = False
        mnuRainboPaint.Checked = False
    End If
    
End Sub

Private Sub mnuSave_Click()
    
    With dlgCommon
        
        .DialogTitle = "Choose a filename to save"
        .Filter = "Bitmap files (*.BMP)|*.BMP"
        .FilterIndex = 1
        .FileName = ""
        .ShowSave
        
        If .FileName = "" Then Exit Sub
        
        SavePicture picMainPic.Image, .FileName
    End With
    
End Sub

Private Sub picMainPic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Shift = 1 And Button = 2 Then
        
        PopupMenu mnuOptions
        GoTo MenuEnd
        
    End If
    
    If Button = 1 Then

        picMainPic.CurrentX = x
        picMainPic.CurrentY = y
        blnDraw = True
        picMainPic.Line -(x, y), RGB(0, 0, 0)
    
        If blnLeft = True Then
            
            picMainPic.DrawWidth = intDrawWidth / 3.14
            picMainPic.DrawStyle = intStyle
            picMainPic.Line -(x, y), RGB(intRed, intGreen, intBlue)
        
            Exit Sub
        
        End If
    
        If blnErase = True Then
            picMainPic.DrawWidth = intDrawWidth / 3.14
            picMainPic.DrawStyle = intStyle
            picMainPic.Line -(x, y), RGB(255, 255, 255)
        
            Exit Sub
    
        End If
    
        If blnErase = False Then
        
            picMainPic.DrawWidth = intDrawWidth / 3.14
            picMainPic.DrawStyle = intStyle
            picMainPic.Line -(x, y), RGB(intRed, intGreen, intBlue)
        
            Exit Sub
    
        End If
    
        If blnPaintBuck = True Then
            picMainPic.BackColor = RGB(intRed, intGreen, intBlue)
            Exit Sub
        End If
    
        If blnVertM Then
            Call MirrorDraw(x, y, oldx, oldy, 0)
        End If
            
        If blnHorM Then
            Call MirrorDraw(x, y, oldx, oldy, 1)
        End If
    
        If blnBothM Then
            Call MirrorDraw(x, y, oldx, oldy, 2)
        End If
        
        If blnRainboPaint = True Then
            
            Call RainboPaint
            
        End If

    ForeColorSample.BackColor = RGB(intRed, intGreen, intBlue)
    picMainPic.Line -(x, y), RGB(intRed, intGreen, intBlue)
    picMainPic.DrawWidth = intDrawWidth / 3.14
    picMainPic.DrawStyle = intStyle

ElseIf Button = 2 Then
        picMainPic.CurrentX = x
        picMainPic.CurrentY = y
        blnDraw = True
    
        If blnRight = True Then
            picMainPic.DrawWidth = intDrawWidth / 3.14
            picMainPic.DrawStyle = intStyle
            picMainPic.Line -(x, y), RGB(BCRed, BCGreen, BCBlue)
        
            Exit Sub
    
        End If
    
        If blnErase = True Then
            picMainPic.DrawWidth = intDrawWidth / 3.14
            picMainPic.DrawStyle = intStyle
            picMainPic.Line -(x, y), RGB(BCRed, BCGreen, BCBlue)
        
            Exit Sub
    
        End If
    
        If blnErase = False Then
        
            picMainPic.DrawWidth = intDrawWidth / 3.14
            picMainPic.DrawStyle = intStyle
            picMainPic.Line -(x, y), RGB(BCRed, BCGreen, BCBlue)
        
            Exit Sub
    
        End If
    
        If blnPaintBuck = True Then
            picMainPic.BackColor = RGB(BCRed, BCGreen, BCBlue)
            Exit Sub
        End If
    
        If blnVertM Then
            Call MirrorDraw(x, y, oldx, oldy, 0)
        End If
            
        If blnHorM Then
            Call MirrorDraw(x, y, oldx, oldy, 1)
        End If
    
        If blnBothM Then
            Call MirrorDraw(x, y, oldx, oldy, 2)
        End If
        
        If blnRainboPaint = True Then
            
            Call RainboPaint
            
        End If
            
            
            BackColorSample.ForeColor = RGB(BCRed, BCGreen, BCBlue)
            picMainPic.Line -(x, y), RGB(BCRed, BCGreen, BCBlue)
            picMainPic.DrawWidth = intDrawWidth / 3.14
            picMainPic.DrawStyle = intStyle
            
End If

MenuEnd:

End Sub

Private Sub picMainPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
    
        If Left = True Then
        
            If blnDraw Then
                
                picMainPic.Line -(x, y), RGB(intRed, intGreen, intBlue)
                picMainPic.DrawWidth = intDrawWidth / 3.14
                picMainPic.DrawStyle = intStyle
        
                Exit Sub
    
            End If
        
        End If
    
        If blnErase = True Then
        
            picMainPic.DrawWidth = intDrawWidth / 3.14
            picMainPic.DrawStyle = intStyle
            picMainPic.Line -(x, y), RGB(255, 255, 255)
            Exit Sub
        
        End If
    
        If blnErase = False Then
        
            picMainPic.Line -(x, y), QBColor(intPaint)
            picMainPic.DrawWidth = intDrawWidth / 3.14
            picMainPic.DrawStyle = intStyle
        
            Exit Sub
        
        End If
    
        If blnVertM = True Then
            Call MirrorDraw(x, y, oldx, oldy, 0)
        End If
        
        If blnHorM = True Then
            Call MirrorDraw(x, y, oldx, oldy, 1)
        End If
    
        If blnBothM = True Then
            Call MirrorDraw(x, y, oldx, oldy, 2)
        End If
        
        oldx = x
        oldy = y
        
        If blnRainboPaint = True Then
            
            Call RainboPaint
            ForeColorSample.BackColor = RGB(intRed, intGreen, intBlue)
            picMainPic.Line -(x, y), RGB(intRed, intGreen, intBlue)
            picMainPic.DrawWidth = intDrawWidth / 3.14
            picMainPic.DrawStyle = intStyle
            
        End If
    
End If

    If Button = 2 Then
        
        If blnRight = True Then
        
            If blnDraw Then
                
                picMainPic.Line -(x, y), RGB(BCRed, BCGreen, BCBlue)
                picMainPic.DrawWidth = intDrawWidth / 3.14
                picMainPic.DrawStyle = intStyle
        
                Exit Sub
    
            End If
        
        End If
    
        If blnErase = True Then
        
            picMainPic.DrawWidth = intDrawWidth / 3.14
            picMainPic.DrawStyle = intStyle
            picMainPic.Line -(x, y), RGB(255, 255, 255)
            Exit Sub
        
        End If
    
        If blnErase = False Then
        
            picMainPic.Line -(x, y), RGB(BCRed, BCGreen, BCBlue)
            picMainPic.DrawWidth = intDrawWidth / 3.14
            picMainPic.DrawStyle = intStyle
        
            Exit Sub
        
        End If
    
        If blnVertM = True Then
            Call MirrorDraw(x, y, oldx, oldy, 0)
        End If
        
        If blnHorM = True Then
            Call MirrorDraw(x, y, oldx, oldy, 1)
        End If
    
        If blnBothM = True Then
            Call MirrorDraw(x, y, oldx, oldy, 2)
        End If
        
        oldx = x
        oldy = y
        
        If blnRainboPaint = True Then
            
            Call RainboPaint
            BackColorSample.BackColor = RGB(BCRed, BCGreen, BCBlue)
            picMainPic.Line -(x, y), RGB(BCRed, BCGreen, BCBlue)
            picMainPic.DrawWidth = intDrawWidth / 3.14
            picMainPic.DrawStyle = intStyle
            
        End If
    
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Pencil = 1
    Eraser = 2
    PaintBucket = 3
    mirror = 4
    
    Select Case Button.Key
        
        Case "Pencil"
        
        Toolbar1.Buttons(1).value = tbrPressed
        
        blnErase = False
        blnDraw = True
        
        If blnHorM = False Then
        HorizontalMirrorLine.Visible = False
        End If
        
        If blnVertM = False Then
        VertMirrorLine.Visible = False
        End If
        
        If blnBothM = True Then
        VertMirriorLine.Visible = True
        HorizontalMirrorLine.Visible = True
        'Call Mirrior
        End If
        
        blnBothM = False
        
        Case "Eraser"

        Toolbar1.Buttons(2).value = tbrPressed
        'Toolbar1.Buttons(1).value = tbrUnpressed

        blnErase = True
        blnDraw = False
        blnCustomPaint = False
        
        If blnHorM = False Then
        HorizontalMirrorLine.Visible = False
        End If
        
        If blnVertM = False Then
        VertMirrorLine.Visible = False
        End If
        
        If blnBothM = True Then
        VertMirriorLine.Visible = True
        HorizontalMirrorLine.Visible = True
        'Call Mirrior
        End If
            
        Case "PaintBucket"
        
        Toolbar1.Buttons(3).value = tbrPressed
        VertMirrorLine.Visible = False
        HorizontalMirrorLine.Visible = False
        
        If blnHorM = False Then
        HorizontalMirrorLine.Visible = False
        End If
        
        If blnVertM = False Then
        VertMirrorLine.Visible = False
        End If
        
        If blnBothM = True Then
        VertMirriorLine.Visible = True
        HorizontalMirrorLine.Visible = True
        'Call Mirrior
        End If
        
        blnErase = False
        
        Case "Mirror"
        frmMirror.Show
        frmPaint.Enabled = False
        
    
    End Select
    
End Sub

Private Sub VScrPicMainPic_Change()
    picMainPic.Height = picMainPic.Height + VScrPicMainPic.value
End Sub

Private Sub VScrPicMainPic_Scroll()
    picMainPic.Height = picMainPic.Height + VScrPicMainPic.value
End Sub
