VERSION 5.00
Begin VB.Form frmConfirmation 
   Caption         =   "Confirmation"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblClearMess 
      Alignment       =   2  'Center
      Caption         =   "Do you wish to clear the picture?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label lblSaveMess 
      Alignment       =   2  'Center
      Caption         =   "Do you want to save before clearing the image?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   4935
   End
End
Attribute VB_Name = "frmConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Yes As Integer
Private Sub cmdNo_Click()
    
    If Yes = 1 Then
    frmPaint.picMainPic.Cls
    End If
    
    frmPaint.Enabled = True
    Unload Me
End Sub

Private Sub cmdYes_Click()
    
    Yes = Yes + 1
    
    If Yes = 2 Then
    GoTo 2
    End If
    
    lblClearMess.Visible = False
    lblSaveMess.Visible = True
    If Yes = 1 Then
        GoTo 1
    End If
    
2:
    Save
    frmPaint.picMainPic.Cls
    Unload Me
    
1:
    
    frmPicSize.Show
    frmConfirmation.Visible = False
    frmConfirmation.Enabled = False
    
    Unload Me
    
End Sub

Sub Save()
    
    With frmPaint.dlgCommon
        .DialogTitle = "Choose a filename to save"
        .Filter = "Bitmap files (*.BMP)|*.BMP"
        .FilterIndex = 1
        .FileName = ""
        .ShowSave
        
        If .FileName = "" Then Exit Sub
        
        SavePicture frmPaint.picMainPic.Image, .FileName
    End With
    
    
End Sub

Private Sub Form_Load()
    Yes = 0
End Sub
