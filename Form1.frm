VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4410
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7080
   ScaleWidth      =   4410
   StartUpPosition =   1  'Fenstermitte
   Begin MSComDlg.CommonDialog CommonDialogButtonCol 
      Left            =   345
      Top             =   255
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Projekt1.Button ButtonExit 
      Height          =   765
      Left            =   1530
      TabIndex        =   0
      ToolTipText     =   "Exit"
      Top             =   4785
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1349
      ForeColor       =   255
      TX              =   "Exit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Projekt1.Button ButtonChangeCol 
      Height          =   510
      Left            =   1170
      TabIndex        =   1
      ToolTipText     =   "Change Button Color"
      Top             =   3735
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1349
      ForeColor       =   0
      TX              =   "Change Button Color"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Projekt1.Button Button1 
      Height          =   510
      Left            =   1170
      TabIndex        =   2
      Top             =   2715
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1349
      ForeColor       =   0
      TX              =   "Dummy"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Projekt1.Button Button2 
      Height          =   510
      Left            =   1185
      TabIndex        =   3
      Top             =   1755
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1349
      ForeColor       =   0
      TX              =   "Dummy"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This XP-Style Button can be placed on any background
'MouseOver and TabStop will be highlighted
'The color of the button can be adapted to any color during runtime

'Please feel invited to visit my homepage
'http://home.t-online.de/home/l.kobarg/clk/
'There you can find a calculator using the XP-Style Button

'if you got any improvements, maybe round, or oval shapes please let me know
'l.kobarg@t-onlien.de

'Based on Leo Barsukov's cool Totally skinned Calculator********
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=38467&lngWId=1

'and Gez Lemon's transparent tutorials
'http://www.juicystudio.com/tutorial/vb/winapi.asp

'known issues:
'during programming the auto-type (auto completion) will not work properly
'if a form using the XP-Style button is open


Dim AktButtonColor As Long


Private Sub ButtonChangeCol_Click()
    On Error Resume Next
    
    Err.Clear
    CommonDialogButtonCol.CancelError = True
    CommonDialogButtonCol.ShowColor
    If Not (Err.Number > 0) Then
        AktButtonColor = CommonDialogButtonCol.Color
        SetButtonColor
    End If
    
End Sub

Private Sub ButtonExit_Click()
    End
End Sub

Private Sub SetButtonColor()
    ButtonChangeCol.cFace = AktButtonColor
    ButtonExit.cFace = AktButtonColor
    
    ButtonChangeCol.Refesh
    ButtonExit.Refesh

End Sub

Private Sub Form_Initialize()
    AktButtonColor = 8454016
    SetButtonColor
End Sub


