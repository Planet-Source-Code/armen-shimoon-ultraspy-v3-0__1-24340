VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Form1.frx":08CA
   ScaleHeight     =   7860
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMakeSmall 
      Caption         =   "<--"
      Height          =   465
      Left            =   8370
      TabIndex        =   26
      Top             =   2520
      Width           =   555
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   555
      Left            =   3060
      TabIndex        =   25
      Top             =   2430
      Width           =   1455
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   555
      Left            =   1620
      TabIndex        =   24
      Top             =   2430
      Width           =   1365
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4770
      Top             =   2430
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Picture"
      Height          =   555
      Left            =   180
      TabIndex        =   22
      Top             =   2430
      Width           =   1365
   End
   Begin VB.Frame Frame3 
      Caption         =   "Screenshot"
      Height          =   4605
      Left            =   90
      TabIndex        =   21
      Top             =   3150
      Width           =   8745
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         Height          =   4065
         Left            =   180
         ScaleHeight     =   4005
         ScaleWidth      =   8235
         TabIndex        =   23
         Top             =   360
         Width           =   8295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Altering Information"
      Height          =   2175
      Left            =   4410
      TabIndex        =   2
      Top             =   90
      Width           =   4425
      Begin VB.TextBox txtStatic 
         Height          =   285
         Left            =   1710
         TabIndex        =   18
         Text            =   "<null>"
         Top             =   990
         Width           =   2445
      End
      Begin VB.TextBox txtDis 
         Height          =   285
         Left            =   1710
         TabIndex        =   16
         Text            =   "<null>"
         Top             =   630
         Width           =   2445
      End
      Begin VB.TextBox txtKill 
         Height          =   285
         Left            =   1710
         TabIndex        =   14
         Text            =   "<null>"
         Top             =   1710
         Width           =   2445
      End
      Begin VB.TextBox txtClassText 
         Height          =   285
         Left            =   1710
         TabIndex        =   12
         Text            =   "<null>"
         Top             =   1350
         Width           =   2445
      End
      Begin VB.TextBox txtEnable 
         Height          =   285
         Left            =   1710
         TabIndex        =   11
         Text            =   "<null>"
         Top             =   270
         Width           =   2445
      End
      Begin VB.Label Label10 
         Caption         =   "Window to change:"
         Height          =   285
         Left            =   270
         TabIndex        =   17
         Top             =   1080
         Width           =   2265
      End
      Begin VB.Label Label9 
         Caption         =   "Disable Window:"
         Height          =   195
         Left            =   270
         TabIndex        =   15
         Top             =   720
         Width           =   2085
      End
      Begin VB.Label Label8 
         Caption         =   "Minimize Window:"
         Height          =   195
         Left            =   270
         TabIndex        =   13
         Top             =   1800
         Width           =   1545
      End
      Begin VB.Label Label7 
         Caption         =   "Class to change:"
         Height          =   330
         Left            =   270
         TabIndex        =   10
         Top             =   1440
         Width           =   2325
      End
      Begin VB.Label Label3 
         Caption         =   "Enable Window:"
         Height          =   240
         Left            =   270
         TabIndex        =   3
         Top             =   360
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "UltraSpy Information"
      Height          =   2175
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3975
      Begin VB.TextBox txtHandle 
         Height          =   285
         Left            =   810
         TabIndex        =   20
         Top             =   360
         Width           =   2265
      End
      Begin VB.TextBox txtParent 
         Height          =   285
         Left            =   810
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   2265
      End
      Begin VB.TextBox txtClass 
         Height          =   285
         Left            =   810
         TabIndex        =   5
         Top             =   720
         Width           =   2265
      End
      Begin VB.TextBox txtUM 
         Height          =   285
         Left            =   810
         TabIndex        =   4
         Top             =   1080
         Width           =   2265
      End
      Begin VB.PictureBox Picture1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   1
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label11 
         Caption         =   "Handle:"
         Height          =   285
         Left            =   180
         TabIndex        =   19
         Top             =   450
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Parent:"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1530
         Width           =   570
      End
      Begin VB.Label Label2 
         Caption         =   "Text:"
         Height          =   180
         Left            =   180
         TabIndex        =   8
         Top             =   1170
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Class:"
         Height          =   270
         Left            =   180
         TabIndex        =   7
         Top             =   810
         Width           =   450
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim verylong As String * 100
Dim gParent As String * 100
Dim SndMsg As String * 100
Dim windowname As String * 100
Dim sztext As String * 100
Dim mousemove As Boolean
Dim Pic01 As Boolean
Dim SmallBL As Boolean







Private Sub cmdAbout_Click()
MsgBox "UltraSpy v3.0 by Shimoon Technologies." & vbCrLf & vbCrLf & "Compiled on June 22, 2001" & vbCrLf & "Coded by Armen Shimoon." & vbCrLf & "©2001 Shimoon Technologies.", vbOKOnly, "About"


End Sub

Private Sub cmdExit_Click()
Unload frmSplash
Unload Form1



End Sub

Private Sub cmdMakeSmall_Click()

If SmallBL = True Then
    Form1.Height = 8235
    cmdMakeSmall.Caption = "<--"
    SmallBL = False
    
Else
    Form1.Height = 3495
    cmdMakeSmall.Caption = "-->"
    SmallBL = True
End If

End Sub

Private Sub cmdSave_Click()
 Dim CheckFile As Boolean
 Dim strFileName As String
 
 CommonDialog1.ShowSave
 
 
 strFileName = CommonDialog1.FileName

 

    
    DoEvents
     On Error GoTo 20
    stdole.SavePicture Picture2.Image, strFileName
    DoEvents
MsgBox "Saved to " & strFileName, vbInformation, "Saved"

20: Exit Sub




End Sub

Private Sub Form_Load()
Picture1.Picture = LoadResPicture(102, vbResIcon)
Form1.Icon = LoadResPicture(101, vbResIcon)
Form1.Caption = LoadResString(101) & " " & LoadResString(102) & " v3.0" & "  -  " & LoadResString(103)

mousemove = False
TextRO txtClass

TextRO txtParent
TextRO txtUM

Pic01 = True


KeepOnTop Form1
SmallBL = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmSplash
Unload Form1

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture1.Picture = Nothing
Form1.MouseIcon = LoadResPicture(102, vbResIcon)
Form1.MousePointer = 99
mousemove = True

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim cursorpos1 As POINTAPI
   Dim wintext As String
   Dim garmon As String
   Dim gIcon As Image
   Dim OldX As Integer
   Dim OldY As Integer
   Dim ttxt As String
   Dim abc As String
   Dim WndRECT As RECT
   Dim Width1 As Integer, Height1 As Integer
   
 If mousemove = True Then
 
    r = GetCursorPos(cursorpos1)
    hwnd1 = WindowFromPoint(cursorpos1.x, cursorpos1.y)
    r = GetClassName(hwnd1, sztext, 100)
    hwnd2 = WindowFromPoint(cursorpos1.x, cursorpos1.y)
    p = GetWindowText(hwnd2, windowname, 100)
    hwnd3 = WindowFromPoint(cursorpos1.x, cursorpos1.y)
    q = GetParent(hwnd3)
    Call GetWindowRect(hwnd1, WndRECT)
    v& = GetDC(hwnd1)
    Width1 = WndRECT.Left + WndRECT.Right
    Height1 = WndRECT.Top + WndRECT.Bottom
    
    Picture2.Picture = Nothing
    
    
    Call BitBlt(Picture2.hdc, 0, 0, 600, 500, v&, 0, 0, vbSrcCopy)
    Call ReleaseDC(hwnd1, v&)
    

              ttxt = Space(100)
              errval = GetCursorPos(cursorpos1)
              thwnd = WindowFromPoint(cursorpos1.x, cursorpos1.y)
              errval = SendMessage(thwnd, WM_GETTEXT, ByVal TXT_LEN, ByVal ttxt)
              ttxt = RTrim(ttxt)
              
              
    txtUM.Text = ttxt
    txtHandle.Text = hwnd1
    txtParent.Text = q

    txtClass.Text = sztext



If txtClass.Text = txtClassText.Text Then
    a = InputBox("New string for " & txtClass & ":", "New string")
    b = SetWindowText(hwnd1, a)
    txtClassText.Text = "<null>"
    Picture1.Picture = LoadResPicture(102, vbResIcon)
    Form1.MousePointer = 0
ElseIf txtUM.Text = Frame2.Caption Then
    Picture1.Picture = LoadResPicture(102, vbResIcon)
    Form1.MousePointer = 0
    mousemove = False
ElseIf txtKill.Text = txtUM.Text Then
    a = CloseWindow(hwnd2)
    txtKill.Text = "<null>"
    Picture1.Picture = LoadResPicture(102, vbResIcon)
    Form1.MousePointer = 0
ElseIf txtDis.Text = txtUM.Text Then
    a = EnableWindow(hwnd2, 0)
    txtDis.Text = "<null>"
    Picture1.Picture = LoadResPicture(102, vbResIcon)
    Form1.MousePointer = 0
ElseIf txtUM.Text = txtStatic.Text Then
    abc = InputBox("New string for static " & txtUM.Text, "New string")
    Call SendMessage(hwnd2, WM_SETTEXT, 0&, ByVal abc)
    txtStatic.Text = "<null>"
    Picture1.Picture = LoadResPicture(102, vbResIcon)
    Form1.MousePointer = 0
ElseIf txtUM.Text = txtEnable.Text Then
    Call EnableWindow(hwnd2, 0&)
    txtEnable.Text = "<null>"
    Picture1.Picture = LoadResPicture(102, vbResIcon)
    Form1.MousePointer = 1
End If
 
 
    
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture1.Picture = LoadResPicture(102, vbResIcon)
Form1.MousePointer = 0
mousemove = False
End Sub


Private Function TextRO(textbx As TextBox)
a = SendMessage(textbx.hwnd, EM_SETREADONLY, 1, 0)
End Function



Public Function AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)


    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hwnd, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Function


