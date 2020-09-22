VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gradientext"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12285
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   12285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana Ref"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10680
      Picture         =   "frmMain.frx":0BF2
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6260
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar pbProgress 
      Height          =   195
      Left            =   9360
      TabIndex        =   16
      Top             =   6960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox txtFont 
      BeginProperty Font 
         Name            =   "Verdana Ref"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5520
      TabIndex        =   15
      Text            =   "Verdana, Arial, Helvetica, Sans-Serif"
      Top             =   6360
      Width           =   5055
   End
   Begin VB.PictureBox pboxBColour 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana Ref"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   13
      Top             =   6720
      Width           =   375
   End
   Begin VB.CommandButton cmdHTML 
      Caption         =   "HTML"
      BeginProperty Font 
         Name            =   "Verdana Ref"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10680
      Picture         =   "frmMain.frx":6944
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5350
      Width           =   1575
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana Ref"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   9
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana Ref"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   8
      Top             =   5400
      Width           =   615
   End
   Begin VB.PictureBox pboxColour 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Verdana Ref"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   5520
      ScaleHeight     =   795
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox txtHTMLText 
      BeginProperty Font 
         Name            =   "Verdana Ref"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmMain.frx":863E
      Top             =   4430
      Width           =   4695
   End
   Begin SHDocVwCtl.WebBrowser wbPreview 
      Height          =   4335
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12255
      ExtentX         =   21616
      ExtentY         =   7646
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "Verdana Ref"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5520
      TabIndex        =   3
      Top             =   4800
      Width           =   5055
   End
   Begin VB.TextBox txtSize 
      BeginProperty Font 
         Name            =   "Verdana Ref"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5520
      TabIndex        =   2
      Text            =   "-1"
      Top             =   4440
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog comdiag 
      Left            =   7920
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Verdana Ref"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10680
      Picture         =   "frmMain.frx":865F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1575
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   17
      Top             =   7185
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16193
            Text            =   "Ready..."
            TextSave        =   "Ready..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2823
            MinWidth        =   2823
            TextSave        =   "06/11/2001"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana Ref"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblFont 
      Alignment       =   1  'Right Justify
      Caption         =   "Font:"
      BeginProperty Font 
         Name            =   "Verdana Ref"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   14
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label lblBack 
      Alignment       =   1  'Right Justify
      Caption         =   "Back:"
      BeginProperty Font 
         Name            =   "Verdana Ref"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   6780
      Width           =   615
   End
   Begin VB.Label lblColour 
      Alignment       =   1  'Right Justify
      Caption         =   "Colour:"
      BeginProperty Font 
         Name            =   "Verdana Ref"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label lblPath 
      Alignment       =   1  'Right Justify
      Caption         =   "Path:"
      BeginProperty Font 
         Name            =   "Verdana Ref"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblSize 
      Alignment       =   1  'Right Justify
      Caption         =   "Size:"
      BeginProperty Font 
         Name            =   "Verdana Ref"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   4440
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCreate_Click()

    sbStatus.Panels(1).Text = "Creating..."

    'Create Text Gradient
    CreateGradientFX txtHTMLText.Text, txtPath.Text, txtSize.Text, txtFont.Text
    
    sbStatus.Panels(1).Text = "Ready..."
    pbProgress.Value = 0
    

    'Preview Web Page
    wbPreview.Navigate txtPath.Text


End Sub

Private Sub cmdAdd_Click()


    'If there isn't already 10 Pbox's then add another
    If pboxColour.UBound <> 9 Then
        AddPicBox pboxColour(pboxColour.UBound).BackColor
    End If


End Sub

Private Sub cmdHTML_Click()


    'Open HTML File in Notepad
    Shell "C:\WINDOWS\NOTEPAD.EXE " & txtPath.Text, vbNormalFocus


End Sub

Private Sub cmdRemove_Click()


    'If there isn't only one left, delete the one furthest to the right
    If pboxColour.UBound <> 0 Then
        Unload pboxColour(pboxColour.UBound)
    End If


End Sub

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    'Add the starting pictureboxes
    AddPicBox vbYellow
    AddPicBox vbGreen
    AddPicBox vbCyan
    AddPicBox vbBlue
    AddPicBox vbMagenta
    
    'Set File Path
    txtPath.Text = "C:\Windows\Desktop\Gradient.html"
    
    'Create Text Gradient
    CreateGradientFX txtHTMLText.Text, txtPath.Text, txtSize.Text, txtFont.Text

    'Preview Web Page
    wbPreview.Navigate txtPath.Text
    

End Sub

Private Sub Form_Resize()

    With pbProgress
        .Left = sbStatus.Panels(2).Left + 30
        .Top = sbStatus.Top + 45
        .Width = sbStatus.Panels(2).Width - 45
        .Height = sbStatus.Height - 75
        .Value = 0
    End With
    
    With wbPreview
        .Width = Me.Width - 120
    End With

End Sub

Private Sub pboxColour_Click(index As Integer)

    'Show common dialog
    comdiag.ShowColor
    
    'Paint Picturebox
    pboxColour(index).BackColor = comdiag.Color


End Sub

Private Sub pboxBColour_Click()

    'Show common dialog
    comdiag.ShowColor
    
    'Paint Picturebox
    pboxBColour.BackColor = comdiag.Color


End Sub

Public Function ResizeControls()


    'Resize big text box
    With txtHTMLText
        .Left = Me.Width - .Width - 120
        .Height = Me.Height - pbProgress.Height - cmdCreate.Height - cmdHTML.Height - 720
    End With
    
    'Resize Preview Box
    With wbPreview
        .Width = txtHTMLText.Left - 120
        .Height = txtHTMLText.Height + 120 + cmdCreate.Height + cmdHTML.Height
    End With
    
    'Resize Progress Bar
    With pbProgress
        .Top = wbPreview.Height + 200
        .Width = Me.Width - 120
    End With
    
    'Resize Create Button
    With cmdCreate
        .Top = txtHTMLText.Height + 180
        .Left = txtHTMLText.Left + txtHTMLText.Width - .Width
    End With
    
    'Resize HTML Button
    cmdHTML.Top = cmdCreate.Top + cmdCreate.Height
    cmdHTML.Left = cmdCreate.Left
    
    'Resize small textboxes
    txtSize.Top = txtHTMLText.Height + 180
    txtPath.Top = txtSize.Top + txtSize.Height + 105
    txtFont.Top = txtPath.Top + txtPath.Height + 105 + txtPath.Height + 105
    pboxBColour.Top = txtFont.Top
    
    'Resize Colour Boxes
    For x = 0 To pboxColour.UBound Step 1
        pboxColour(x).Top = txtPath.Top + txtPath.Height + 105
    Next x
    
    'Resize Labels
    lblSize.Top = txtSize.Top
    lblPath.Top = txtPath.Top
    lblColour.Top = pboxColour(0).Top
    lblColour.Left = wbPreview.Width - 15
    lblSize.Left = lblColour.Left
    lblPath.Left = lblColour.Left
    lblBack.Left = lblColour.Left
    lblBack.Top = txtFont.Top
    lblFont.Top = lblBack.Top
    lblFont.Left = lblBack.Left + lblBack.Width + 345
    
    'Resize other buttons
    cmdAdd.Top = lblColour.Top
    cmdRemove.Top = lblColour.Top
    cmdRemove.Left = cmdCreate.Left - 105 - cmdRemove.Width
    cmdAdd.Left = cmdRemove.Left - cmdAdd.Width
    
    'Resize small textboxes
    txtSize.Left = lblColour.Left + lblColour.Width + 105
    txtPath.Left = txtSize.Left
    pboxBColour.Left = txtPath.Left
    
    'Resize Colour Boxes
    For x = 0 To pboxColour.UBound Step 1
        pboxColour(x).Left = txtPath.Left + (x * pboxColour(x).Width)
    Next x
    
    txtFont.Left = lblFont.Left + lblFont.Width + 105
    txtFont.Width = txtPath.Left + txtPath.Width - txtFont.Left
    
    wbPreview.Height = txtFont.Top + txtFont.Height - wbPreview.Top
    
    cmdHTML.Height = wbPreview.Top + wbPreview.Height - cmdHTML.Top

End Function


