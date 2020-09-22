VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Shortcut"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Text            =   "New Shortcut"
      Top             =   1920
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   600
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdCreateShortcut 
      Caption         =   "Create Shortcut"
      Height          =   375
      Left            =   1673
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   285
      Left            =   3840
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtTarget 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Text            =   "\Startup"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.OptionButton optOther 
      Caption         =   "StartMenu > Other"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.OptionButton optPrograms 
      Caption         =   "StartMenu > Programs"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.OptionButton optDesktop 
      Caption         =   "Desktop"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Name for shortcut :"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Target :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "Create shortcut in :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''#####################################################''
''##                                                 ##''
''##  Created By BelgiumBoy_007                      ##''
''##                                                 ##''
''##  Visit BartNet @ www.bartnet.be for more Codes  ##''
''##                                                 ##''
''##  Copyright 2003 BartNet Corp.                   ##''
''##                                                 ##''
''#####################################################''

'This is the API that will create the shortcut.  It's used in the cmdCreateShortcut_Click() sub.
Private Declare Function fCreateShellLink Lib "Vb5stkit.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long

Private Sub cmdBrowse_Click()
On Error GoTo skip                                      'An error will occur when the user presses Cancel.
    CommonDialog.DialogTitle = "Select a File"          'Set the text at the top of the Dialog (if we don't then the text Open will appear).
    CommonDialog.FileName = App.Path                    'Set the default path.
    CommonDialog.Filter = "Applications|*.exe"          'Set the file types to allow.
    CommonDialog.ShowOpen                               'Show the Dialog.
    txtTarget.Text = CommonDialog.FileName              'Set the text to the selected file.
    txtTarget.SelStart = Len(txtTarget.Text)            'Put the cursor at the back of the text.
skip:

    'With a CommonDialog we can specify the file types to accept.  It will only show files that it can accept.
    'You can also add multiple file types to accept.  Here is an example :
    
    '"Images|*.JPG; *.jpg; *.gif; *.GIF"                  -   This will display every file with the exention .JPG, .jpg, .gif or .GIF.
    '"Images|*.JPG; *.jpg; *.gif; *.GIF|All Files@*.*"    -   This will display every file with the exention .JPG, .jpg, .gif or .GIF AND it will allow the user to select 'All Files' in the dropdown combo.
    
    'To understand these exaples better, just try them out and for all means EXPERIMENT with them !
End Sub

Private Sub cmdCreateShortcut_Click()
    txtOther_LostFocus                                  'Make sure this happens.
    
'On Error GoTo err
    Open txtTarget.Text For Binary Access Read As #1    'Open the file which we want to create a shortcut to.
    Close #1                                            'Why ? Simple : To make sure that the file exists, if it doesn't, an error will occur.
On Error Resume Next                                    'If an error occurs from this point forward it has nothing to do with the file not existing.
    Dim lReturn As Long                                 'This variable will hold the result of out API call.
    
    If optDesktop.Value = True Then
        lReturn = fCreateShellLink("..\..\Desktop", txtName.Text, txtTarget.Text, "")
    Else
        If optPrograms.Value = True Then
            lReturn = fCreateShellLink("", txtName.Text, txtTarget.Text, "")
        Else
            lReturn = fCreateShellLink(txtOther.Text, txtName.Text, txtTarget.Text, "")
        End If
    End If
    
    MsgBox "Shortcut created.", vbOKOnly + vbInformation, "Result"
    
    Exit Sub
    
err:
    MsgBox "The File you entered does not exist.", vbOKOnly + vbCritical, "Error"
End Sub

Private Sub Form_Load()
    optDesktop.Value = True
    txtOther.Enabled = False
    cmdCreateShortcut.Enabled = False                   'We don't have enough info to create the shortcut yet.
End Sub

Private Sub optDesktop_Click()
    txtOther.Enabled = False
    txtOther_Change
End Sub

Private Sub optOther_Click()
    txtOther.Enabled = True
    txtOther_Change
End Sub

Private Sub optPrograms_Click()
    txtOther.Enabled = False
    txtOther_Change
End Sub

Private Sub txtName_Change()
    txtOther_Change                                     'It does the same so why type it twice ?
End Sub

Private Sub txtOther_Change()
    If Len(Trim(txtOther.Text)) = 0 Then                'Make sure txtOther is not empty.
        If optOther.Value = True Then
            cmdCreateShortcut.Enabled = False
        Else
            cmdCreateShortcut.Enabled = True
        End If
    Else
        If Len(Trim(txtTarget.Text)) = 0 Then           'Make sure txtTarget is not empty.
            cmdCreateShortcut.Enabled = False
        Else
            If Len(Trim(txtName.Text)) = 0 Then
                cmdCreateShortcut.Enabled = False
            Else
                cmdCreateShortcut.Enabled = True
            End If
        End If
    End If
    
    'The Len() function returns the length of a string.
    
    'For example : Len("TEST") will return the value 4.
    
    'The Trim() function removes all the spaces at the beginning and end of a string.
    
    'For example : Trim("    TEST    ") will return the value "TEST".
    
    'We can combine these 2 functions.
    
    'For example : Len(Trim("    TEST    ")) will return the value 4.
    
    'This means, that if someone just adds spaces that the program will know.
    
    'For example : Len(Trim("        ")) = 0.
    
    'This way we can check to see if the value entered into the textbox is acceptable.
End Sub

Private Sub txtOther_LostFocus()
    If Mid(txtOther.Text, 1, 1) <> "\" Then             'The destination must start with a \.
        txtOther.Text = "\" & txtOther.Text
    End If
End Sub

Private Sub txtTarget_Change()
    txtOther_Change                                     'It does the same so why type it twice ?
End Sub
