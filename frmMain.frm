VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "File Merger By Mike"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3480
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFiles 
      Height          =   1620
      Left            =   75
      OLEDropMode     =   1  'Manual
      TabIndex        =   6
      Top             =   75
      Width           =   3315
   End
   Begin VB.PictureBox PercentBox 
      Height          =   275
      Left            =   75
      ScaleHeight     =   210
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   4
      Top             =   2250
      Width           =   3315
      Begin VB.Label lblPercent 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         Height          =   270
         Left            =   1350
         TabIndex        =   7
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Bar 
         BackColor       =   &H0080C0FF&
         Height          =   270
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   60
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   390
      Left            =   1275
      TabIndex        =   3
      Top             =   2775
      Width           =   915
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear List"
      Height          =   390
      Left            =   75
      TabIndex        =   2
      Top             =   2775
      Width           =   915
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Left            =   75
      TabIndex        =   1
      Text            =   "C:\windows\desktop\test.mp3"
      Top             =   1800
      Width           =   3315
   End
   Begin VB.CommandButton cmdMerge 
      Caption         =   "Merge"
      Default         =   -1  'True
      Height          =   390
      Left            =   2475
      TabIndex        =   0
      Top             =   2775
      Width           =   915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'File Merger By Mike Canejo
'mikecanejo@hotmail.com


Option Explicit
'Makes all vars error upon compilation
'if not declared before they are used.




Private TheFiles(250) As String
'String array to hold up to 250 max file paths.
'Used this for the drag and drop feature to save the paths.


Private Sub cmdMerge_Click()
Dim i As Long, x As Long
Dim SavedSpot As Long
Dim theByte() As Byte
Dim Length As Long
'Declares variables
'I used '()' in the byte array because
'i don't know what the file size will be
'so this allows you to resize it later once
'i know. (ReDim theByte(10) would resize it to 10)


lstFiles.Enabled = False
txtOutput.Enabled = False
cmdClear.Enabled = False
Bar.Visible = True
'To give the right effect of the progress bar
'and to make sure the merging process can't be interupted
'by a change of the location or by dragging more files into the listbox.

SavedSpot = 1
'A type long variable to save the position in the
'output file to put the next file thus merging.


For i = 0 To lstFiles.ListCount - 1

    Me.Caption = "File Merger By Mike - [File:" & i + 1 & " of " & lstFiles.ListCount & "]"
    'Displays the status in the form's caption
    
    Length = FileLen(TheFiles(i))
    'Puts the file length in the type long var.
    
    ReDim theByte(Length - 1)
    'This will resize 'theByte' array to the file size
    'and each spot in the array will hold 1 byte.
    
    Open TheFiles(i) For Binary Access Read As #1 'Opens the file to read from it in binary
        
        Get #1, , theByte()
        'Loads the whole file into 'theByte' array
    
    Close #1

    Open txtOutput For Binary As #1  'Opens the file to write to it in binary
    
        Put #1, SavedSpot, theByte()
        'Puts the next file in the listbox in the right position
        'using the 'SavedSpot' variable.
        
    Close #1

    If Bar.Width > 49 Then
        lblPercent.ForeColor = &H80&
        'Just an effect for the progress bar when it exceeds 49%.
    End If
    
    Bar.Width = Int((100 / lstFiles.ListCount) * i + 1)
    'Sets the width to the current percent of the status
    'using the picturebox at the Scaled width of 100. (100%)
    
    lblPercent.Caption = Int((100 / lstFiles.ListCount) * i + 1) & "%"
    'Sets the percent left in the merge process to the 'lblPercent' label
    
    SavedSpot = SavedSpot + Length
    'Saves the spot in the output file to put the next
    'file in the spot it left off at.
    
    DoEvents
    'Provides a convenient way to allow a task to be canceled
    'thus allowing the form not to freeze up.
Next i


Bar.Width = 100
'Just incase the progress bar doesn't come out to 100$, SET IT!!!!!  ;]

lblPercent.Caption = "100%"
'Same reason as above, sometimes it could come to 93% from a odd number.

MsgBox "DONE!", vbInformation, "Alert"
'Tells you when it's over, hurray!

lblPercent.ForeColor = &H80000012
'Resets the forecolor for the next merge.

lblPercent.Caption = "0%"
'Resets back to 0%

lstFiles.Enabled = True
txtOutput.Enabled = True
cmdClear.Enabled = True
'Enabled the controls above.

Bar.Visible = False
Bar.Width = 0
'Hides the bar because vb makes the width '15' and this will be visible,
'we don't want that so we'll hide it until the next merge. We set the width to
'0 also for the next time you merge something.
End Sub


Private Sub cmdClose_Click()
    Close #1    'Upon exit, make sure the file is closed.
    End         'Terminates the program.
End Sub

Private Sub cmdClear_Click()
    lstFiles.Clear       'Clears the listbox, i hope you learned this by now ;]
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width / 2 - Me.Left / 2), (Screen.Height / 2 - Me.Height / 2)
    'Centers the form in the center of the users screen.
    
    txtOutput = GetSetting("1", "1", "1", "C:\windows\desktop\test.mp3")
    'Loads the last typed location to put the output but if its the first
    'time loading, load my default path i came up with.
End Sub

Private Sub lstFiles_DblClick()
    lstFiles.Clear       'Clears the listbox, i hope you learned this by now ;]
End Sub

Private Sub lstFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo EndIt                             'If theres an error from dragging and dropping the files goto EndIt.
Dim i As Integer, z As Integer                  'Declares the vars.
z = lstFiles.ListCount - 1                      'Saves the total number of items to load into 'TheFiles' array.
    For i = 1 To 250                            'Up to 250 files can be dragged in the listbox at a time, increase for slower performance.
        lstFiles.AddItem GetFile(Data.Files(i)) 'Adds the file name of the file's path using the GetFile function.
        TheFiles(z + i) = Data.Files(i)         'Adds 'i' to the z var to hold its place and increase it by i.
    Next i                                      'Uhh...Next i buddy ;]
EndIt:                                          'Incase of an error, go here.
    lstFiles.ListIndex = lstFiles.ListCount - 1 'Highlites the last item in the listbox.
End Sub

Private Function GetFile(ThePath) As String
Dim i As Integer
    If InStr(1, ThePath, "\", vbTextCompare) <> 0 Then 'Checks if there's a backslash in the string.
        For i = Len(ThePath) To 1 Step -1              'Goes backwards in the string to find the first existance of the backslash.
            If Mid(ThePath, i, 1) = "\" Then           'When the backslash is found.
                GetFile = Mid(ThePath, i + 1)          'Put the file name in the function string.
                Exit For                               'End the for loop.
            End If
        Next i
    End If
End Function

Private Sub txtOutput_Change()
    SaveSetting "1", "1", "1", txtOutput
    'Wrote this up really quick to save what you last
    'typed for a location to saved the output to and
    'then be able to load it.
End Sub
