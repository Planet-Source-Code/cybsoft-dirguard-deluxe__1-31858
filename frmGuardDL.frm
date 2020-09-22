VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGuard 
   Caption         =   "DirGuard DeLuxe"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11085
   Icon            =   "frmGuardDL.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   11085
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frHiddenItems 
      Caption         =   " Hidden Items "
      Height          =   3375
      Left            =   120
      TabIndex        =   7
      Top             =   6240
      Width           =   10095
      Begin VB.FileListBox HiddenFileList 
         Height          =   1455
         Left            =   6240
         TabIndex        =   35
         Top             =   360
         Width           =   1935
      End
      Begin VB.Timer Timer1 
         Left            =   240
         Top             =   2160
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1575
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   0
         Cols            =   7
         FixedRows       =   0
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "(c) 2002 by Cybsoft"
         Height          =   255
         Left            =   8280
         TabIndex        =   42
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      Begin VB.Frame Frame7 
         Caption         =   " What to detect "
         ForeColor       =   &H00FF0000&
         Height          =   3015
         Left            =   9240
         TabIndex        =   36
         Top             =   120
         Width           =   1695
         Begin VB.CheckBox chkACC 
            Caption         =   "No access"
            Height          =   255
            Left            =   360
            TabIndex        =   43
            Top             =   2520
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkAtt 
            Caption         =   "Attributes"
            Height          =   255
            Left            =   360
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1920
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkDat 
            Caption         =   "Date / Time"
            Height          =   255
            Left            =   360
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1560
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox ChkSiz 
            Caption         =   "Size change"
            Height          =   255
            Left            =   360
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkDel 
            Caption         =   "Deletions"
            Height          =   255
            Left            =   360
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   720
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkAdd 
            Caption         =   "Additions"
            Height          =   255
            Left            =   360
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   360
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00E0E0E0&
            X1              =   0
            X2              =   1680
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            X1              =   0
            X2              =   1680
            Y1              =   1080
            Y2              =   1080
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " Statistics "
         ForeColor       =   &H00FF0000&
         Height          =   3015
         Left            =   6840
         TabIndex        =   17
         Top             =   120
         Width           =   2295
         Begin VB.Line Line1 
            BorderColor     =   &H00000000&
            X1              =   0
            X2              =   2280
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label lblChanges 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1320
            TabIndex        =   30
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Changes found"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Idle"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1320
            TabIndex        =   27
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Status"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label lblRuns 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1320
            TabIndex        =   25
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Runs"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label lblRefresh 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "5 Sec"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1440
            TabIndex        =   23
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Refresh Time"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblFileCount 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1440
            TabIndex        =   21
            Top             =   720
            Width           =   615
         End
         Begin VB.Label lblDirCount 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1440
            TabIndex        =   20
            Top             =   360
            Width           =   615
         End
         Begin VB.Label label8 
            Caption         =   "Guarded Files"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Guarded Dirs"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " Refresh time"
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   2520
         TabIndex        =   15
         Top             =   2400
         Width           =   4215
         Begin MSComctlLib.Slider Slider1 
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   661
            _Version        =   393216
            LargeChange     =   10
            SmallChange     =   5
            Min             =   5
            Max             =   30
            SelStart        =   5
            Value           =   5
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " LogFile "
         ForeColor       =   &H00FF0000&
         Height          =   2175
         Left            =   2520
         TabIndex        =   9
         Top             =   3240
         Width           =   8415
         Begin VB.CheckBox chkScroll 
            Caption         =   "Auto Scroll"
            Height          =   255
            Left            =   2520
            TabIndex        =   44
            Top             =   1800
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CommandButton cmdStart 
            Caption         =   "&Start"
            Height          =   255
            Left            =   5160
            TabIndex        =   34
            ToolTipText     =   "Start Guarding"
            Top             =   1800
            Width           =   975
         End
         Begin VB.CommandButton cmdGuardStop 
            Caption         =   "&Stop"
            Height          =   255
            Left            =   6240
            TabIndex        =   33
            ToolTipText     =   "Stop Guarding"
            Top             =   1800
            Width           =   975
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "&Exit"
            Height          =   255
            Left            =   7320
            TabIndex        =   28
            ToolTipText     =   "Exit program"
            Top             =   1800
            Width           =   975
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Print"
            Height          =   255
            Left            =   720
            TabIndex        =   14
            ToolTipText     =   "Print the logfile"
            Top             =   1800
            Width           =   495
         End
         Begin VB.CommandButton cmdClearLoG 
            Caption         =   "Clear"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            ToolTipText     =   "Clear the logfile"
            Top             =   1800
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Height          =   255
            Left            =   1320
            TabIndex        =   12
            ToolTipText     =   "Save the logfile"
            Top             =   1800
            Width           =   495
         End
         Begin VB.CommandButton cmdInsert 
            Caption         =   "Add"
            Height          =   255
            Left            =   1920
            TabIndex        =   11
            ToolTipText     =   "Add a commentline to logfile"
            Top             =   1800
            Width           =   495
         End
         Begin RichTextLib.RichTextBox rtbChangedfiles 
            Height          =   1455
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   240
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   2566
            _Version        =   393217
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            Appearance      =   0
            TextRTF         =   $"frmGuardDL.frx":030A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin VB.Frame Frame3 
         Caption         =   " Guarded directorys "
         ForeColor       =   &H00FF0000&
         Height          =   2175
         Left            =   2520
         TabIndex        =   5
         Top             =   120
         Width           =   4215
         Begin VB.CommandButton cmdOK 
            Caption         =   "Add to Guardlist"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CommandButton cmdClearDirs 
            Caption         =   "Clear Guardlist"
            Height          =   255
            Left            =   2280
            TabIndex        =   31
            Top             =   1680
            Width           =   1815
         End
         Begin VB.ListBox lstGuardDirs 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   1395
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Navigation "
         ForeColor       =   &H00FF0000&
         Height          =   5295
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2295
         Begin VB.FileListBox lstFiles 
            Appearance      =   0  'Flat
            Height          =   2175
            Hidden          =   -1  'True
            Left            =   120
            System          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   2940
            Width           =   2055
         End
         Begin VB.DirListBox lstMap 
            Appearance      =   0  'Flat
            Height          =   2115
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   2055
         End
         Begin VB.DriveListBox Drivestation 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2055
         End
      End
   End
End
Attribute VB_Name = "frmGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NoChange As Boolean                     ' Signal changes
Dim GuardStart As Boolean                   ' Signals Guard-start
Private Sub cmdExit_Click()
Timer1.Interval = 0
Unload Me
End
End Sub

Private Sub cmdGuardStop_Click()
lblStatus.Caption = "End run"
lblStatus.Refresh
Timer1.Interval = 0
GuardStart = False

cmdOK.Enabled = True
cmdClearDirs.Enabled = True
cmdClearLoG.Enabled = True
cmdPrint.Enabled = True
cmdSave.Enabled = True
Drivestation.Enabled = True
lstMap.Enabled = True

If lstGuardDirs.ListCount > 0 Then                  ' dont if no dirs where guarded
    Call LogFooter
    Call RefreshData
End If
lblStatus.Caption = "Idle"
Me.Refresh

If chkScroll.Value = 1 Then
    rtbChangedfiles.SelStart = Len(rtbChangedfiles) + 1         ' autoscroll
End If

End Sub
Private Sub cmdStart_Click()
If lstGuardDirs.ListCount = 0 Then Exit Sub ' nothing to guard
GuardStart = True
cmdOK.Enabled = False
cmdClearDirs.Enabled = False
cmdClearLoG.Enabled = False
cmdPrint.Enabled = False
cmdSave.Enabled = False
Drivestation.Enabled = False
lstMap.Enabled = False

updatetime = Slider1.Value                  ' whats the value of the slider ?
lblRefresh.Caption = CStr(updatetime) & " Sec"
Timer1.Interval = (updatetime * 1000)
lblRuns.Caption = "0"
lblStatus.Caption = "Running"
Call LogHeader
Call RefreshData
Me.Refresh
End Sub

Private Sub Timer1_Timer()
NoChange = True
If GuardStart = False Then Exit Sub
lblStatus.Caption = "Check"
lblStatus.Refresh

Call CheckFilesAdd                                      ' are files added
Call CheckFilesDel                                      ' are files deleted

'                                                   Get stored data
For ChangeCounter = 0 To MSFlexGrid1.Rows - 1              ' number of guarded files
    MSFlexGrid1.Row = ChangeCounter
    
    MSFlexGrid1.Col = 0
    StoredFile = MSFlexGrid1.Text
    MSFlexGrid1.Col = 1
    StoredFileName = MSFlexGrid1.Text               ' bestandsnaam
    MSFlexGrid1.Col = 2
    StoredDateTime = MSFlexGrid1.Text                ' date en time
    MSFlexGrid1.Col = 3
    StoredSize = Val(MSFlexGrid1.Text)               ' Size
    MSFlexGrid1.Col = 4
    StoredAttrib = Val(MSFlexGrid1.Text)             ' attrib

' =============================================== Get actural data ==============
    Position = InStrRev(StoredFile, "\")            ' visualisation
    ToMark = Left(StoredFile, Position - 1)         ' Find last "\" in filenamestring
    For MarkCounter = 0 To lstGuardDirs.ListCount - 1
    MarkDir = lstGuardDirs.List(MarkCounter)        ' find active directory
    If MarkDir = ToMark Then
        lstGuardDirs.Selected(MarkCounter) = True   ' Visualisation
    End If
    Next MarkCounter
    
    LetsFind = Dir(StoredFile, vbDirectory)         ' first check if file still there !
        If Len(LetsFind) = 0 Then                   ' Not found / no longer there
            GoTo WasMissing                         ' = next MainCounter
        End If
        
        ActualDT = (FileDateTime(StoredFile))       ' get DateAndTime
        ActualSize = Val((FileLen(StoredFile)))     ' Get Size
        ActualAttrib = Val(GetAttr(StoredFile))    ' get Attribute
        
' =================================== Compair and log ============================

        If StoredSize <> ActualSize Then
         If ChkSiz.Value = 0 Then GoTo chkTime              ' skip
         
            NoChange = False
            WeHave = SizeResult(Val(StoredSize), Val(ActualSize))
            rtbChangedfiles.Text = rtbChangedfiles.Text & Date & "  " & Time & "    " & StoredFile & " Size changed now : " & WeHave & vbCrLf
            lblChanges.Caption = Val(lblChanges.Caption) + 1
        End If
        
chkTime:
        
        If StoredDateTime <> CStr(ActualDT) Then              ' change Date / time
        
            If chkDat.Value = 0 Then GoTo chkAtri               ' skip
            
            NoChange = False
            rtbChangedfiles.Text = rtbChangedfiles.Text & Date & "  " & Time & "    " & StoredFile & " Date/Time changed" & vbCrLf
            lblChanges.Caption = Val(lblChanges.Caption) + 1
        End If
chkAtri:
            
        If StoredAttrib <> ActualAttrib Then
            If chkAtt.Value = 0 Then GoTo WasMissing              ' skip
            
            NoChange = False
            NewAt = CAttr(Val(ActualAttrib))                      ' convert to tekst
            rtbChangedfiles.Text = rtbChangedfiles.Text & Date & "  " & Time & "    " & StoredFile & " Attribute changed in : " & NewAt & vbCrLf
            lblChanges.Caption = Val(lblChanges.Caption) + 1
        End If
WasMissing:
Next ChangeCounter
If NoChange = False Then rtbChangedfiles.Refresh            ' only if something happend
lblRuns.Caption = Val(lblRuns.Caption) + 1
lblRuns.Refresh
lblStatus.Caption = "Running"

If chkScroll.Value = 1 Then
    rtbChangedfiles.SelStart = Len(rtbChangedfiles) + 1         ' autoscroll
End If
lblStatus.Refresh
Call RefreshData
End Sub
Private Sub CheckFilesAdd()

If chkAdd.Value = 0 Then Exit Sub

For CollectCounter = 0 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.Row = CollectCounter
    MSFlexGrid1.Col = 0
    Filenames = MSFlexGrid1.Text
    FileString = FileString & " " & Filenames
Next CollectCounter                                         ' guarded filenames

For Counter = 0 To lstGuardDirs.ListCount - 1
    lstGuardDirs.Selected(Counter) = True                       'Visualisation
    HiddenFileList.Path = lstGuardDirs.List(Counter)              ' directories

    HiddenFileList.Refresh
    
        For SubCounter = 0 To HiddenFileList.ListCount - 1  ' Files in directories
            Diritem = HiddenFileList.List(SubCounter)           ' a file
            Diritem = HiddenFileList.Path & "\" & Diritem
            MSFlexGrid1.Col = 0
            
                If InStr(FileString, Diritem) = 0 Then                   ' file added
                    rtbChangedfiles.Text = rtbChangedfiles.Text & Date & "  " & Time & "    " & Diritem & "  " & "was added" & vbCrLf
                    lblChanges.Caption = Val(lblChanges.Caption) + 1
                End If
        Next SubCounter
Next Counter
End Sub
Private Sub CheckFilesDel()

If chkDel.Value = 0 Then Exit Sub

For CollectCounter = 0 To lstGuardDirs.ListCount - 1

    HiddenFileList.Path = lstGuardDirs.List(CollectCounter)
    lstGuardDirs.Selected(Counter) = True                   ' visualisation
    If HiddenFileList.ListCount <= 0 Then Exit Sub          'nothing there
    HiddenFileList.Refresh
        For SubCounter = 0 To HiddenFileList.ListCount
            Filenames = HiddenFileList.List(SubCounter)
            Filenames = HiddenFileList.Path & "\" & Filenames
            FileString = FileString & " " & Filenames
        Next SubCounter
Next CollectCounter

MSFlexGrid1.Col = 0
For Counter = 0 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.Row = Counter
    Checkitem = MSFlexGrid1.Text
        
        If InStr(FileString, Checkitem) = 0 Then            ' deleted
               rtbChangedfiles.Text = rtbChangedfiles.Text & Date & "  " & Time & "    " & Checkitem & "  " & "was deleted" & vbCrLf
               lblChanges.Caption = Val(lblChanges.Caption) + 1
        End If
 Next Counter
End Sub
Private Sub RefreshData()
lblStatus.Caption = "Refresh"
lblStatus.Refresh
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 0
MSFlexGrid1.Cols = 5

' Of course you can keep the database in the Flexgrid and just add or remove
' items from it but i choose to update / refresh it completely every run.
' This makes the code much more simple.

lblFileCount.Caption = "0"

For FillCounter = 0 To lstGuardDirs.ListCount - 1
    lstGuardDirs.Selected(FillCounter) = True       ' Visualisation
    FilePad = lstGuardDirs.List(FillCounter)
    HiddenFileList.Path = FilePad
    HiddenFileList.Refresh
    
    For SubCounter = 0 To HiddenFileList.ListCount - 1
        If HiddenFileList.ListCount <= 0 Then GoTo SkipFile ' nothing there
        Item = HiddenFileList.List(SubCounter)

        ' get the fileinfo

        Totalen = FilePad & "\" & Item              ' gehele filename   Col = 1
                                                    ' bestandnaam       Col = 2
        Firstcheck = Dir(Totalen, vbDirectory)      ' does it exists
            If Len(Firstcheck) = 0 Then             ' can not be accesed
                If chkACC.Value = 1 Then
                    rtbChangedfiles.Text = rtbChangedfiles.Text & "<GUARD NOTE>    " & Totalen & "   " & "can not be accesed (NOT guarded) !" & vbCrLf
                End If
                GoTo SkipFile
            End If
            
        TimeStamp = (FileDateTime(Totalen))         ' TimeAndDAte       Col = 3
        FileSize = (FileLen(Totalen))               ' FileSize          Col = 4
        Attrib = (GetAttr(Totalen))                 ' Attribute         Col = 5
        
        MSFlexGrid1.AddItem Totalen & Chr(9) & Item & Chr(9) & TimeStamp & Chr(9) & FileSize & Chr(9) & Attrib, SubCounter
        lblFileCount.Caption = MSFlexGrid1.Rows
    Next SubCounter
SkipFile:
Next FillCounter

' CLEARING THE APPENDED EMPTY GRID-FIELDS dont know why they are still there but
' hey...now they are gone . .
lstGuardDirs.Selected(FillCounter - 1) = False
ClearAppend = Val(lblFileCount.Caption)
For ClearCounter = MSFlexGrid1.Rows To (MSFlexGrid1.Rows + 1) Step -1
   MSFlexGrid1.RemoveItem (ClearCounter)
Next ClearCounter

MSFlexGrid1.Refresh
If GuardStart = True Then
    lblStatus.Caption = "Running"
 Else
    lblStatus.Caption = "End run"
End If

lblStatus.Refresh
End Sub


Private Sub Form_Activate()
Me.Caption = " Directory Guard V" & App.Major & "." & App.Minor & " DeLuxe"
GuardStart = False
NoChange = True
End Sub
Private Sub Form_Resize()

    If WindowState = conMinimized Then
        Caption = " Directory Guard V" & App.Major & "." & App.Minor & " DeLuxe"
    Else
        Caption = "  DirGuard"
        
    End If
End Sub
Private Sub cmdOK_Click()
FilePad = lstFiles.Path
MSFlexGrid1.Cols = 5
' aantal cols is aantal colomen welke je gebruikt invoer gescheiden door chr(9)

guarded = False
' Check for double if not then add

If UCase(lstMap.Path) = "C:\" Then
      MsgBox "DONT MONITOR C:\ ROOT-DIRECTORIE !!!!" & vbCrLf & "The program will probably crash !", vbExclamation + vbOKOnly, "FATAL ERROR"
    Exit Sub                                           'NO ROOT_DIRS !!!!!
End If

For Counter = 0 To lstGuardDirs.ListCount - 1           'No double
    StoredItem = lstGuardDirs.List(Counter)
    If StoredItem = lstMap.Path Then guarded = True
Next Counter

If guarded = False Then
lstGuardDirs.AddItem (lstMap.Path)
lblDirCount.Caption = lstGuardDirs.ListCount
lblFileCount.Caption = Val(lblFileCount.Caption) + lstFiles.ListCount
End If


For Counter = 0 To lstFiles.ListCount - 1
Item = lstFiles.List(Counter)

' get the fileinfo
Totalen = FilePad & "\" & Item              ' gehele filename   Col = 1
                                            ' bestandnaam       Col = 2
Firstcheck = Dir(Totalen, vbDirectory)      ' does it exists
    If Len(Firstcheck) = 0 Then             ' can not be accessed
        If chkACC.Value = 1 Then
                rtbChangedfiles.Text = rtbChangedfiles.Text & rtbChangedfiles.Text & "<GUARD NOTE>    " & Totalen & "   " & "can not be accesed (NOT guarded) !" & vbCrLf
        End If
        GoTo SkipFile
    End If
TimeStamp = (FileDateTime(Totalen))         ' TimeAndDAte       Col = 3
FileSize = (FileLen(Totalen))               ' FileSize          Col = 4
Attrib = (GetAttr(Totalen))                 ' Attribute         Col = 5
        
' add to flexgrid

MSFlexGrid1.AddItem Totalen & Chr(9) & Item & Chr(9) & TimeStamp & Chr(9) & FileSize & Chr(9) & Attrib
SkipFile:
Next Counter

TimeAdjust = MSFlexGrid1.Rows               ' Adjusting timer to compensate
    
    Select Case TimeAdjust                  ' large file-amounts
    
    Case Is > 100                           ' > 100 guardfiles
        Slider1.Min = 10
        Slider1.Max = 40
    Case Is > 200
        Slider1.Min = 15                    ' > 200 guardfiles
        Slider1.Max = 45
    Case Is > 300
        Slider1.Min = 20                    ' > 300 guardFiles
        Slider1.Max = 60
    End Select

End Sub

Private Sub Drivestation_Change()                   ' we change drive
    On Error GoTo error                             ' in case a disk is not available
    lstMap.Path = Drivestation.Drive                ' set the directory for the map-list
    Exit Sub
error:                                              ' disk was not available
    Dim answer As Integer
    answer = MsgBox(Err.Description, 5, "Device error !")
    If Annswer = 4 Then Resume                      ' they pressed ok
End Sub
Private Sub lstMap_Change()                         ' the map-information                                     ' if not stopped...stop it now
lstMap.Refresh                                      ' refresh it
lstFiles.Path = lstMap.Path                         ' set the map for the filelisting
End Sub
Private Sub Slider1_Change()                ' change the update-time for the timer
Dim updatetime As Integer                   ' there is no need for all declarations
                                            ' its best to get used to it and use them
updatetime = Slider1.Value                  ' whats the value of the slider ?
If updatetime < Slider1.Min Then            ' is it smaler than the minimum that is
    updatetime = Slider1.Min                ' set by the filecount
End If
lblRefresh.Caption = CStr(updatetime) & " Sec"
lblRefresh.Refresh

If GuardStart = True Then                   ' unless we guard . . .  !!
    Timer1.Interval = (updatetime * 1000)   ' make seconds from milliseconds
    rtbChangedfiles.Text = rtbChangedfiles.Text & Date & "  " & Time & "  Refreshtime changed now : " & lblRefresh.Caption & vbCrLf
End If

If chkScroll.Value = 1 Then
    rtbChangedfiles.SelStart = Len(rtbChangedfiles) + 1         ' autoscroll
End If

End Sub

Private Sub cmdClearDirs_Click()
lstGuardDirs.Clear
MSFlexGrid1.Clear
End Sub
Private Sub LogHeader()
Message1 = "Directory Guard V" & App.Major & "." & App.Minor & " DeLuxe" & vbCrLf
Message2 = "Guarding started on : " & Date & " At : " & Time & " Refresh-Time : " & lblRefresh.Caption & vbCrLf
Message4 = "=========================================================================" & vbCrLf
rtbChangedfiles.Text = rtbChangedfiles.Text & Message1 & Message2 & Message4
End Sub
Private Sub LogFooter()
Message1 = "=========================================================================" & vbCrLf
Message2 = "Guarding stopped on : " & Date & " At : " & Time & vbCrLf
Message3 = "=========================================================================" & vbCrLf
rtbChangedfiles.Text = rtbChangedfiles.Text & Message1 & Message2 & Message3
End Sub
Private Sub cmdPrint_Click()
rtbChangedfiles.SelPrint (Printer.hDC)  ' could be better but it works
Printer.EndDoc                          ' just print the lof and eject page
cmdStart.SetFocus
End Sub
Private Sub cmdClearLoG_Click()
rtbChangedfiles = ""
lblChanges.Caption = "0"
cmdStart.SetFocus
End Sub
Private Sub cmdSave_Click()
frmSave.Show
End Sub
Private Sub cmdInsert_Click()
FrmNote.Show
cmdStart.SetFocus
If chkScroll.Value = 1 Then
    rtbChangedfiles.SelStart = Len(rtbChangedfiles) + 1         ' autoscroll
End If
End Sub
