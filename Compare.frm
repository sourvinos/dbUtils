VERSION 5.00
Object = "{839D0F5D-B7D7-41B7-A3B4-85D69300B8C1}#98.0#0"; "iGrid300_10Tec.ocx"
Begin VB.Form Compare 
   Caption         =   "Database Comparison"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10380
   Icon            =   "Compare.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame myFrame 
      Caption         =   "Frames"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   2550
      TabIndex        =   4
      Top             =   6900
      Width           =   3690
   End
   Begin VB.PictureBox Seperator 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   90
      Left            =   1050
      MousePointer    =   7  'Size N S
      ScaleHeight     =   90
      ScaleWidth      =   5865
      TabIndex        =   3
      Top             =   6375
      Width           =   5865
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear grid"
      Height          =   465
      Left            =   2025
      TabIndex        =   2
      Top             =   5475
      Width           =   1890
   End
   Begin iGrid300_10Tec.iGrid grdGrid 
      Height          =   5340
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9419
      BorderStyle     =   1
      DefaultRowHeight=   19
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Condensed"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483631
      FrozenCols      =   1
      RowMode         =   -1  'True
   End
   Begin VB.CommandButton cmdCompare 
      Caption         =   "Start"
      Height          =   465
      Left            =   75
      TabIndex        =   0
      Top             =   5475
      Width           =   1890
   End
End
Attribute VB_Name = "Compare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lngOldY As Long
Private blnIsMoving As Boolean

Private Sub cmdClear_Click()

    grdGrid.Clear True
    cmdClear.Enabled = False
    cmdCompare.Enabled = True

End Sub

Private Sub cmdCompare_Click()

    'Local μεταβλητές
    Dim lngRow As Long
    Dim lngCol As Long
    Dim dBase As Database
    Dim myFile As String
    Dim myName As String
    Dim myPath As String
    Dim intTables As Integer
    Dim intFields As Integer
    Dim intIndexes As Integer
    
    'Αρχικές τιμές
    myPath = App.Path
    myName = Dir(myPath, vbNormal)
    myFile = Dir(strDataDirectory & "\*.mdb", vbNormal)
    grdGrid.Redraw = False
    lngRow = 1
    
    Do While myFile <> ""
        grdGrid.AddCol myFile, myFile, lwidth:=300
        Set dBase = DBEngine.OpenDatabase(strDataDirectory & "\" & myFile, False, True)
        For intTables = 0 To dBase.TableDefs.Count - 1
            If dBase.TableDefs(intTables).Attributes = 0 Then
                If grdGrid.RowCount < lngRow Then
                    grdGrid.AddRow
                    lngRow = grdGrid.RowCount
                End If
                grdGrid.CellValue(lngRow, grdGrid.ColCount) = dBase.TableDefs(intTables).Name
                lngRow = lngRow + 1
                For intFields = 0 To dBase.TableDefs(intTables).Fields.Count - 1
                    If grdGrid.RowCount < lngRow Then
                        grdGrid.AddRow
                        lngRow = grdGrid.RowCount
                    End If
                    grdGrid.CellValue(lngRow, grdGrid.ColCount) = "     " & _
                        dBase.TableDefs(intTables).Fields(intFields).Name & _
                        " | " & dBase.TableDefs(intTables).Fields(intFields).Size
                    lngRow = lngRow + 1
                Next intFields
                For intIndexes = 0 To dBase.TableDefs(intTables).Indexes.Count - 1
                    If grdGrid.RowCount < lngRow Then
                        grdGrid.AddRow
                        lngRow = grdGrid.RowCount
                    End If
                    grdGrid.CellValue(lngRow, grdGrid.ColCount) = "     " & _
                        dBase.TableDefs(intTables).Indexes(intIndexes).Name & " " & _
                        dBase.TableDefs(intTables).Indexes(intIndexes).Primary & " " & _
                        dBase.TableDefs(intTables).Indexes(intIndexes).Fields
                    lngRow = lngRow + 1
                Next intIndexes
            End If
        Next intTables
        dBase.Close
        myFile = Dir
        lngRow = 1
    Loop
    
    grdGrid.AddCol
    For lngRow = 1 To grdGrid.RowCount
        For lngCol = 2 To grdGrid.ColCount - 1
            If grdGrid.CellValue(lngRow, lngCol) <> grdGrid.CellValue(lngRow, lngCol - 1) Then
                grdGrid.CellForeColor(lngRow, grdGrid.ColCount) = vbRed
                grdGrid.CellValue(lngRow, grdGrid.ColCount) = "Error!"
            End If
        Next lngCol
    Next lngRow
    
    grdGrid.Redraw = True
    cmdCompare.Enabled = False
    cmdClear.Enabled = True

End Sub

Private Sub Form_Load()

    strDataDirectory = App.Path & "\Data"
    
    Me.WindowState = vbMaximized
    Me.Width = Screen.Width
    Me.Height = Screen.Height - 1250
    
    grdGrid.Width = Me.Width - 150
    grdGrid.Height = Me.Height - 150
    grdGrid.ForeColor = &H400040
    
    grdGrid.Header.Font = "Ubuntu Condensed"
    grdGrid.Header.Font.Size = 11
    
    cmdCompare.Top = grdGrid.Height + 150
    cmdClear.Top = grdGrid.Height + 150
    
    cmdCompare.Enabled = True
    cmdClear.Enabled = False
    
    Seperator.Left = 0
    Seperator.Width = Me.Width
    Seperator.Top = LoadSeperatorPositionFromRegistry
    
    PositionControls

End Sub

Private Sub Seperator_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        lngOldY = Y
        blnIsMoving = True
    End If
    
End Sub

Private Sub Seperator_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If blnIsMoving Then
        Seperator.Top = Seperator.Top - (lngOldY - Y)
        PositionControls
    End If
    
End Sub

Private Sub Seperator_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    blnIsMoving = False
    
    SaveSetting App.EXEName, "Seperator Position", "Top", Seperator.Top

End Sub

Private Function LoadSeperatorPositionFromRegistry()

    LoadSeperatorPositionFromRegistry = GetSetting(App.EXEName, "Seperator Position", "Top")

End Function

Private Function PositionControls()

    myFrame.Top = Seperator.Top + 200
    myFrame.Height = grdGrid.Height - Seperator.Top - 200

End Function
