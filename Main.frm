VERSION 5.00
Begin VB.Form Main 
   Caption         =   "Database Utilities"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6285
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu�������� 
      Caption         =   "Tasks"
      Begin VB.Menu mnulDatabaseComparison 
         Caption         =   "Database comparison"
      End
      Begin VB.Menu mnuSeperatorB 
         Caption         =   "-"
      End
      Begin VB.Menu mnu������ 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnulDatabaseComparison_Click()

   Compare.Show

End Sub

Private Sub mnu������_Click()

    End

End Sub
