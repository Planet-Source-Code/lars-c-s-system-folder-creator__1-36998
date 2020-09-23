Attribute VB_Name = "status"
Option Explicit


Public Sub GoToWeb(WhatURL)
Dim Success As Long

Success = ShellExecute(0&, vbNullString, WhatURL, vbNullString, "C:\", SW_SHOWNORMAL)

End Sub
