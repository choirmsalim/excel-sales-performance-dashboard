Attribute VB_Name = "FullScreen"
Sub HideAll()
ActiveWindow.DisplayHorizontalScrollBar = False
ActiveWindow.DisplayVerticalScrollBar = False
ActiveWindow.DisplayHeadings = False
ActiveWindow.DisplayWorkbookTabs = False
Application.DisplayStatusBar = False
Application.DisplayFormulaBar = False
Application.ExecuteExcel4Macro "Show.Toolbar(""Ribbon"",False)"
Application.DisplayFullScreen = True
Range("B1").Value = 2
End Sub
Sub ShowAll()
Application.DisplayFullScreen = False
ActiveWindow.DisplayHorizontalScrollBar = True
ActiveWindow.DisplayVerticalScrollBar = True
ActiveWindow.DisplayHeadings = True
ActiveWindow.DisplayWorkbookTabs = True
Application.DisplayStatusBar = True
Application.DisplayFormulaBar = True
Application.ExecuteExcel4Macro "Show.Toolbar(""Ribbon"",True)"
Range("B1").Value = 1
End Sub

Sub Swtchprg()
If Range("B1").Value = 2 Then
Range("B1").Value = 1
Call ShowAll
Else
Range("B1").Value = 2
Call HideAll
End If
End Sub
