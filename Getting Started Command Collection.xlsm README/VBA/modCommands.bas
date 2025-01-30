Attribute VB_Name = "modCommands"
Option Explicit

Sub AutoFitColumnWidthSelection()

    If Not Selection Is Nothing Then
        Selection.Columns.AutoFit
    Else
        MsgBox "Please select a range first.", vbExclamation
    End If

End Sub

Sub AutoFitRowHeightSelection()

    If Not Selection Is Nothing Then
        Selection.Rows.AutoFit
    Else
        MsgBox "Please select a range first.", vbExclamation
    End If

End Sub


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Say Hello
' Description:            Open's a message box with "Hello, __ !" where the contents of named range replace __
' Macro Expression:       modCommands.SayHello()
' Generated:              06/06/2024 09:59 AM
'----------------------------------------------------------------------------------------------------
Sub SayHello()
    On Error GoTo ErrorHandler
    MsgBox "Hello, " & ActiveWorkbook.Names("NameForHello").RefersToRange.Value & "!", vbOKOnly, "Greetings from OA Robot      "

    Exit Sub

ErrorHandler:
    MsgBox "We encountered an error saying hello.  Make sure the NameForHello named range exists in the active workbook.", _
     vbOKOnly, "Say Hello Error"
    
End Sub


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Make Cell Pretty
' Description:            Applies the pretty formatting to the active cell
' Macro Expression:       modCommands.MakeCellPretty()
' Generated:              06/06/2024 11:15 AM
'----------------------------------------------------------------------------------------------------
Sub MakeCellPretty()

    With ActiveCell
        .Font.Bold = True
        .Font.Color = -9428616
        .Interior.Color = 16115392
        
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeRight).Weight = xlThick
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeBottom).Weight = xlThick
        
        .Borders(xlEdgeRight).Color = 10384908
        .Borders(xlEdgeTop).Color = 10384908
        .Borders(xlEdgeBottom).Color = 10384908
        .Borders(xlEdgeLeft).Color = 10384908
    End With
        
End Sub


Sub MakeSelectionPretty()

    With Selection
    
        .Font.Color = RGB(21, 96, 130)
        .Font.Name = "Blackadder ITC"
        .Font.Size = 14
        
        .HorizontalAlignment = xlCenterAcrossSelection
        
        .Interior.Color = RGB(218, 242, 208)
        
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeRight).Weight = xlThick
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeBottom).Weight = xlThick
        
        .Borders(xlEdgeRight).Color = RGB(60, 125, 34)
        .Borders(xlEdgeTop).Color = RGB(60, 125, 34)
        .Borders(xlEdgeBottom).Color = RGB(60, 125, 34)
        .Borders(xlEdgeLeft).Color = RGB(60, 125, 34)
    
    End With
        
End Sub
        


Sub MakeTabPretty()

    With ActiveWorkbook.ActiveSheet.Tab
        .Color = RGB(11, 48, 64)
    End With
        
End Sub


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Robot Sticker
' Description:            Sticks a Robot Sticker on the active workbook
' Macro Expression:       modCommands.RobotSticker()
' Generated:              06/13/2024 10:17 AM
'----------------------------------------------------------------------------------------------------
Sub RobotSticker()

    'Insert Logo
    ThisWorkbook.Worksheets("Command Overview").Shapes("Robot Sticker").Copy
    ActiveSheet.Paste
    
    'Move down and right
    Selection.ShapeRange.IncrementLeft 4.5
    Selection.ShapeRange.IncrementTop 3
    
    ActiveCell.Activate
    
ExitHandler:
    Exit Sub

ErrorHandler:
    MsgBox "Unable to insert robot sticker.", vbOKOnly, "Guided Tour Command Collection"

End Sub
