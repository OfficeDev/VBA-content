---
title: Make a Cell Blink
ms.prod: excel
ms.assetid: 0494fc11-b3d5-4462-aa57-31756cd5a2e7
ms.date: 06/08/2017
---


# Make a Cell Blink

This example shows how to make cell B2 on sheet 1 blink by changing the color and the text back and forth from red to white in the  **StartBlinking** procedure. The **StopBlinking** procedure shows how to stop the blinking by clearing the value of the cell and setting the **[ColorIndex](interior-colorindex-property-excel.md)** property to white.

 **Sample code provided by:** Tom Urtis, [Atlas Programming Management](http://www.atlaspm.com/)



```vb
Option Explicit

Public NextBlink As Double
'The cell that you want to blink
Public Const BlinkCell As String = "Sheet1!B2"

'Start blinking
Private Sub StartBlinking()
    Application.Goto Range("A1"), 1
    'If the color is red, change the color and text to white
    If Range(BlinkCell).Interior.ColorIndex = 3 Then
        Range(BlinkCell).Interior.ColorIndex = 0
        Range(BlinkCell).Value = "White"
    'If the color is white, change the color and text to red
    Else
        Range(BlinkCell).Interior.ColorIndex = 3
        Range(BlinkCell).Value = "Red"
    End If
    'Wait one second before changing the color again
    NextBlink = Now + TimeSerial(0, 0, 1)
    Application.OnTime NextBlink, "StartBlinking", , True
End Sub

'Stop blkinking
Private Sub StopBlinking()
    'Set color to white
    Range(BlinkCell).Interior.ColorIndex = 0
    'Clear the value in the cell
    Range(BlinkCell).ClearContents
    On Error Resume Next
    Application.OnTime NextBlink, "StartBlinking", , False
    Err.Clear
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

MVP Tom Urtis is the founder of Atlas Programming Management, a full-service Microsoft Office and Excel business solutions company in Silicon Valley. Tom has over 25 years of experience in business management and developing Microsoft Office applications, and is the coauthor of "Holy Macro! It's 2,500 Excel VBA Examples." 


