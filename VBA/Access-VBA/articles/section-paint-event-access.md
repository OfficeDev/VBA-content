---
title: Section.Paint Event (Access)
keywords: vbaac10.chm14238
f1_keywords:
- vbaac10.chm14238
ms.prod: access
api_name:
- Access.Section.Paint
ms.assetid: f68d981d-8371-cf0d-9da4-063aaa0f0907
ms.date: 06/08/2017
---


# Section.Paint Event (Access)

Occurs when the specified section is redrawn.


## Syntax

 _expression_. **Paint**

 _expression_ A variable that represents a **Section** object.


## Example

The following example shows how to set the  **BackColor** property of a control based on its value.

 **Sample code provided by:** The[Microsoft Access 2010 Programmer's Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.html)





```vb
Private Sub SetControlFormatting()
    If (Me.AvgOfRating >= 8) Then
        Me.AvgOfRating.BackColor = vbGreen
    ElseIf (Me.AvgOfRating >= 5) Then
        Me.AvgOfRating.BackColor = vbYellow
    Else
        Me.AvgOfRating.BackColor = vbRed
    End If
End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    ' size the width of the rectangle
    Dim lngOffset As Long
    lngOffset = (Me.boxInside.Left - Me.boxOutside.Left) * 2
    Me.boxInside.Width = (Me.boxOutside.Width * (Me.AvgOfRating / 10)) - lngOffset
    
    ' do conditional formatting for the control in print preview
    SetControlFormatting
End Sub

Private Sub Detail_Paint()
    ' do conditional formatting for the control in report view
    SetControlFormatting
End Sub
```


## About the Contributors
<a name="AboutContributors"> </a>

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[Section Object](section-object-access.md)

