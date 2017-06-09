---
title: ToggleButton.FontName Property (Access)
keywords: vbaac10.chm11723
f1_keywords:
- vbaac10.chm11723
ms.prod: access
api_name:
- Access.ToggleButton.FontName
ms.assetid: 7b1d51d8-5307-1446-344a-f406f2758a36
ms.date: 06/08/2017
---


# ToggleButton.FontName Property (Access)

You can use the  **FontName** property to specify the font for text in the following situations:


- When displaying or printing controls on forms and reports.
    
- When using the  **[Print](report-print-method-access.md)** method on a report.
    

Read/write  **String**.


## Syntax

 _expression_. **FontName**

 _expression_ A variable that represents a **ToggleButton** object.


## Remarks

The  **FontName** property setting is the name of the font that the text is displayed in.

You can set the default for this property by using a control's default control style or the  **DefaultControl** property in Visual Basic.

For reports, you can set this property only in an event procedure or in a macro specified by the  **OnPrint** event property setting.

Font availability depends on your system and printer. If you select a font that your system can't display or that isn't installed, Windows substitutes a similar font.


## Example

The following example uses the  **Print** method to display text on a report named Report1. It uses the **TextWidth** and **TextHeight** methods to center the text vertically and horizontally.


```vb
Private Sub Detail_Format(Cancel As Integer, _ 
 FormatCount As Integer) 
 Dim rpt as Report 
 Dim strMessage As String 
 Dim intHorSize As Integer, intVerSize As Integer 
 
 Set rpt = Me 
 strMessage = "DisplayMessage" 
 With rpt 
 'Set scale to pixels, and set FontName and 
 'FontSize properties. 
 .ScaleMode = 3 
 .FontName = "Courier" 
 .FontSize = 24 
 End With 
 ' Horizontal width. 
 intHorSize = Rpt.TextWidth(strMessage) 
 ' Vertical height. 
 intVerSize = Rpt.TextHeight(strMessage) 
 ' Calculate location of text to be displayed. 
 Rpt.CurrentX = (Rpt.ScaleWidth/2) - (intHorSize/2) 
 Rpt.CurrentY = (Rpt.ScaleHeight/2) - (intVerSize/2) 
 ' Print text on Report object. 
 Rpt.Print strMessage 
End Sub
```


## See also


#### Concepts


[ToggleButton Object](togglebutton-object-access.md)

