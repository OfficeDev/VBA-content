---
title: ToggleButton.FontSize Property (Access)
keywords: vbaac10.chm11724
f1_keywords:
- vbaac10.chm11724
ms.prod: access
api_name:
- Access.ToggleButton.FontSize
ms.assetid: 0175a789-55cb-afeb-33ad-81705983a28d
ms.date: 06/08/2017
---


# ToggleButton.FontSize Property (Access)

You can use the  **FontSize** property to specify the point size for text in the following situations:


- When displaying or printing controls on forms and reports.
    
- When using the  **Print** method on a report.
    

Read/write  **Integer**.


## Syntax

 _expression_. **FontSize**

 _expression_ A variable that represents a **ToggleButton** object.


## Remarks

The  **FontSize** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|8|(Default for all reports and controls except command buttons) The text is 8-point type.|
|10|(Default for command buttons) The text is 10-point type.|
|Other sizes|The text is the indicated size.|
You can set the default for this property by using a control's default control style or the  **DefaultControl** property in Visual Basic.

For reports, you can set this property only in an event procedure or in a macro specified by the  **OnPrint** event property setting.

In Visual Basic, you set the  **FontSize** property by using a numeric expression equal to the desired size of the font. The setting for the **FontSize** property can be between 1 and 127, inclusive.


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

