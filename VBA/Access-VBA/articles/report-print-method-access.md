---
title: Report.Print Method (Access)
keywords: vbaac10.chm13788
f1_keywords:
- vbaac10.chm13788
ms.prod: access
api_name:
- Access.Report.Print
ms.assetid: 6f8523cc-7b17-ec27-e2c9-a7ae3d5a8c3f
ms.date: 06/08/2017
---


# Report.Print Method (Access)

The  **Print** method prints text on a **[Report](report-object-access.md)** object using the current color and font.


## Syntax

 _expression_. **Print**( ** _Expr_** )

 _expression_ A variable that represents a **Report** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Expr_|Required|**String**|The string expressions to print. If this argument is omitted, the  **Print** method prints a blank line. Multiple expressions can be separated with a space, a semicolon (;), or a comma. A space has the same effect as a semicolon.|

### Return Value

Nothing


## Remarks

You can use this method only in a event procedure or macro specified by a section's  **OnPrint** event property setting.

The expressions specified by the  _Expr_ argument are printed on the object starting at the position indicated by the **[CurrentX](report-currentx-property-access.md)** and **[CurrentY](report-currenty-property-access.md)** property settings.

When the  _Expr_ argument is printed, a carriage return is usually appended so that the next **Print** method begins printing on the next line. When a carriage return occurs, the **CurrentY** property setting is increased by the height of the _Expr_ argument (the same as the value returned by the **[TextHeight](report-textheight-method-access.md)** method) and the **CurrentX** property is set to 0.

When a semicolon follows the  _Expr_ argument, no carriage return is appended, and the next **Print** method prints on the same line that the current **Print** method printed on. The **CurrentX** and **CurrentY** properties are set to the point immediately after the last character printed. If the _Expr_ argument itself contains carriage returns, each such embedded carriage return sets the **CurrentX** and **CurrentY** properties as described for the **Print** method without a semicolon.

When a comma follows the  _Expr_ argument, the **CurrentX** and **CurrentY** properties are set to the next print zone on the same line.

When the  _Expr_ argument is printed on a **Report** object, lines that can't fit in the specified position don't scroll. The text is clipped to fit the object.

Because the  **Print** method usually prints with proportionally spaced characters, it's important to remember that there's no correlation between the number of characters printed and the number of fixed-width columns those characters occupy. For example, a wide letter (such as W) occupies more than one fixed-width column, whereas a narrow letter (such as I) occupies less. You should make sure that your tabular columns are positioned far enough apart to accommodate the text you wish to print. Alternately, you can print with a fixed-pitch font (such as Courier) to ensure that each character uses only one column.


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


[Report Object](report-object-access.md)

