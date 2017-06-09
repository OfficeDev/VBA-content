---
title: Report.TextHeight Method (Access)
keywords: vbaac10.chm13787
f1_keywords:
- vbaac10.chm13787
ms.prod: access
api_name:
- Access.Report.TextHeight
ms.assetid: cac67d4c-e140-06ae-ccbd-961cdee3d087
ms.date: 06/08/2017
---


# Report.TextHeight Method (Access)

The  **TextHeight** method returns the height of a text string as it would be printed in the current font of a **[Report](report-object-access.md)** object.


## Syntax

 _expression_. **TextHeight**( ** _Expr_** )

 _expression_ A variable that represents a **Report** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Expr_|Required|**String**|The text string for which the text height will be determined.|

### Return Value

Single


## Remarks

You can use the  **TextHeight** method to determine the amount of vertical space a text string will require in the current font when the report is formatted and printed. For example, a text string formatted in 9-point Arial will require a different amount of space than one formatted in 12-point Courier. To determine the current font and font size for text in a report, check the settings for the report's **FontName** and **FontSize** properties.

The value returned by the  **TextHeight** method is expressed in terms of the coordinate system in effect for the report, as defined by the **Scale** method. You can use the **ScaleMode** property to determine the coordinate system currently in effect for the report.

If the  _strexpr_ argument contains embedded carriage returns, the **TextHeight** method returns the cumulative height of the lines, including the leading space above and below each line. You can use the value returned by the **TextHeight** method to calculate the necessary space and positioning for multiple lines of text within a report.


## Example

The following example uses the  **TextHeight** and **TextWidth** methods to determine the amount of vertical and horizontal space required to print a text string in the report's current font.

To try this example in Microsoft Access, create a new report. Set the  **OnPrint** property of the Detail section to [Event Procedure]. Enter the following code in the report's module, then switch to Print Preview.




```vb
Private Sub Detail_Print(Cancel As Integer, _ 
 PrintCount As Integer) 
 ' Set unit of measure to twips (default scale). 
 Me.Scalemode = 1 
 ' Print name and font size of report font. 
 Debug.Print "Report Font: "; Me.FontName 
 Debug.Print "Report Font Size: "; Me.FontSize 
 ' Print height and width required for text string. 
 Debug.Print "Text Height (Twips): "; _ 
 Me.TextHeight("Product Report") 
 Debug.Print "Text Width (Twips): "; _ 
 Me.TextWidth("Product Report") 
End Sub
```


## See also


#### Concepts


[Report Object](report-object-access.md)

