---
title: AboveAverage Object (Excel)
keywords: vbaxl10.chm823072
f1_keywords:
- vbaxl10.chm823072
ms.prod: excel
api_name:
- Excel.AboveAverage
ms.assetid: dd4ea82f-7986-5d6f-2b0e-fe0ca38226e2
ms.date: 06/08/2017
---


# AboveAverage Object (Excel)

Represents an above average visual of a conditional formatting rule. Applying a color or fill to a range or selection to help you see the value of a cells relative to other cells.


## Remarks

All conditional formatting objects are contained within a  **[FormatConditions](formatconditions-object-excel.md)** collection object, which is a child of a **[Range](range-object-excel.md)** collection. You can create an above average formatting rule by using either the **[Add](formatconditions-add-method-excel.md)** or **[AddAboveAverage](formatconditions-addaboveaverage-method-excel.md)** method of the **FormatConditions** collection.


## Example

The following example builds a dynamic data set and applies color to the above average values through conditional formatting rules.


```
Sub AboveAverageCF() 
 
' Building data for Melanie 
 Range("A1").Value = "Name" 
 Range("B1").Value = "Number" 
 Range("A2").Value = "Melanie-1" 
 Range("A2").AutoFill Destination:=Range("A2:A26"), Type:=xlFillDefault 
 Range("B2:B26").FormulaArray = "=INT(RAND()*101)" 
 Range("B2:B26").Select 
 
' Applying Conditional Formatting to items above the average. Should appear green fill and dark green font. 
 Selection.FormatConditions.AddAboveAverage 
 Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority 
 Selection.FormatConditions(1).AboveBelow = xlAboveAverage 
 With Selection.FormatConditions(1).Font 
 .Color = -16752384 
 .TintAndShade = 0 
 End With 
 With Selection.FormatConditions(1).Interior 
 .PatternColorIndex = xlAutomatic 
 .Color = 13561798 
 .TintAndShade = 0 
 End With 
MsgBox "Added an Above Average Conditional Format to Melanie's data. Press F9 to update values.", vbInformation 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Delete](aboveaverage-delete-method-excel.md)|
|[ModifyAppliesToRange](aboveaverage-modifyappliestorange-method-excel.md)|
|[SetFirstPriority](aboveaverage-setfirstpriority-method-excel.md)|
|[SetLastPriority](aboveaverage-setlastpriority-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[AboveBelow](aboveaverage-abovebelow-property-excel.md)|
|[Application](aboveaverage-application-property-excel.md)|
|[AppliesTo](aboveaverage-appliesto-property-excel.md)|
|[Borders](aboveaverage-borders-property-excel.md)|
|[CalcFor](aboveaverage-calcfor-property-excel.md)|
|[Creator](aboveaverage-creator-property-excel.md)|
|[Font](aboveaverage-font-property-excel.md)|
|[Interior](aboveaverage-interior-property-excel.md)|
|[NumberFormat](aboveaverage-numberformat-property-excel.md)|
|[NumStdDev](aboveaverage-numstddev-property-excel.md)|
|[Parent](aboveaverage-parent-property-excel.md)|
|[Priority](aboveaverage-priority-property-excel.md)|
|[PTCondition](aboveaverage-ptcondition-property-excel.md)|
|[ScopeType](aboveaverage-scopetype-property-excel.md)|
|[StopIfTrue](aboveaverage-stopiftrue-property-excel.md)|
|[Type](aboveaverage-type-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
