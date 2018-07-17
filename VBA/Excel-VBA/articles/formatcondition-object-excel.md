---
title: FormatCondition Object (Excel)
keywords: vbaxl10.chm511072
f1_keywords:
- vbaxl10.chm511072
ms.prod: excel
api_name:
- Excel.FormatCondition
ms.assetid: 38a2bca9-9b28-3ef2-8c7a-4d35a27229ec
ms.date: 06/08/2017
---


# FormatCondition Object (Excel)

Represents a conditional format.


## Remarks

 The **FormatCondition** object is a member of the **[FormatConditions](formatconditions-object-excel.md)** collection. The **FormatConditions** collection can now contain more than three conditional formats for a given range.

Use the  **[Add](formatconditions-add-method-excel.md)** method to create a new conditional format. If a range has mulitple formats, you can use the **[Modify](formatcondition-modify-method-excel.md)** method to change one of the formats, or you can use the **[Delete](formatcondition-delete-method-excel.md)** method to delete a format and then use the **Add** method to create a new format.

Use the  **[Font](formatcondition-font-property-excel.md)**, **[Borders](formatcondition-borders-property-excel.md)**, and **[Interior](formatcondition-interior-property-excel.md)** properties of the **FormatCondition** object to control the appearance of formatted cells. Some properties of these objects aren?t supported by the conditional format object model. Some of the properties that can be used with conditional formatting are listed in the following table.



|**Object**|**Properties**|
|:-----|:-----|
|**[Font](font-object-excel.md)**|**Bold** **Color** **ColorIndex** **FontStyle** **Italic** **Strikethrough** **Underline** The accounting underline styles cannot be used.|
|**[Border](border-object-excel.md)**|**Bottom** **Color** **Left** **Right** **Style** The following border styles can be used (all others aren?t supported): **xlNone**, **xlSolid**, **xlDash**, **xlDot**, **xlDashDot**, **xlDashDotDot**, **xlGray50**, **xlGray75**, and **xlGray25**. **Top** **Weight** The following border weights can be used (all others aren?t supported): **xlWeightHairline** and **xlWeightThin**.|
|**[Interior](interior-object-excel.md)**|**Color** **ColorIndex** **Pattern** **PatternColorIndex**|

## Example

Use  **[FormatConditions](range-formatconditions-property-excel.md)** ( _index_ ), where _index_ is the index number of the conditional format, to return a **FormatCondition** object. The following example sets format properties for an existing conditional format for cells E1:E10.


```
With Worksheets(1).Range("e1:e10").FormatConditions(1) 
 With .Borders 
 .LineStyle = xlContinuous 
 .Weight = xlThin 
 .ColorIndex = 6 
 End With 
 With .Font 
 .Bold = True 
 .ColorIndex = 3 
 End With 
End With
```


## Methods



|**Name**|
|:-----|
|[Delete](formatcondition-delete-method-excel.md)|
|[Modify](formatcondition-modify-method-excel.md)|
|[ModifyAppliesToRange](formatcondition-modifyappliestorange-method-excel.md)|
|[SetFirstPriority](formatcondition-setfirstpriority-method-excel.md)|
|[SetLastPriority](formatcondition-setlastpriority-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](formatcondition-application-property-excel.md)|
|[AppliesTo](formatcondition-appliesto-property-excel.md)|
|[Borders](formatcondition-borders-property-excel.md)|
|[Creator](formatcondition-creator-property-excel.md)|
|[DateOperator](formatcondition-dateoperator-property-excel.md)|
|[Font](formatcondition-font-property-excel.md)|
|[Formula1](formatcondition-formula1-property-excel.md)|
|[Formula2](formatcondition-formula2-property-excel.md)|
|[Interior](formatcondition-interior-property-excel.md)|
|[NumberFormat](formatcondition-numberformat-property-excel.md)|
|[Operator](formatcondition-operator-property-excel.md)|
|[Parent](formatcondition-parent-property-excel.md)|
|[Priority](formatcondition-priority-property-excel.md)|
|[PTCondition](formatcondition-ptcondition-property-excel.md)|
|[ScopeType](formatcondition-scopetype-property-excel.md)|
|[StopIfTrue](formatcondition-stopiftrue-property-excel.md)|
|[Text](formatcondition-text-property-excel.md)|
|[TextOperator](formatcondition-textoperator-property-excel.md)|
|[Type](formatcondition-type-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
