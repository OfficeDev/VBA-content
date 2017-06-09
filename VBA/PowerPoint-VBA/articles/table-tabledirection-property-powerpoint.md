---
title: Table.TableDirection Property (PowerPoint)
keywords: vbapp10.chm622006
f1_keywords:
- vbapp10.chm622006
ms.prod: powerpoint
api_name:
- PowerPoint.Table.TableDirection
ms.assetid: 3fbb1c4b-6cdb-f97e-7b85-c41897bc5ced
ms.date: 06/08/2017
---


# Table.TableDirection Property (PowerPoint)

Returns or sets the direction in which the table cells are ordered. Read/write.


## Syntax

 _expression_. **TableDirection**

 _expression_ A variable that represents a **Table** object.


### Return Value

PpDirection


## Remarks

The default value of the  **TableDirection** property is **ppDirectionLefttToRight**, unless the **[LanguageSettings](application-languagesettings-property-powerpoint.md)** property or the **[DefaultLanguageID](presentation-defaultlanguageid-property-powerpoint.md)** property is set to a right-to-left language, in which case the default value is **ppDirectionRightToLeft**.

The value of the  **TableDirection** property can be one of these **PpDirection** constants.


||
|:-----|
|**ppDirectionLeftToRight**|
|**ppDirectionMixed**|
|**ppDirectionRightToLeft**|
When you are using the  **TextDirection** property, The **ppDirectionMixed** constant may be returned.


## Example

This example sets the direction in which cells in the selected table are ordered to left to right (first column is the leftmost column).


```vb
ActiveWindow.Selection.ShapeRange.Table.TableDirection = _
    ppDirectionLeftToRight
```


## See also


#### Concepts


[Table Object](table-object-powerpoint.md)

