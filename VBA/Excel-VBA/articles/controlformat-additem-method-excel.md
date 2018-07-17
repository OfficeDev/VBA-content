---
title: ControlFormat.AddItem Method (Excel)
keywords: vbaxl10.chm630073
f1_keywords:
- vbaxl10.chm630073
ms.prod: excel
api_name:
- Excel.ControlFormat.AddItem
ms.assetid: fffc243b-3f94-14ab-f7b4-83c56325aa5e
ms.date: 06/08/2017
---


# ControlFormat.AddItem Method (Excel)

Adds an item to a list box or a combo box.


## Syntax

 _expression_ . **AddItem**( **_Text_** , **_Index_** )

 _expression_ A variable that represents a **ControlFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Text_|Required| **String**|The text to be added.|
| _Index_|Optional| **Variant**|The position of the new entry. If the list has fewer entries than the specified index, blank items from the end of the list are added to the specified position. If this argument is omitted, the item is appended to the existing list.|

## Remarks

Using this method clears any range specified by the  **[ListFillRange](controlformat-listfillrange-property-excel.md)** property.


## Example

This example creates a list box and fills it with integers from 1 to 10.


```vb
With Worksheets(1) 
 Set lb = .Shapes.AddFormControl(xlListBox, 100, 10, 100, 100) 
 For x = 1 To 10 
 lb.ControlFormat.AddItem x 
 Next 
End With
```


## See also


#### Concepts


[ControlFormat Object](controlformat-object-excel.md)

