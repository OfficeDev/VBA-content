---
title: ControlFormat.MultiSelect Property (Excel)
keywords: vbaxl10.chm630087
f1_keywords:
- vbaxl10.chm630087
ms.prod: excel
api_name:
- Excel.ControlFormat.MultiSelect
ms.assetid: 5ec1e5b6-37ab-465b-bf81-4955f6fd0f31
ms.date: 06/08/2017
---


# ControlFormat.MultiSelect Property (Excel)

Returns or sets the selection mode of the specified list box. Can be one of the following constants:  **xlNone** , **xlSimple** , or **xlExtended** . Read/write **Long** .


## Syntax

 _expression_ . **MultiSelect**

 _expression_ A variable that represents a **ControlFormat** object.


## Remarks

Single select ( **xlNone** ) allows only one item at a time to be selected. Clicking the mouse or pressing the SPACEBAR cancels the selection and selects the clicked item.

Simple multiselect ( **xlSimple** ) toggles the selection on an item in the list when click it with the mouse or press the SPACEBAR when the focus is on the item. This mode is appropriate for pick lists, in which there are often multiple items selected.

Extended multiselect ( **xlExtended** ) usually acts like a single-selection list box, so when you click an item, you cancel all other selections. When you hold down SHIFT while clicking the mouse or pressing an arrow key, you select items sequentially from the current item. When you hold down CTRL while clicking the mouse, you add single items to the list. This mode is appropriate when multiple items are allowed but not often used.

You can use the  **Value** or **ListIndex** property to return and set the selected item in a single-select list box.

You cannot link multiselect list boxes by using the  **LinkedCell** property.


## Example

This example creates a simple multiselect list box.


```vb
Set lb = Worksheets(1).Shapes.AddFormControl(xlListBox, _ 
 Left:=10, Top:=10, Height:=100, Width:100) 
lb.ControlFormat.MultiSelect = xlSimple
```


## See also


#### Concepts


[ControlFormat Object](controlformat-object-excel.md)

