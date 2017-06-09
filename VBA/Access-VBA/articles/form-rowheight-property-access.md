---
title: Form.RowHeight Property (Access)
keywords: vbaac10.chm13395
f1_keywords:
- vbaac10.chm13395
ms.prod: access
api_name:
- Access.Form.RowHeight
ms.assetid: 1575cb30-54ab-d45b-bb64-844f12336eca
ms.date: 06/08/2017
---


# Form.RowHeight Property (Access)

You can use the  **RowHeight** property to specify the height of all rows in Datasheet view. Read/write **Integer**.


## Syntax

 _expression_. **RowHeight**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **RowHeight** property applies to all fields in Datasheet view and to form controls when the form is in Datasheet view.

The  **RowHeight** property setting represents the datasheet row height in twips. To specify the default height for the current font, you can set the **RowHeight** property to **True**.


## Example

This example takes effect in Datasheet view of the open Customers form. It sets the row height to 450 twips and sizes the column to fit the size of the visible text.


```vb
Forms![Customers].RowHeight = 450 
Forms![Customers]![Address].ColumnWidth = -2
```


## See also


#### Concepts


[Form Object](form-object-access.md)

