---
title: Table.AutoFitBehavior Method (Word)
keywords: vbawd10.chm156303380
f1_keywords:
- vbawd10.chm156303380
ms.prod: word
api_name:
- Word.Table.AutoFitBehavior
ms.assetid: 74e162a5-cde0-bdd3-2ea6-f78fb0ecca5a
ms.date: 06/08/2017
---


# Table.AutoFitBehavior Method (Word)

Determines how Microsoft Word resizes a table when the AutoFit feature is used.


## Syntax

 _expression_ . **AutoFitBehavior**( **_Behavior_** )

 _expression_ Required. A variable that represents a **[Table](table-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Behavior_|Required| **WdAutoFitBehavior**|How Word resizes the specified table with the AutoFit feature is used.|

## Remarks

Word can resize the table based on the content of the table cells or the width of the document window. You can also use this method to turn off AutoFit so that the table size is fixed, regardless of cell contents or window width.

Setting the  **AutoFitBehavior** property to **wdAutoFitContent** or **wdAutoFitWindow** sets the **AllowAutoFit** property to **True** if it is currently **False** . Likewise, setting the **AutoFitBehavior** property to **wdAutoFitFixed** sets the **AllowAutoFit** property to **False** if it is currently **True** .


## Example

This example sets the AutoFit behavior for the first table in the active document to automatically resize based on the width of the document window.


```vb
ActiveDocument.Tables(1).AutoFitBehavior _ 
 wdAutoFitWindow
```


## See also


#### Concepts


[Table Object](table-object-word.md)

