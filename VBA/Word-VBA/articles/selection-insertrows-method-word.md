---
title: Selection.InsertRows Method (Word)
keywords: vbawd10.chm158663184
f1_keywords:
- vbawd10.chm158663184
ms.prod: word
api_name:
- Word.Selection.InsertRows
ms.assetid: 326ad049-4d39-1ca6-a203-ddba0e77cba4
ms.date: 06/08/2017
---


# Selection.InsertRows Method (Word)

Inserts the specified number of new rows above the row that contains the selection. If the selection isn't in a table, an error occurs.


## Syntax

 _expression_ . **InsertRows**( **_NumRows_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NumRows_|Optional| **Variant**|The number of rows to be added.|

## Remarks

You can also insert rows by using the  **[Add](rows-add-method-word.md)** method of the **Rows** object.


## Example

This example inserts two new rows above the row that contains the selection, and then it removes the borders from the new rows.


```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.InsertRows NumRows:=2 
 Selection.Borders.Enable =False 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

