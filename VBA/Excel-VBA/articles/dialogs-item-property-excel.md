---
title: Dialogs.Item Property (Excel)
keywords: vbaxl10.chm254074
f1_keywords:
- vbaxl10.chm254074
ms.prod: excel
api_name:
- Excel.Dialogs.Item
ms.assetid: f9200ca3-711b-92ee-81b2-7c9cf1d104af
ms.date: 06/08/2017
---


# Dialogs.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **Dialogs** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **XlBuiltInDialog**| **Variant** . The name or index number of the object.|

## Example

This example displays the  **Open** dialog box and selects the **Read-Only** option.


```vb
Application.Dialogs.Item(xlDialogOpen).Show arg3:=True
```


## See also


#### Concepts


[Dialogs Object](dialogs-object-excel.md)

