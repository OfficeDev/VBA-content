---
title: AddIns.Item Property (Excel)
keywords: vbaxl10.chm187075
f1_keywords:
- vbaxl10.chm187075
ms.prod: excel
api_name:
- Excel.AddIns.Item
ms.assetid: 417987d5-322c-2784-c51e-18a1fa7578d1
ms.date: 06/08/2017
---


# AddIns.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents an **AddIns** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Example

This example displays the status of the Analysis ToolPak add-in. Note that the string used as the index to the  **AddIns** method is the **Title** property of the **AddIn** object.


```vb
If ThisWorkbook.Application.AddIns.Item("Analysis ToolPak").Installed = True Then 
 MsgBox "Analysis ToolPak add-in is installed" 
Else 
 MsgBox "Analysis ToolPak add-in is not installed" 
End If
```


## See also


#### Concepts


[AddIns Collection](addins-object-excel.md)

