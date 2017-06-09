---
title: AddIns2.Item Property (Excel)
keywords: vbaxl10.chm867075
f1_keywords:
- vbaxl10.chm867075
ms.prod: excel
api_name:
- Excel.AddIns2.Item
ms.assetid: 6a0ee1ca-75ba-14ed-1bbb-606480aeafbe
ms.date: 06/08/2017
---


# AddIns2.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that returns a **AddIns2** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Example

This example displays the status of the Analysis ToolPak add-in. Note that the string used as the index to the  **AddIns2** method is the **Title** property of the **AddIn** object.


```vb
If ThisWorkbook.Application.AddIns2.Item("Analysis ToolPak").Installed = True Then 
 MsgBox "Analysis ToolPak add-in is installed" 
Else 
 MsgBox "Analysis ToolPak add-in is not installed" 
End If
```


## See also


#### Concepts


[AddIns2 Object](addins2-object-excel.md)

