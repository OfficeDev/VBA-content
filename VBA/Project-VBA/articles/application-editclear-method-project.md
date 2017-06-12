---
title: Application.EditClear Method (Project)
keywords: vbapj.chm205
f1_keywords:
- vbapj.chm205
ms.prod: project-server
api_name:
- Project.Application.EditClear
ms.assetid: 0f87ca1c-c87c-774a-e8dd-2f4d29a40e28
ms.date: 06/08/2017
---


# Application.EditClear Method (Project)

Clears the selected cells.


## Syntax

 _expression_. **EditClear**( ** _Contents_**, ** _Formats_**, ** _Notes_**, ** _Hyperlinks_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Contents_|Optional|**Boolean**|**True** if the contents of the selected cells are cleared. The default value is **True**.|
| _Formats_|Optional|**Boolean**|**True** if the formats of the selected cells are cleared. The default value is **False**.|
| _Notes_|Optional|**Boolean**|**True** if the notes of the assignment, resource, or task in the selected cells are cleared. The default value is **False**.|
| _Hyperlinks_|Optional|**Boolean**|**True** if the hyperlinks associated with the selected cells are removed. The default value is **False**.|

### Return Value

 **Boolean**


## Example

The following example clears the contents, formats, and notes of the selected cells.


```vb
Sub ClearAll() 
 EditClear True, True, True 
End Sub
```


