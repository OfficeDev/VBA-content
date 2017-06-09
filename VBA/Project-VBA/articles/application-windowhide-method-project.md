---
title: Application.WindowHide Method (Project)
keywords: vbapj.chm703
f1_keywords:
- vbapj.chm703
ms.prod: project-server
api_name:
- Project.Application.WindowHide
ms.assetid: 37219d9d-1e50-3341-7618-9827d077d4d8
ms.date: 06/08/2017
---


# Application.WindowHide Method (Project)

Hides a window.


## Syntax

 _expression_. **WindowHide**( ** _Name_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the window to hide. The name of a window is the exact text that appears in the title bar of the window. The default is the active window.|

### Return Value

 **Boolean**


## Example

The following example hides all windows except the active window.


```vb
Sub HideAllWindowsExceptActive() 
 
 Dim I As Long ' Index for For...Next loop 
 
 For I = 1 To Windows.Count 
 If Windows(I) <> ActiveWindow And Windows(I).Visible Then 
 
 WindowHide Windows(I).Caption 
 End If 
 Next I 
 
End Sub
```


