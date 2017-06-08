---
title: Application.BoxProgressMarksShow Method (Project)
keywords: vbapj.chm46
f1_keywords:
- vbapj.chm46
ms.prod: project-server
api_name:
- Project.Application.BoxProgressMarksShow
ms.assetid: fd0ff0bd-7069-5e41-fa50-a47a4b09e9f6
ms.date: 06/08/2017
---


# Application.BoxProgressMarksShow Method (Project)

Shows or hides progress marks in the active Network Diagram.


## Syntax

 _expression_. **BoxProgressMarksShow**( ** _Show_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Show_|Optional|**Boolean**|**True** if progress marks display in the active view. The default value is **True** if progress marks are hidden and **False** if they are visible.|

### Return Value

 **Boolean**


## Example

The following example first displays and then hides the progress marks.


```vb
Sub BoxProgress_MarksShow() 
 
 Dim Result As Boolean 
 
 'Activate the Network Diagram view 
 ViewApply Name:="Network Diagram" 
 
 Result = BoxProgressMarksShow(True) 
 Result = BoxProgressMarksShow(False) 
 
End Sub
```


