---
title: Application.BoxShowHideFields Method (Project)
keywords: vbapj.chm905
f1_keywords:
- vbapj.chm905
ms.prod: project-server
api_name:
- Project.Application.BoxShowHideFields
ms.assetid: b100c012-8ab9-2e39-c8c8-569b1498c5da
ms.date: 06/08/2017
---


# Application.BoxShowHideFields Method (Project)

Shows or hides the task data fields of the active Network Diagram.


## Syntax

 _expression_. **BoxShowHideFields**( ** _Show_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Show_|Optional|**Boolean**|**True** if the fields of Network Diagram boxes are displayed in the active Network Diagram. **False** if only task ID numbers are displayed. The default value is **True** if the active Network Diagram isn't showing fields and **False** if it is.|

### Return Value

 **Boolean**


## Example

The following example first hides the fields of Network Diagram boxes and then displays them.


```vb
Sub BoxShow_HideFields() 
 
 Dim Result As Boolean 
 
 'Activate the Network Diagram view 
 ViewApply Name:="Network Diagram" 
 
 Result = BoxShowHideFields(False) 
 Result = BoxShowHideFields(True) 
 
End Sub
```


