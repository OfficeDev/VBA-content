---
title: Application.BoxLinkLabelsShow Method (Project)
keywords: vbapj.chm47
f1_keywords:
- vbapj.chm47
ms.prod: project-server
api_name:
- Project.Application.BoxLinkLabelsShow
ms.assetid: 8dbb1406-10e8-d096-540a-4c7cfd61a413
ms.date: 06/08/2017
---


# Application.BoxLinkLabelsShow Method (Project)

Shows or hides link labels in the active Network Diagram.


## Syntax

 _expression_. **BoxLinkLabelsShow**( ** _Show_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Show_|Optional|**Boolean**|**True** if link labels display in the active view. The default value is **True** if link labels are hidden and **False** if they are visible.|

### Return Value

 **Boolean**


## Example

The following example first displays and then hides the labels.


```vb
Sub ShowBoxLink() 
 
 'Activate the Network Diagram view 
 ViewApply Name:="Network Diagram" 
 
 Result = BoxLinkLabelsShow(True) 
 Result = BoxLinkLabelsShow(False) 
End Sub
```


