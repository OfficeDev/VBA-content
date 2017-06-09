---
title: Application.Form Method (Project)
keywords: vbapj.chm1004
f1_keywords:
- vbapj.chm1004
ms.prod: project-server
api_name:
- Project.Application.Form
ms.assetid: 23e7c800-bda9-c931-bc27-084dec872953
ms.date: 06/08/2017
---


# Application.Form Method (Project)

Displays a custom form. The  **Form** method produces an error if a resource form is specified when the active view is a task view, and vice versa.


## Syntax

 _expression_. **Form**( ** _Name_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of a custom form. The default is a task form when the active view is a task view, and a resource form when the active view is a resource view.|

### Return Value

 **Boolean**


## Example

The following example displays the Cost Tracking form.


```vb
Sub DisplayCostTrackingForm 
 Form("Cost Tracking") 
End Sub
```


