---
title: Application.BoxAlign Method (Project)
keywords: vbapj.chm29
f1_keywords:
- vbapj.chm29
ms.prod: project-server
api_name:
- Project.Application.BoxAlign
ms.assetid: 2b27c9a0-36fa-1bbd-96e3-267b95ad5407
ms.date: 06/08/2017
---


# Application.BoxAlign Method (Project)

Aligns the specified part of the selected boxes in the active Network Diagram view with the same part of the box that has the focus.


## Syntax

 _expression_. **BoxAlign**( ** _Alignment_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Alignment_|Required|**Long**|Specifies which side or portion of a box to use for the alignment. Can be one of the  **[PjAlign](pjalign-enumeration-project.md)** constants.|

### Return Value

 **Boolean**


## Remarks

If only one box is selected, the  **BoxAlign** method has no effect.


