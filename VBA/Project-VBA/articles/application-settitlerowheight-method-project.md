---
title: Application.SetTitleRowHeight Method (Project)
keywords: vbapj.chm2120
f1_keywords:
- vbapj.chm2120
ms.prod: project-server
api_name:
- Project.Application.SetTitleRowHeight
ms.assetid: 7ee0d6db-9fd5-bcd4-e495-14d0a270ed99
ms.date: 06/08/2017
---


# Application.SetTitleRowHeight Method (Project)

Sets the title row height of the active view.


## Syntax

 _expression_. **SetTitleRowHeight**( ** _TitleHeight_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TitleHeight_|Optional|**Integer**|The height of the title row of the active view.|

### Return Value

 **Boolean**


## Remarks

Using the  **SetTitleRowHeight** method without specifying any arguments sets the title row height to the default height of the active view.

The  **SetTitleRowHeight** method applies only to sheet views. Project returns a trappable error (error code 1100) in a non-sheet view such as the **Network Diagram** or **Calendar** view.


