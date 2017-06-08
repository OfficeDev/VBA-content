---
title: Application.SheetFollowHyperlink Event (Excel)
keywords: vbaxl10.chm504093
f1_keywords:
- vbaxl10.chm504093
ms.prod: excel
api_name:
- Excel.Application.SheetFollowHyperlink
ms.assetid: 656e0ec6-64ea-1685-f088-a7e30bfaef38
ms.date: 06/08/2017
---


# Application.SheetFollowHyperlink Event (Excel)

Occurs when you click any hyperlink in Microsoft Excel. For worksheet-level events, see the Help topic for the  **[FollowHyperlink](worksheet-followhyperlink-event-excel.md)** event.


## Syntax

 _expression_ . **SheetFollowHyperlink**( **_Sh_** , **_Target_** )

 _expression_ An expression that returns a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The  **[Worksheet](worksheet-object-excel.md)** object that contains the hyperlink.|
| _Target_|Required| **Hyperlink**|The  **Hyperlink** object that represents the destination of the hyperlink.|

## See also


#### Concepts


[Application Object](application-object-excel.md)

