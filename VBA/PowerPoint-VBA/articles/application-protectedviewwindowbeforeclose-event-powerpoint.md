---
title: Application.ProtectedViewWindowBeforeClose Event (PowerPoint)
keywords: vbapp10.chm621028
f1_keywords:
- vbapp10.chm621028
ms.prod: powerpoint
api_name:
- PowerPoint.Application.ProtectedViewWindowBeforeClose
ms.assetid: e10ffe16-aad8-1e2d-fd75-82243a56ef05
ms.date: 06/08/2017
---


# Application.ProtectedViewWindowBeforeClose Event (PowerPoint)

Occurs immediately before a protected view window or a document in a protected view window closes.


## Syntax

 _expression_. **ProtectedViewWindowBeforeClose**( **_ProtViewWindow_**, **_ProtectedViewCloseReason_**, **_Cancel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ProtViewWindow_|Required|**ProtectedViewWindow**|The protected view window that is closed.|
| _ProtectedViewCloseReason_|Required|**PpProtectedViewCloseReason**|A constant that specifies the reason the protected view window is closed.|
| _Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the window does not close when the procedure is finished.|

### Return Value

nothing


## Remarks

If the  **ProtectedViewWindowsBeforeClose** event is called as part of the[ProtectedViewWindow.Edit](protectedviewwindow-edit-method-powerpoint.md) method, setting _Cancel_ to **True** produces no action.


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

