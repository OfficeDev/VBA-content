---
title: IRtdServer.Heartbeat Method (Excel)
keywords: vbaxl10.chm500009
f1_keywords:
- vbaxl10.chm500009
ms.prod: excel
api_name:
- Excel.IRtdServer.Heartbeat
ms.assetid: 9dc61d35-30cb-fcbe-6aaf-acb2df61d535
ms.date: 06/08/2017
---


# IRtdServer.Heartbeat Method (Excel)

Determines if the real-time data server is still active. Returns a  **Long** value. Zero or a negative number indicates failure; a positive number indicates that the server is active.


## Syntax

 _expression_ . **Heartbeat**

 _expression_ A variable that represents an **IRtdServer** object.


### Return Value

Long


## Remarks

The  **Heartbeat** method is called by Microsoft Excel if the **[HeartbeatInterval](irtdupdateevent-heartbeatinterval-property-excel.md)** property has elapsed since the last time Excel was called with the **[UpdateNotify](irtdupdateevent-updatenotify-method-excel.md)** method.


## See also


#### Concepts


[IRtdServer Object](irtdserver-object-excel.md)

