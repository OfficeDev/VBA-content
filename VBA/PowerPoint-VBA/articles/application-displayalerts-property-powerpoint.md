---
title: Application.DisplayAlerts Property (PowerPoint)
keywords: vbapp10.chm502050
f1_keywords:
- vbapp10.chm502050
ms.prod: powerpoint
api_name:
- PowerPoint.Application.DisplayAlerts
ms.assetid: e18cf1f5-c456-8cd5-40e7-eec69c40811d
ms.date: 06/08/2017
---


# Application.DisplayAlerts Property (PowerPoint)

Sets or returns whether Microsoft PowerPoint displays alerts while running a macro. Read/write.


## Syntax

 _expression_. **DisplayAlerts**

 _expression_ A variable that represents an **Application** object.


### Return Value

PpAlertLevel


## Remarks

The value of the  **DisplayAlerts** property is not reset once a macro stops running; it is maintained throughout a session. It is not stored across sessions, so when PowerPoint begins, it is reset to **ppAlertsNone**.

The value of the  **DisplayAlerts** property can be one of these **PpAlertLevel** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**ppAlertsAll**| All message boxes and alerts are displayed; errors are returned to the macro.|
|**ppAlertsNone**|The default. No alerts or message boxes are displayed. If a macro encounters a message box, the default value is chosen and the macro continues.|

## Example

The following line of code instructs PowerPoint to display all message boxes and alerts, returning errors to the macro.


```vb
Sub SetAlert

    Application.DisplayAlerts = ppAlertsAll

End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

