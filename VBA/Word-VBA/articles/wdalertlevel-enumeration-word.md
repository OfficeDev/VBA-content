---
title: WdAlertLevel Enumeration (Word)
keywords: vbawd10.chm0
f1_keywords:
- vbawd10.chm0
ms.prod: word
api_name:
- Word.WdAlertLevel
ms.assetid: ecfcfb35-0fdc-0881-53aa-3c2e3f5350bf
ms.date: 06/08/2017
---


# WdAlertLevel Enumeration (Word)

Specifies the way certain alerts and messages are handled while a macro is running.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **wdAlertsAll**|-1|All message boxes and alerts are displayed; errors are returned to the macro.|
| **wdAlertsMessageBox**|-2|Only message boxes are displayed; errors are trapped and returned to the macro.|
| **wdAlertsNone**|0|No alerts or message boxes are displayed. If a macro encounters a message box, the default value is chosen and the macro continues.|

