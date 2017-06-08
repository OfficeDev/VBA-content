---
title: DisplayAlerts Property
keywords: vbagr10.chm65879
f1_keywords:
- vbagr10.chm65879
ms.prod: excel
api_name:
- Excel.DisplayAlerts
ms.assetid: 630e60be-23e3-795b-1ed9-26b791fb7efc
ms.date: 06/08/2017
---


# DisplayAlerts Property

 **True** if Microsoft Graph displays certain alerts and messages while a macro is running. Read/write **Boolean**.


## Remarks

The default value is  **True**. Set this property to  **False** if you don't want to be disturbed by prompts and alert messages while a macro is running; any time a message requires a response, Microsoft Graph chooses the default response.

If you set this property to  **False**, Microsoft Graph doesn't automatically set it back to  **True** when your macro stops running. Write your macro such that it always sets this property back to **True** when it stops running.


