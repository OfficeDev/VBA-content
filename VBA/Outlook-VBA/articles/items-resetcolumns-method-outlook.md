---
title: Items.ResetColumns Method (Outlook)
keywords: vbaol11.chm69
f1_keywords:
- vbaol11.chm69
ms.prod: outlook
api_name:
- Outlook.Items.ResetColumns
ms.assetid: 0543dd17-1e65-5484-ab21-d4791b3b1194
ms.date: 06/08/2017
---


# Items.ResetColumns Method (Outlook)

Clears the properties that have been cached with the  **[SetColumns](items-setcolumns-method-outlook.md)** method.


## Syntax

 _expression_ . **ResetColumns**

 _expression_ A variable that represents an **Items** object.


## Remarks

All properties are accessible after calling the  **ResetColumns** method. **SetColumns** should be reused to store new properties again. **ResetColumns** does nothing if **SetColumns** has not been called first.


## See also


#### Concepts


[Items Object](items-object-outlook.md)

