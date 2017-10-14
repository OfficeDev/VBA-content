---
title: Results.ResetColumns Method (Outlook)
keywords: vbaol11.chm509
f1_keywords:
- vbaol11.chm509
ms.prod: outlook
api_name:
- Outlook.Results.ResetColumns
ms.assetid: 1839dd92-cbab-5fac-a59b-b1ceb6ef874a
ms.date: 06/08/2017
---


# Results.ResetColumns Method (Outlook)

Clears the properties that have been cached with the  **[SetColumns](results-setcolumns-method-outlook.md)** method.


## Syntax

 _expression_ . **ResetColumns**

 _expression_ A variable that represents a **Results** object.


## Remarks

All properties are accessible after calling the  **ResetColumns** method. **SetColumns** should be reused to store new properties again. **ResetColumns** does nothing if **SetColumns** has not been called first.


## See also


#### Concepts


[Results Object](results-object-outlook.md)

