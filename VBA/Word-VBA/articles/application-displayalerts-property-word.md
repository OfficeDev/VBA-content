---
title: Application.DisplayAlerts Property (Word)
keywords: vbawd10.chm158335070
f1_keywords:
- vbawd10.chm158335070
ms.prod: word
api_name:
- Word.Application.DisplayAlerts
ms.assetid: 23d35e76-d5be-c1ed-4312-b6c220413882
ms.date: 06/08/2017
---


# Application.DisplayAlerts Property (Word)

Returns or sets the way certain alerts and messages are handled while a macro is running. Read/write  **WdAlertLevel** .


## Syntax

 _expression_ . **DisplayAlerts**

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example sets Word to display all alerts and message boxes when it is running macros.


```vb
Application.DisplayAlerts = wdAlertsAll
```

This example returns the current setting of the  **DisplayAlerts** property.




```vb
Dim lngTemp As Long 
 
lngTemp = Application.DisplayAlerts
```


## See also


#### Concepts


[Application Object](application-object-word.md)

