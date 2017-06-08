---
title: Inspector.BeforeMove Event (Outlook)
keywords: vbaol11.chm470
f1_keywords:
- vbaol11.chm470
ms.prod: outlook
api_name:
- Outlook.Inspector.BeforeMove
ms.assetid: 52a4445e-4d76-7b55-ce28-d972fba87a9b
ms.date: 06/08/2017
---


# Inspector.BeforeMove Event (Outlook)

Occurs when the  **[Inspector](inspector-object-outlook.md)** is moved by the user.


## Syntax

 _expression_ . **BeforeMove**( **_Cancel_** )

 _expression_ A variable that represents an **Inspector** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the operation is not completed and the inspector is not moved.|

## Remarks

This event can be cancelled after it has started.


## See also


#### Concepts


[Inspector Object](inspector-object-outlook.md)

