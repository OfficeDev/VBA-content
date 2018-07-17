---
title: OlkCheckBox.Exit Event (Outlook)
keywords: vbaol11.chm1000155
f1_keywords:
- vbaol11.chm1000155
ms.prod: outlook
api_name:
- Outlook.OlkCheckBox.Exit
ms.assetid: a89b3d32-c540-ea72-b018-fabc9b9760f3
ms.date: 06/08/2017
---


# OlkCheckBox.Exit Event (Outlook)

Occurs just after the focus passes from this control to another control on the same form.


## Syntax

 _expression_ . **Exit**( **_Cancel_** )

 _expression_ A variable that represents an **OlkCheckBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the exit operation is not completed and the focus remains in this control.|

## See also


#### Concepts


[OlkCheckBox Object](olkcheckbox-object-outlook.md)

