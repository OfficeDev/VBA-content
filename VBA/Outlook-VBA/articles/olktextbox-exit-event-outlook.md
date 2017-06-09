---
title: OlkTextBox.Exit Event (Outlook)
keywords: vbaol11.chm1000077
f1_keywords:
- vbaol11.chm1000077
ms.prod: outlook
api_name:
- Outlook.OlkTextBox.Exit
ms.assetid: ea36905e-bd5a-2d6c-6ea6-0ad33d965741
ms.date: 06/08/2017
---


# OlkTextBox.Exit Event (Outlook)

Occurs just after the focus passes from this control to another control on the same form.


## Syntax

 _expression_ . **Exit**( **_Cancel_** )

 _expression_ A variable that represents an **OlkTextBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the exit operation is not completed and the focus remains in this control.|

## See also


#### Concepts


[OlkTextBox Object](olktextbox-object-outlook.md)

