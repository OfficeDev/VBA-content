---
title: OlkTextBox.AfterUpdate Event (Outlook)
keywords: vbaol11.chm1000082
f1_keywords:
- vbaol11.chm1000082
ms.prod: outlook
api_name:
- Outlook.OlkTextBox.AfterUpdate
ms.assetid: f61b5a19-4f3d-9287-d681-d5ac7b8979a4
ms.date: 06/08/2017
---


# OlkTextBox.AfterUpdate Event (Outlook)

Occurs after the data in the control has been changed through the user interface.


## Syntax

 _expression_ . **AfterUpdate**

 _expression_ A variable that represents an **OlkTextBox** object.


## Remarks

 **BeforeUpdate** and **AfterUpdate** can occur any time the data in the control is being saved to the item. The typical sequence of events involving **AfterUpdate** for this control is as follows:


1. User focuses on the control
    
2.  **BeforeUpdate**
    
3. Control data is updated
    
4.  ** AfterUpdate**
    
5.  **Exit** : User moves focus away from control
    



## See also


#### Concepts


[OlkTextBox Object](olktextbox-object-outlook.md)

