---
title: OlkCommandButton.AfterUpdate Event (Outlook)
keywords: vbaol11.chm1000130
f1_keywords:
- vbaol11.chm1000130
ms.prod: outlook
api_name:
- Outlook.OlkCommandButton.AfterUpdate
ms.assetid: 2f968ed1-7043-a3de-8219-927c27e12832
ms.date: 06/08/2017
---


# OlkCommandButton.AfterUpdate Event (Outlook)

Occurs after the data in the control has been changed through the user interface.


## Syntax

 _expression_ . **AfterUpdate**

 _expression_ A variable that represents an **OlkCommandButton** object.


## Remarks

 **BeforeUpdate** and **AfterUpdate** can occur any time the data in the control is being saved to the item. The typical sequence of events involving **AfterUpdate** for this control is as follows:


1. User focuses on the control
    
2.  **BeforeUpdate**
    
3. Control data is updated
    
4.  ** AfterUpdate**
    
5.  **Exit** : User moves focus away from control
    



## See also


#### Concepts


[OlkCommandButton Object](olkcommandbutton-object-outlook.md)

