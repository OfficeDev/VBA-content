---
title: OlkOptionButton.AfterUpdate Event (Outlook)
keywords: vbaol11.chm1000190
f1_keywords:
- vbaol11.chm1000190
ms.prod: outlook
api_name:
- Outlook.OlkOptionButton.AfterUpdate
ms.assetid: aa573288-f4fb-656c-304b-f564335c8c2d
ms.date: 06/08/2017
---


# OlkOptionButton.AfterUpdate Event (Outlook)

Occurs after the data in the control has been changed through the user interface.


## Syntax

 _expression_ . **AfterUpdate**

 _expression_ A variable that represents an **OlkOptionButton** object.


## Remarks

 **BeforeUpdate** and **AfterUpdate** can occur any time the data in the control is being saved to the item. The typical sequence of events involving **AfterUpdate** for this control is as follows:


1. User focuses on the control
    
2.  **BeforeUpdate**
    
3. Control data is updated
    
4.  ** AfterUpdate**
    
5.  **Exit** : User moves focus away from control
    



## See also


#### Concepts


[OlkOptionButton Object](olkoptionbutton-object-outlook.md)

