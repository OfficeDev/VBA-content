---
title: OlkDateControl.AfterUpdate Event (Outlook)
keywords: vbaol11.chm1000374
f1_keywords:
- vbaol11.chm1000374
ms.prod: outlook
api_name:
- Outlook.OlkDateControl.AfterUpdate
ms.assetid: 7086c185-99a2-94e1-6041-64c58869067f
ms.date: 06/08/2017
---


# OlkDateControl.AfterUpdate Event (Outlook)

Occurs after the data in the control has been changed through the user interface.


## Syntax

 _expression_ . **AfterUpdate**

 _expression_ A variable that represents an **OlkDateControl** object.


## Remarks

 **BeforeUpdate** and **AfterUpdate** can occur any time the data in the control is being saved to the item. The typical sequence of events involving **AfterUpdate** for this control is as follows:


1. User focuses on the control
    
2.  **BeforeUpdate**
    
3. Control data is updated
    
4.  ** AfterUpdate**
    
5.  **Exit** : User moves focus away from control
    



## See also


#### Concepts


[OlkDateControl Object](olkdatecontrol-object-outlook.md)

