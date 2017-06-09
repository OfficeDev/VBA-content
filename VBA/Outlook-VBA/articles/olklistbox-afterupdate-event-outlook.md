---
title: OlkListBox.AfterUpdate Event (Outlook)
keywords: vbaol11.chm1000291
f1_keywords:
- vbaol11.chm1000291
ms.prod: outlook
api_name:
- Outlook.OlkListBox.AfterUpdate
ms.assetid: 140c3cfd-ddad-a6cd-17bb-c8f5297c181e
ms.date: 06/08/2017
---


# OlkListBox.AfterUpdate Event (Outlook)

Occurs after the data in the control has been changed through the user interface.


## Syntax

 _expression_ . **AfterUpdate**

 _expression_ A variable that represents an **OlkListBox** object.


## Remarks

 **BeforeUpdate** and **AfterUpdate** can occur any time the data in the control is being saved to the item. The typical sequence of events involving **AfterUpdate** for this control is as follows:


1. User focuses on the control
    
2.  **BeforeUpdate**
    
3. Control data is updated
    
4.  ** AfterUpdate**
    
5.  **Exit** : User moves focus away from control
    



## See also


#### Concepts


[OlkListBox Object](olklistbox-object-outlook.md)

