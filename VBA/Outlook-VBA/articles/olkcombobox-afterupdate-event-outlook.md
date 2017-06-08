---
title: OlkComboBox.AfterUpdate Event (Outlook)
keywords: vbaol11.chm1000247
f1_keywords:
- vbaol11.chm1000247
ms.prod: outlook
api_name:
- Outlook.OlkComboBox.AfterUpdate
ms.assetid: d130f15a-832e-f2d1-a6f4-13edcfb5bd9d
ms.date: 06/08/2017
---


# OlkComboBox.AfterUpdate Event (Outlook)

Occurs after the data in the control has been changed through the user interface.


## Syntax

 _expression_ . **AfterUpdate**

 _expression_ A variable that represents an **OlkComboBox** object.


## Remarks

 **BeforeUpdate** and **AfterUpdate** can occur any time the data in the control is being saved to the item. The typical sequence of events involving **AfterUpdate** for this control is as follows:


1. User focuses on the control
    
2.  **BeforeUpdate**
    
3. Control data is updated
    
4.  ** AfterUpdate**
    
5.  **Exit** : User moves focus away from control
    



## See also


#### Concepts


[OlkComboBox Object](olkcombobox-object-outlook.md)

