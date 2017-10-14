---
title: OlkCheckBox.AfterUpdate Event (Outlook)
keywords: vbaol11.chm1000160
f1_keywords:
- vbaol11.chm1000160
ms.prod: outlook
api_name:
- Outlook.OlkCheckBox.AfterUpdate
ms.assetid: a207e36b-9afe-b7e3-9dd4-9af2ae16cf7d
ms.date: 06/08/2017
---


# OlkCheckBox.AfterUpdate Event (Outlook)

Occurs after the data in the control has been changed through the user interface.


## Syntax

 _expression_ . **AfterUpdate**

 _expression_ A variable that represents an **OlkCheckBox** object.


## Remarks

 **BeforeUpdate** and **AfterUpdate** can occur any time the data in the control is being saved to the item. The typical sequence of events involving **AfterUpdate** for this control is as follows:


1. User focuses on the control
    
2.  **BeforeUpdate**
    
3. Control data is updated
    
4.  ** AfterUpdate**
    
5.  **Exit** : User moves focus away from control
    



## See also


#### Concepts


[OlkCheckBox Object](olkcheckbox-object-outlook.md)

