---
title: OlkTextBox.BeforeUpdate Event (Outlook)
keywords: vbaol11.chm1000083
f1_keywords:
- vbaol11.chm1000083
ms.prod: outlook
api_name:
- Outlook.OlkTextBox.BeforeUpdate
ms.assetid: bf65f754-abcd-8a0a-53de-2c2ee76d27f8
ms.date: 06/08/2017
---


# OlkTextBox.BeforeUpdate Event (Outlook)

Occurs when the data in the control is changed through the user interface and is about to be saved to the item. 


## Syntax

 _expression_ . **BeforeUpdate**( **_Cancel_** )

 _expression_ A variable that represents an **OlkTextBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the operation will not be completed and the property bound to the control will not be updated.|

## Remarks

Canceling this property will revert the control to the current value of the property and return the focus to the control.

 **BeforeUpdate** and **AfterUpdate** can occur any time the data in the control is being saved to the item. The typical sequence of events involving **BeforeUpdate** for this control is as follows:


1. User focuses on the control
    
2.  **BeforeUpdate**
    
3. Control data is updated
    
4.  ** AfterUpdate**
    
5.  **Exit** : User moves focus away from control
    



## See also


#### Concepts


[OlkTextBox Object](olktextbox-object-outlook.md)

