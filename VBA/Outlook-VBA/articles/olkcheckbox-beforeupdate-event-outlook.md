---
title: OlkCheckBox.BeforeUpdate Event (Outlook)
keywords: vbaol11.chm1000161
f1_keywords:
- vbaol11.chm1000161
ms.prod: outlook
api_name:
- Outlook.OlkCheckBox.BeforeUpdate
ms.assetid: e12072d3-cd24-ce5d-0738-80d44a9c9154
ms.date: 06/08/2017
---


# OlkCheckBox.BeforeUpdate Event (Outlook)

Occurs when the data in the control is changed through the user interface and is about to be saved to the item. 


## Syntax

_expression_. **BeforeUpdate** (**_Cancel_**)

_expression_ A variable that represents an **OlkCheckBox** object.


### Parameters

|Name|Required/Optional|Data Type|Description|
|:-----|:-----|:-----|:-----|
|_Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the operation will not be completed, and the property bound to the control will not be updated.|

<br/>

## Remarks

Canceling this property will revert the control to the current value of the property and return the focus to the control.

**BeforeUpdate** and **AfterUpdate** can occur any time the data in the control is being saved to the item. The typical sequence of events involving **BeforeUpdate** for this control is as follows:

1. User focuses on the control.
    
2.  **BeforeUpdate**
    
3. Control data is updated.
    
4.  **AfterUpdate**
    
5.  **Exit**: User moves focus away from control.
    
    
## See also

#### Concepts

- [OlkCheckBox Object](olkcheckbox-object-outlook.md)

