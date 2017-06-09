---
title: CanRedo Property
keywords: fm20.chm2000860
f1_keywords:
- fm20.chm2000860
ms.prod: office
api_name:
- Office.CanRedo
ms.assetid: 18b4b51d-3a8a-e03d-14b2-b262f6a12c78
ms.date: 06/08/2017
---


# CanRedo Property



Indicates whether the most recent Undo can be reversed.
 **Syntax**
 _object_. **CanRedo**
The  **CanRedo** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Return Values**
The  **CanRedo** property return values are:


|**Value**|**Description**|
|:-----|:-----|
|**True**|The most recent Undo can be reversed.|
|**False**|The most recent Undo is irreversible.|
 **Remarks**
 **CanRedo** is read-only.
To Redo an action means to reverse an Undo; it does not necessarily mean to repeat the last user action.
The following user actions illustrate using Undo and Redo:


1. Change the setting of an option button.
    
2. Enter text into a text box.
    
3. Click Undo. The text disappears from the text box.
    
4. Click Undo. The option button reverts to its previous setting.
    
5. Click Redo. The value of the option button changes.
    
6. Click Redo. The text reappears in the text box.
    


