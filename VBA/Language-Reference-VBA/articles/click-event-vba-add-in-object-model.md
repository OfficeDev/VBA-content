---
title: Click Event (VBA Add-In Object Model)
keywords: vbob6.chm1098932
f1_keywords:
- vbob6.chm1098932
ms.prod: office
ms.assetid: ac72bf41-e582-be84-d204-96545e8db71e
ms.date: 06/08/2017
---


# Click Event (VBA Add-In Object Model)



Occurs when the  **OnAction**[property](vbe-glossary.md) of a corresponding command bar control is set.
 **Syntax**
 **Sub**_object_**_Click (ByVal** **_ctrl_** **As Object**, **ByRef** **_handled_** **As Boolean**, **ByRef** **_canceldefault_** **As Boolean)**
The Click event syntax has these [named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_ctrl_**|Required; [Object](vbe-glossary.md). Specifies the object that is the source of the Click event.|
|**_handled_**|Required; [Boolean](vbe-glossary.md). If  **True**, other[add-ins](vbe-glossary.md) should handle the event. If **False**, the action of the command bar item has not been handled.|
|**_canceldefault_**|Required;  **Boolean**. If **True**, default behavior is performed unless canceled by a downstream add-in. If **False**, default behavior is not performed unless restored by a downstream add-in.|
 **Remarks**
The Click event is specific to the  **CommandBarEvents** object. Use a[variable](vbe-glossary.md) declared using the **WithEvents** keyword to receive the Click event for a **CommandBar** control. This variable should be set to the return value of the **CommandBarEvents** property of the **Events** object. The **CommandBarEvents** property takes the **CommandBar** control as an[argument](vbe-glossary.md). When the  **CommandBar** control is clicked (for the variable you declared using the **WithEvents** keyword), the code is executed.

