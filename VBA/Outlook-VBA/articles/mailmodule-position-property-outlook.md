---
title: MailModule.Position Property (Outlook)
keywords: vbaol11.chm2818
f1_keywords:
- vbaol11.chm2818
ms.prod: outlook
api_name:
- Outlook.MailModule.Position
ms.assetid: 46cfff8e-07ac-f929-94be-c7b39980daa8
ms.date: 06/08/2017
---


# MailModule.Position Property (Outlook)

Returns or sets a  **Long** value that represents the ordinal position of the **[MailModule](mailmodule-object-outlook.md)** object when it is displayed in the Navigation Pane. Read/write.


## Syntax

 _expression_ . **Position**

 _expression_ A variable that represents a **MailModule** object.


## Remarks

This property can only be set to a value between 1 and 9. An error occurs if you attempt to set this property to a value outside that range.

Changing the value of this property for a given  **MailModule** object changes the **Position** values of other navigaton modules in a **[NavigationModules](navigationmodules-object-outlook.md)** collection, depending on the relative change between the new value and the original value.


- If the new value is less than the original value, the specified  **MailModule** object moves up to the new position and the other navigation modules that are already at or below that new position move down.
    
- If the new value is greater than the original value, the specified  **MailModule** object moves down to the new position and the other navigation modules that are between the old position and the new position move up, filling the old position.
    

## See also


#### Concepts


[MailModule Object](mailmodule-object-outlook.md)

