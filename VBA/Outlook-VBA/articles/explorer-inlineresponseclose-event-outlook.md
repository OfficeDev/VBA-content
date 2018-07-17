---
title: Explorer.InlineResponseClose Event (Outlook)
keywords: vbaol11.chm3598
f1_keywords:
- vbaol11.chm3598
ms.assetid: ff3f3286-995a-409c-aca5-706290e26252
ms.date: 06/08/2017
ms.prod: outlook
---


# Explorer.InlineResponseClose Event (Outlook)
Occurs when the user performs an action that causes the active inline response to close in the Reading Pane.

## Syntax

 _expression_ . **InlineResponseClose**_(Item)_

 _expression_ A variable that represents an **[Explorer](explorer-object-outlook.md)** object.


## Remarks

This event fires when a new inline response or draft inline response is closed for the following reasons:


- The user selects the  **Pop Out** command.
    
    The user selects the  **Discard** command.
    
    The user sends the inline response.
    
    The user navigates to a different message in the message list.
    
    The user navigates to a different folder. 
    
    The user switches modules, for example from the mail module to the Calendar module.
    

## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

