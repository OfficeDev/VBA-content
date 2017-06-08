---
title: WhatsThisHelp Property
keywords: vblr6.chm916695
f1_keywords:
- vblr6.chm916695
ms.prod: office
api_name:
- Office.WhatsThisHelp
ms.assetid: f36a9ddc-c0d3-c2d7-8cf8-03b49bd00679
ms.date: 06/08/2017
---


# WhatsThisHelp Property



Returns a [Boolean](vbe-glossary.md) value that determines whether context-sensitive Help uses the pop-up window provided by Windows 95 Help or the main Help window. Read-only at[run time](vbe-glossary.md). This property does not apply to the Macintosh.
 **Remarks**
The settings for the  **WhatsThisHelp** property are:


|**Setting**|**Description**|
|:-----|:-----|
|**True**|The application uses one of the What's This access techniques to start Windows Help and load a topic identified by the  **WhatsThisHelpID** property.|
|**False**|(Default) The application uses the F1 key to start Windows Help and load the topic identified by the  **HelpContextID** property.|
 **Remarks**
There are two access techniques for providing What's This Help in an application. The  **WhatsThisHelp** property must be set to **True** for any of these techniques to work.


- Providing a  **What's This** button on the title bar of the **UserForm** using the **WhatsThisButton** property. The mouse pointer changes to the What's This state (arrow with question mark). The topic displayed is identified by the **WhatsThisHelpID** property of the control clicked by the user.
    
- Invoking the  **WhatsThisMode** method of a **UserForm**. This produces the same behavior as clicking the **What's This** button without using a button. For example, you can invoke this method from a command on a menu in the menu bar of the application.
    


