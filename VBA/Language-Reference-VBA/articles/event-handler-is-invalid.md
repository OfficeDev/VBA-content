---
title: Event handler is invalid
keywords: vblr6.chm1107946
f1_keywords:
- vblr6.chm1107946
ms.prod: office
ms.assetid: 98127960-e85a-0d89-ac7c-8e0f6bff8adf
ms.date: 06/08/2017
---


# Event handler is invalid

The [parameter](vbe-glossary.md) list of an event-handling[procedure](vbe-glossary.md) must precisely match the declaration of the event. This error has the following cause and solution:



- Your event-handling procedure has the wrong number of parameters. Eliminate extra parameters or add the missing ones.
    
- One or more of your event-handling procedure parameters has the wrong [data type](vbe-glossary.md).
    
    Make the parameter types match those of the event declaration.
    
- Your event-handling procedure is a  **Function** rather than a **Sub**. Make your procedure a **Sub**. An event handler can't return a value.
    
- Another [type library](vbe-glossary.md) uses the event name for a type of its own.
    
    Qualify the name with the name of the proper type library to avoid the ambiguity.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

