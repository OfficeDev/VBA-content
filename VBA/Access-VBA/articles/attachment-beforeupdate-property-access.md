---
title: Attachment.BeforeUpdate Property (Access)
keywords: vbaac10.chm13937
f1_keywords:
- vbaac10.chm13937
ms.prod: access
api_name:
- Access.Attachment.BeforeUpdate
ms.assetid: 44a17114-bbb6-8ec9-89b5-db09cf60de98
ms.date: 06/08/2017
---


# Attachment.BeforeUpdate Property (Access)

Returns or sets which macro, event procedure, or user-defined function runs when the  **BeforeUpdate** event occurs. Read/write **String**.


## Syntax

 _expression_. **BeforeUpdate**

 _expression_ A variable that represents an **Attachment** object.


## Remarks

Valid values for this property are " _macroname_" where  _macroname_ is the name of macro; "[Event Procedure]" which indicates the event procedure associated with the **BeforeUpdate** event for the specified object; or " **=** _functionname_ **()** " where _functionname_ is the name of a user-defined function.


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

