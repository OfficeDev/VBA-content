---
title: Email Object (Word)
keywords: vbawd10.chm2525
f1_keywords:
- vbawd10.chm2525
ms.prod: word
api_name:
- Word.Email
ms.assetid: ee23a74e-556b-04d8-f0b9-fb95f7aa8cfc
ms.date: 06/08/2017
---


# Email Object (Word)

Represents an e-mail message.


## Remarks

Use the  **[Email](document-email-property-word.md)** property to return the **Email** object. The **Email** object and its properties are valid only if the active document is an unsent forward, reply, or new e-mail message.

This example displays the name of the style associated with the current e-mail author.




```vb
MsgBox ActiveDocument.Email _ 
 .CurrentEmailAuthor.Style.NameLocal
```

The author style name is the same as the value returned by the  **[UserName](application-username-property-word.md)** property.


 **Note**   There is no Emails collection; each **Document** object contains only one **Email** object.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


