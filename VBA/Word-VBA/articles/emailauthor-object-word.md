---
title: EmailAuthor Object (Word)
keywords: vbawd10.chm2519
f1_keywords:
- vbawd10.chm2519
ms.prod: word
api_name:
- Word.EmailAuthor
ms.assetid: 2749e018-42e9-7a1a-f18b-8605b38ff0ae
ms.date: 06/08/2017
---


# EmailAuthor Object (Word)

Represents the author of an e-mail message.


## Remarks

Use the  **[CurrentEmailAuthor](email-currentemailauthor-property-word.md)** property to return the **EmailAuthor** object. The **EmailAuthor** object and its properties are valid only if the active document is an unsent forward, reply, or new e-mail message.

This example returns the style associated with the current author for unsent replies, forwards, or new e-mail messages, and displays the name of the font associated with this style.




```vb
Set MyEmailStyle = _ 
 ActiveDocument.Email.CurrentEmailAuthor.Style 
Msgbox MyEmailStyle.Font.Name
```


 **Note**  There is no EmailAuthors collection; each  **Email** object contains only one **EmailAuthor** object.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


