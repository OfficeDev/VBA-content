---
title: EmailSignature Object (Word)
keywords: vbawd10.chm2524
f1_keywords:
- vbawd10.chm2524
ms.prod: word
api_name:
- Word.EmailSignature
ms.assetid: 9d641321-d52b-ab9a-4117-6f9e11dedbba
ms.date: 06/08/2017
---


# EmailSignature Object (Word)

Contains information about the e-mail signatures used by Microsoft Word when you create and edit e-mail messages and replies.


## Remarks

Use the  **EmailSignature** property to return the **EmailSignature** object.

This example changes the signatures Word appends to new outgoing e-mail messages and e-mail message replies.




```vb
With Application.EmailOptions.EmailSignature 
 .NewMessageSignature = "Signature1" 
 .ReplyMessageSignature = "Reply2" 
End With
```


 **Note**  There is no EmailSignatures collection; each  **[EmailOptions](emailoptions-object-word.md)** object contains only one **EmailSignature** object.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


