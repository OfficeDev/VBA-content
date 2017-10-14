---
title: EmailOptions Object (Word)
ms.prod: word
api_name:
- Word.EmailOptions
ms.assetid: 41fefa03-c993-e218-0f92-0cf30c0bfbd4
ms.date: 06/08/2017
---


# EmailOptions Object (Word)

Contains global application-level attributes used by Microsoft Word when you create and edit e-mail messages and replies.


## Remarks

Use the  **[EmailOptions](application-emailoptions-property-word.md)** property to return the **EmailOptions** object.

This example changes the font color of the default style used to compose new e-mail messages.




```vb
Application.EmailOptions.ComposeStyle.Font.Color = _ 
 wdColorBrightGreen
```

This example sets Word to mark comments in e-mail messages with the initials "WK."




```vb
Application.EmailOptions.MarkCommentsWith = "WK" 
Application.EmailOptions.MarkComments = True
```

This example changes the signatures Word appends to new outgoing e-mail messages and e-mail message replies.




```vb
With Application.EmailOptions.EmailSignature 
 .NewMessageSignature = "Signature1" 
 .ReplyMessageSignature = "Reply2" 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


