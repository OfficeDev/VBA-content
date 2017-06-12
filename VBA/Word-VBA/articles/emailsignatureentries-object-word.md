---
title: EmailSignatureEntries Object (Word)
keywords: vbawd10.chm2533
f1_keywords:
- vbawd10.chm2533
ms.prod: word
api_name:
- Word.EmailSignatureEntries
ms.assetid: 42a63f45-f989-be32-e75a-059c9a77c6f1
ms.date: 06/08/2017
---


# EmailSignatureEntries Object (Word)

A collection of  **[EmailSignatureEntry](emailsignatureentry-object-word.md)** objects that represents all the e-mail signature entries available to Word.


## Remarks

Use the  **[EmailSignatureEntries](emailsignature-emailsignatureentries-property-word.md)** property to return the **EmailSignatureEntries** collection. Use the **[Add](emailsignatureentries-add-method-word.md)** method of the **EmailSignatureEntries** object to add an e-mail signature to Word. The following example creates a new e-mail signature entry based on the author's name and a selection in the active document, and then it sets the new signature entry as the default e-mail signature to use for new messages.


```vb
Sub NewEmailSignature() 
 With Application.EmailOptions.EmailSignature 
 .EmailSignatureEntries.Add "Jeff Smith", Selection.Range 
 .NewMessageSignature = "Jeff Smith" 
 End With 
End Sub
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


