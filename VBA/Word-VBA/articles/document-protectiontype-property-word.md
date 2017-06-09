---
title: Document.ProtectionType Property (Word)
keywords: vbawd10.chm158007356
f1_keywords:
- vbawd10.chm158007356
ms.prod: word
api_name:
- Word.Document.ProtectionType
ms.assetid: b11de5a8-8755-293e-88d4-86ce199cb57f
ms.date: 06/08/2017
---


# Document.ProtectionType Property (Word)

Returns the protection type for the specified document. Can be one of the following  **WdProtectionType** constants: **wdAllowOnlyComments** , **wdAllowOnlyFormFields** , **wdAllowOnlyReading** , **wdAllowOnlyRevisions** , or **wdNoProtection** .


## Syntax

 _expression_ . **ProtectionType**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

If the active document isn't already protected, this example protects the document for comments.


```vb
If ActiveDocument.ProtectionType = wdNoProtection Then 
 ActiveDocument.Protect Type:=wdAllowOnlyComments 
End If
```

This example unprotects the active document if it is protected.




```vb
Set Doc = ActiveDocument 
If Doc.ProtectionType <> wdNoProtection Then Doc.Unprotect
```


## See also


#### Concepts


[Document Object](document-object-word.md)

