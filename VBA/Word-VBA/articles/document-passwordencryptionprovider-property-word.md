---
title: Document.PasswordEncryptionProvider Property (Word)
keywords: vbawd10.chm158007663
f1_keywords:
- vbawd10.chm158007663
ms.prod: word
api_name:
- Word.Document.PasswordEncryptionProvider
ms.assetid: 473e7599-4c04-4a29-6d5c-70228900dedf
ms.date: 06/08/2017
---


# Document.PasswordEncryptionProvider Property (Word)

Returns a  **String** specifying the name of the algorithm encryption provider that Microsoft Word uses when encrypting documents with passwords. Read-only.


## Syntax

 _expression_ . **PasswordEncryptionProvider**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

Use the  **[SetPasswordEncryptionOptions](document-setpasswordencryptionoptions-method-word.md)** method to specify the name of the algorithm encryption provider Word uses when encrypting documents with passwords.


## Example

This example sets the password encryption options if the password encryption algorithm in use is not "Microsoft RSA SChannel Cryptographic Provider."


```vb
Sub PasswordSettings() 
 With ActiveDocument 
 If .PasswordEncryptionProvider <> "Microsoft RSA SChannel Cryptographic Provider" Then 
 .SetPasswordEncryptionOptions _ 
 PasswordEncryptionProvider:="Microsoft RSA SChannel Cryptographic Provider", _ 
 PasswordEncryptionAlgorithm:="RC4", _ 
 PasswordEncryptionKeyLength:=56, _ 
 PasswordEncryptionFileProperties:=True 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

