---
title: Document.PasswordEncryptionAlgorithm Property (Word)
keywords: vbawd10.chm158007664
f1_keywords:
- vbawd10.chm158007664
ms.prod: word
api_name:
- Word.Document.PasswordEncryptionAlgorithm
ms.assetid: 5317832f-936b-5c3b-5acc-6c067563acd6
ms.date: 06/08/2017
---


# Document.PasswordEncryptionAlgorithm Property (Word)

Returns a  **String** indicating the algorithm Microsoft Word uses for encrypting documents with passwords. Read-only.


## Syntax

 _expression_ . **PasswordEncryptionAlgorithm**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

Use the  **[SetPasswordEncryptionOptions](document-setpasswordencryptionoptions-method-word.md)** method to specify the algorithm Word uses for encrypting documents with passwords.


## Example

This example sets the password encryption options if the password encryption algorithm in use is "OfficeXor," which is the password algorithm used in versions of Word prior to Word 97 for Windows.


```vb
Sub PasswordSettings() 
 With ActiveDocument 
 If .PasswordEncryptionAlgorithm = "OfficeXor" Then 
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

