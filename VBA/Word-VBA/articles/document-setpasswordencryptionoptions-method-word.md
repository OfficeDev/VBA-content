---
title: Document.SetPasswordEncryptionOptions Method (Word)
keywords: vbawd10.chm158007657
f1_keywords:
- vbawd10.chm158007657
ms.prod: word
api_name:
- Word.Document.SetPasswordEncryptionOptions
ms.assetid: 4e7c2c0a-cac2-6fa3-f237-f02c897757a1
ms.date: 06/08/2017
---


# Document.SetPasswordEncryptionOptions Method (Word)

Sets the options Microsoft Word uses for encrypting documents with passwords.


## Syntax

 _expression_ . **SetPasswordEncryptionOptions**( **_PasswordEncryptionProvider_** , **_PasswordEncryptionAlgorithm_** , **_PasswordEncryptionKeyLength_** , **_PasswordEncryptionFileProperties_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PasswordEncryptionProvider_|Required| **String**|The name of the encryption provider.|
| _PasswordEncryptionAlgorithm_|Required| **String**|The name of the encryption algorithm. Word supports stream-encrypted algorithms.|
| _PasswordEncryptionKeyLength_|Required| **Long**|The encryption key length. Must be a multiple of 8, starting at 40.|
| _PasswordEncryptionFileProperties_|Optional| **Variant**| **True** for Word to encrypt file properties. Default is **True** .|

## Remarks

For enhanced security, do not use Weak Encryption (XOR) (also called "OfficeXor") or "Office97/2000 Compatible" (also called "OfficeStandard") algorithms.


## Example

This example sets the password encryption to a stronger encryption if the password encryption algorithm in use is "OfficeXor" or "OfficeStandard."


```vb
Sub PasswordSettings() 
 With ActiveDocument 
 If .PasswordEncryptionAlgorithm = "OfficeXor" Or _ 
 .PasswordEncryptionAlgorithm = "OfficeStandard" Then 
 
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

