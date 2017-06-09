---
title: Workbook.PasswordEncryptionAlgorithm Property (Excel)
keywords: vbaxl10.chm199212
f1_keywords:
- vbaxl10.chm199212
ms.prod: excel
api_name:
- Excel.Workbook.PasswordEncryptionAlgorithm
ms.assetid: 2745a8da-2a61-b949-115a-7f1112a0289e
ms.date: 06/08/2017
---


# Workbook.PasswordEncryptionAlgorithm Property (Excel)

Returns a  **String** indicating the algorithm Microsoft Excel uses to encrypt passwords for the specified workbook. Read-only.


## Syntax

 _expression_ . **PasswordEncryptionAlgorithm**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

Use the  **[SetPasswordEncryptionOptions](workbook-setpasswordencryptionoptions-method-excel.md)** method to specify whether Excel encrypts file properties for password-protected workbooks.


## Example

This example sets the password encryption options for the active workbook.


```vb
Sub SetPasswordOptions() 
 
 ActiveWorkbook.SetPasswordEncryptionOptions _ 
 PasswordEncryptionProvider:="Microsoft RSA SChannel Cryptographic Provider", _ 
 PasswordEncryptionAlgorithm:="RC4", _ 
 PasswordEncryptionKeyLength:=56, _ 
 PasswordEncryptionFileProperties:=True 
 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

