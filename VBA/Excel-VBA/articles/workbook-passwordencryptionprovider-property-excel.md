---
title: Workbook.PasswordEncryptionProvider Property (Excel)
keywords: vbaxl10.chm199211
f1_keywords:
- vbaxl10.chm199211
ms.prod: excel
api_name:
- Excel.Workbook.PasswordEncryptionProvider
ms.assetid: d5bcbbf2-8de9-6725-9cac-679d6c023b34
ms.date: 06/08/2017
---


# Workbook.PasswordEncryptionProvider Property (Excel)

Returns a  **String** specifying the name of the algorithm encryption provider that Microsoft Excel uses when encrypting passwords for the specified workbook. Read-only.


## Syntax

 _expression_ . **PasswordEncryptionProvider**

 _expression_ A variable that represents a **Workbook** object.


## Example

This example sets the password encryption options for the specified workbook, if the file properties are not encrypted for password-protected workbooks.


```vb
Sub SetPasswordOptions() 
 
 With ActiveWorkbook 
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


[Workbook Object](workbook-object-excel.md)

