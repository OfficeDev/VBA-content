---
title: Workbook.PasswordEncryptionKeyLength Property (Excel)
keywords: vbaxl10.chm199213
f1_keywords:
- vbaxl10.chm199213
ms.prod: excel
api_name:
- Excel.Workbook.PasswordEncryptionKeyLength
ms.assetid: 2662f2f5-1ad0-4a75-82c0-3268f147948a
ms.date: 06/08/2017
---


# Workbook.PasswordEncryptionKeyLength Property (Excel)

Returns a  **Long** indicating the key length of the algorithm Microsoft Excel uses when encrypting passwords for the specified workbook. Read-only.


## Syntax

 _expression_ . **PasswordEncryptionKeyLength**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

Use the  **[SetPasswordEncryptionOptions](workbook-setpasswordencryptionoptions-method-excel.md)** method to specify whether Excel encrypts file properties for the specified password-protected workbook.


## Example

This example sets the password encryption options for the specified workbook, if the password encryption key length is less than 56.


```vb
Sub SetPasswordOptions() 
 
 With ActiveWorkbook 
 If .PasswordEncryptionKeyLength < 56 Then 
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

