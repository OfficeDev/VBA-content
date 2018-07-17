---
title: Presentation.PasswordEncryptionKeyLength Property (PowerPoint)
keywords: vbapp10.chm583077
f1_keywords:
- vbapp10.chm583077
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.PasswordEncryptionKeyLength
ms.assetid: 4a3d59e4-fd4d-cd8d-8d51-cca6ebd4b758
ms.date: 06/08/2017
---


# Presentation.PasswordEncryptionKeyLength Property (PowerPoint)

Returns the key length of the algorithm Microsoft PowerPoint uses when it encrypts documents with passwords. Read-only.


## Syntax

 _expression_. **PasswordEncryptionKeyLength**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

Long


## Remarks

Use the  **[SetPasswordEncryptionOptions](presentation-setpasswordencryptionoptions-method-powerpoint.md)** method to specify the algorithm PowerPoint uses for encrypting documents with passwords.


## Example

This example sets the password encryption options if the password encryption key length is less than 40.


```vb
Sub PasswordSettings() 
    With ActivePresentation 
        If .PasswordEncryptionKeyLength < 40 Then 
            .SetPasswordEncryptionOptions  


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


[Presentation Object](presentation-object-powerpoint.md)

