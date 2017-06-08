---
title: Presentation.PasswordEncryptionProvider Property (PowerPoint)
keywords: vbapp10.chm583075
f1_keywords:
- vbapp10.chm583075
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.PasswordEncryptionProvider
ms.assetid: 055d4972-a835-f3fb-24df-9f275374ea6e
ms.date: 06/08/2017
---


# Presentation.PasswordEncryptionProvider Property (PowerPoint)

Returns the name of the algorithm encryption provider that Microsoft PowerPoint uses when it encrypts documents with passwords. Read-only.


## Syntax

 _expression_. **PasswordEncryptionProvider**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

String


## Remarks

Use the  **[SetPasswordEncryptionOptions](presentation-setpasswordencryptionoptions-method-powerpoint.md)** method to specify the algorithm PowerPoint uses for encrypting documents with passwords.


## Example

This example sets the password encryption options if the password encryption algorithm in use is not the Microsoft RSA SChannel Cryptographic Provider.


```vb
Sub PasswordSettings()

    With ActivePresentation
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


[Presentation Object](presentation-object-powerpoint.md)

