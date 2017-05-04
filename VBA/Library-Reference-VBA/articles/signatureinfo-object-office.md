---
title: SignatureInfo Object (Office)
keywords: vbaof11.chm286000
f1_keywords:
- vbaof11.chm286000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SignatureInfo
ms.assetid: fe0ffe7d-7cc7-0d82-6888-d5eacca0d3ce
---


# SignatureInfo Object (Office)

Represents the information used to create a digital or in-document signature.


## Example

The following example uses the  **GetCertificationDetails** method of the **SignatureInfo** object to get the expiration date of the digital certificate.


```vb
Sub GetCertDetails() 
Dim objSignatureInfo As SignatureInfo 
Dim varDetail As Variant 
 
strDetail = objSignatureInfo.GetCertificationDetail(certdetExpirationDate) 
 
End Sub
```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

