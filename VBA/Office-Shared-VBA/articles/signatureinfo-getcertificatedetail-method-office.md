---
title: SignatureInfo.GetCertificateDetail Method (Office)
keywords: vbaof11.chm286007
f1_keywords:
- vbaof11.chm286007
ms.prod: office
api_name:
- Office.SignatureInfo.GetCertificateDetail
ms.assetid: f3cab134-5560-be37-25b4-2cbbfcf0693e
ms.date: 06/08/2017
---


# SignatureInfo.GetCertificateDetail Method (Office)

Displays a specified detail related to a digital certificate.


## Syntax

 _expression_. **GetCertificateDetail**( **_certdet_** )

 _expression_ An expression that returns a **SignatureInfo** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _certdet_|Required|**CertificateDetail**|An enumerated value specifying which certificate detail to display.|

### Return Value

Variant


## Example

The following example gets the expiration date of the digital certificate.


```
Sub GetCertDetails() 
Dim objSignatureInfo As SignatureInfo 
Dim varDetail As Variant 
 
strDetail = objSignatureInfo.GetCertificationDetail(certdetExpirationDate) 
 
End Sub 

```


## See also


#### Concepts


[SignatureInfo Object](signatureinfo-object-office.md)
#### Other resources


[SignatureInfo Object Members](signatureinfo-members-office.md)

