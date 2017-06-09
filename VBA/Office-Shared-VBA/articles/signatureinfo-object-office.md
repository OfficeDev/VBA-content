---
title: SignatureInfo Object (Office)
keywords: vbaof11.chm286000
f1_keywords:
- vbaof11.chm286000
ms.prod: office
api_name:
- Office.SignatureInfo
ms.assetid: fe0ffe7d-7cc7-0d82-6888-d5eacca0d3ce
ms.date: 06/08/2017
---


# SignatureInfo Object (Office)

Represents the information used to create a digital or in-document signature.


## Example

The following example uses the  **GetCertificationDetails** method of the **SignatureInfo** object to get the expiration date of the digital certificate.


```
Sub GetCertDetails() 
Dim objSignatureInfo As SignatureInfo 
Dim varDetail As Variant 
 
strDetail = objSignatureInfo.GetCertificationDetail(certdetExpirationDate) 
 
End Sub 

```


## Methods



|**Name**|
|:-----|
|[GetCertificateDetail](signatureinfo-getcertificatedetail-method-office.md)|
|[GetSignatureDetail](signatureinfo-getsignaturedetail-method-office.md)|
|[SelectCertificateDetailByThumbprint](signatureinfo-selectcertificatedetailbythumbprint-method-office.md)|
|[SelectSignatureCertificate](signatureinfo-selectsignaturecertificate-method-office.md)|
|[ShowSignatureCertificate](signatureinfo-showsignaturecertificate-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](signatureinfo-application-property-office.md)|
|[CertificateVerificationResults](signatureinfo-certificateverificationresults-property-office.md)|
|[ContentVerificationResults](signatureinfo-contentverificationresults-property-office.md)|
|[Creator](signatureinfo-creator-property-office.md)|
|[IsCertificateExpired](signatureinfo-iscertificateexpired-property-office.md)|
|[IsCertificateRevoked](signatureinfo-iscertificaterevoked-property-office.md)|
|[IsCertificateUntrusted](signatureinfo-iscertificateuntrusted-property-office.md)|
|[IsValid](signatureinfo-isvalid-property-office.md)|
|[ReadOnly](signatureinfo-readonly-property-office.md)|
|[SignatureComment](signatureinfo-signaturecomment-property-office.md)|
|[SignatureImage](signatureinfo-signatureimage-property-office.md)|
|[SignatureProvider](signatureinfo-signatureprovider-property-office.md)|
|[SignatureText](signatureinfo-signaturetext-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
