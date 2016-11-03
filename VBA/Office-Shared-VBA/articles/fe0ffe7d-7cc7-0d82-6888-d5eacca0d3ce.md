
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
|[GetCertificateDetail](f3cab134-5560-be37-25b4-2cbbfcf0693e.md)|
|[GetSignatureDetail](77a5a835-cc8a-0341-8e5d-6ddb603f9517.md)|
|[SelectCertificateDetailByThumbprint](997010ee-330f-433d-c62c-bf211b8351d6.md)|
|[SelectSignatureCertificate](acf3993f-85b3-a455-e3ee-1a713e7787c6.md)|
|[ShowSignatureCertificate](8fef7299-e110-b0a2-7a0c-552e9068e001.md)|

## Properties



|**Name**|
|:-----|
|[Application](98544420-0b08-3fc4-50cd-a787f52450ae.md)|
|[CertificateVerificationResults](dc661f7e-f02e-79a6-91d6-c124109c6d4c.md)|
|[ContentVerificationResults](18fd1338-1554-7bc6-a947-c3ea1123a38f.md)|
|[Creator](57a91318-cdf5-edd0-a1df-5cfdde1e7293.md)|
|[IsCertificateExpired](22f61a5b-809f-718e-926b-a3c6bc9691f1.md)|
|[IsCertificateRevoked](e68c5c54-19a4-c0ef-21c3-c8b5248d86d2.md)|
|[IsCertificateUntrusted](c52041d5-2522-7656-5a40-4b0f3035005d.md)|
|[IsValid](71c2a187-85c7-430f-626d-5dd055ae33dc.md)|
|[ReadOnly](047fe3f8-825b-ae30-ba8d-adcb434b20d3.md)|
|[SignatureComment](2cd03ccf-4291-ff80-ef13-4c03590aa10b.md)|
|[SignatureImage](4a0fa820-5e65-36c6-1f0c-d5d98c4e8fb1.md)|
|[SignatureProvider](e426f4c6-95f7-dc3f-752d-0fee56bc2c65.md)|
|[SignatureText](09b6b780-aa04-32fd-bb13-a2202f5e7cb6.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)