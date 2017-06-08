---
title: SignatureInfo Members (Office)
ms.prod: office
ms.assetid: 52c19097-8afb-d35c-a9f7-eae81e91c05d
ms.date: 06/08/2017
---


# SignatureInfo Members (Office)
Represents the information used to create a digital or in-document signature.

Represents the information used to create a digital or in-document signature.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[GetCertificateDetail](signatureinfo-getcertificatedetail-method-office.md)|Displays a specified detail related to a digital certificate.|
|[GetSignatureDetail](signatureinfo-getsignaturedetail-method-office.md)|Displays a specified detail related to a signature.|
|[SelectCertificateDetailByThumbprint](signatureinfo-selectcertificatedetailbythumbprint-method-office.md)|Displays a dialog box containing information about a digital certificate following vertification of the user from a thumbprint.|
|[SelectSignatureCertificate](signatureinfo-selectsignaturecertificate-method-office.md)|Displays a dialog box that allows users to select which signature certificate to use for signing a document.|
|[ShowSignatureCertificate](signatureinfo-showsignaturecertificate-method-office.md)|Displays the selected or default digital certificate. |

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](signatureinfo-application-property-office.md)|Gets an  **Application** object that represents the container application for the **SignatureInfo** object. Read-only.|
|[CertificateVerificationResults](signatureinfo-certificateverificationresults-property-office.md)|Gets the results from the verification of a digital certificate. Read-only.|
|[ContentVerificationResults](signatureinfo-contentverificationresults-property-office.md)|Gets a value representing the results of the verification of the hashed contents of a signed document. Read-only.|
|[Creator](signatureinfo-creator-property-office.md)|Gets a 32-bit integer that indicates the application in which the  **SignatureInfo** object was created. Read-only.|
|[IsCertificateExpired](signatureinfo-iscertificateexpired-property-office.md)|Gets a  **Boolean** value indicating whether the digital certificate is expired. Read-only.|
|[IsCertificateRevoked](signatureinfo-iscertificaterevoked-property-office.md)|Gets a  **Boolean** value indicating whether the digital certificate is revoked. Read-only.|
|[IsCertificateUntrusted](signatureinfo-iscertificateuntrusted-property-office.md)|Gets a  **Boolean** value indicating whether the digital certificate used to digitally sign a document comes from a trusted source. Read-only.|
|[IsValid](signatureinfo-isvalid-property-office.md)|Gets a  **Boolean** value indicating whether the signature was successfully validated following signature verification. Read-only.|
|[ReadOnly](signatureinfo-readonly-property-office.md)|Gets a  **Boolean** value indicating whether the **SignatureInfo** object is read-only. Read-only.|
|[SignatureComment](signatureinfo-signaturecomment-property-office.md)|Gets or sets a value containing comments included in a signature packet. Read/write.|
|[SignatureImage](signatureinfo-signatureimage-property-office.md)|Gets or sets the value of the image used to sign the document. Read/write.|
|[SignatureProvider](signatureinfo-signatureprovider-property-office.md)|Gets a value identifying an installed signature provider add-in. Read-only.|
|[SignatureText](signatureinfo-signaturetext-property-office.md)|Gets or sets the value of the signature text used to sign this document. Read/write.|

