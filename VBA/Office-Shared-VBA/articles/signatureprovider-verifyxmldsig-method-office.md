---
title: SignatureProvider.VerifyXmlDsig Method (Office)
keywords: vbaof11.chm287006
f1_keywords:
- vbaof11.chm287006
ms.prod: office
api_name:
- Office.SignatureProvider.VerifyXmlDsig
ms.assetid: 8b72f282-ace5-4b51-e90a-e2df79affcb1
ms.date: 06/08/2017
---


# SignatureProvider.VerifyXmlDsig Method (Office)

Verifies a signature based on the signed state of the document and the legitimacy of the certificate used for signing.


## Syntax

 _expression_. **VerifyXmlDsig**( **_QueryContinue_**, **_psigsetup_**, **_psiginfo_**, **_XmlDsigStream_**, **_pcontverres_**, **_pcertverres_** )

 _expression_ An expression that returns a **SignatureProvider** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _QueryContinue_|Required|**IQueryContinue**|Provides a way to query the host application for permission to continue the verification operation.|
| _psigsetup_|Required|**SignatureSetup**|Specifies configuration information about a signature line.|
| _psiginfo_|Required|**SignatureInfo**|Specifies information captured from the signing ceremony.|
| _XmlDsigStream_|Required|**IStream**|Represents a steam of data containing XML, which represents an  **XMLDSIG** object.|
| _pcontverres_|Required|**ContentVerificationResults**|Specifies the status of the signature verification action.|
| _pcertverres_|Required|**CertificateVerificationResults**|Specifies the status of the signing certificate verification.|

## Remarks

XMLDSIG is a standards-based signature format (http://www.w3.org/TR/xmldsig-core/), verifiable by third parties. This is the default format for signatures in Microsoft Office.


## Example

The following example, written in C#, shows the implementation of the  **VerifyXmlDsig** method in a custom signature provider project.


```
 public void VerifyXmlDsig(object queryContinue, SignatureSetup sigsetup, SignatureInfo siginfo, object xmldsigStream, ref ContentVerificationResults contverresults, ref CertificateVerificationResults certverresults) 
 { 
 using (COMStream comstream = new COMStream(xmldsigStream)) 
 { 
 XmlDocument xmldsig = new XmlDocument(); 
 xmldsig.PreserveWhitespace = true; 
 xmldsig.Load(comstream); 
 
 XmlElement signature = xmldsig.DocumentElement; 
 SignedXml signedXml = new SignedXml(); 
 signedXml.LoadXml(signature); 
 
 contverresults = signedXml.CheckSignature() ? 
 Microsoft.Office.Core.ContentVerificationResults.contverresValid : 
 Microsoft.Office.Core.ContentVerificationResults.contverresModified; 
 } 
 }
```


 **Note**  Signature providers are implemented exclusively in custom COM add-ins created in managed and unmanaged code and cannot be implemented in Microsoft Visual Basic® for Applications (VBA). 


## See also


#### Concepts


[SignatureProvider Object](signatureprovider-object-office.md)
#### Other resources


[SignatureProvider Object Members](signatureprovider-members-office.md)

