---
title: SignatureProvider.SignXmlDsig Method (Office)
keywords: vbaof11.chm287004
f1_keywords:
- vbaof11.chm287004
ms.prod: office
api_name:
- Office.SignatureProvider.SignXmlDsig
ms.assetid: d278f48f-4128-b8b1-f32d-d81ccbbf6771
ms.date: 06/08/2017
---


# SignatureProvider.SignXmlDsig Method (Office)

Used to sign the XMLDSIG template.


## Syntax

 _expression_. **SignXmlDsig**( **_QueryContinue_**, **_psigsetup_**, **_psiginfo_**, **_XmlDsigStream_** )

 _expression_ An expression that returns a **SignatureProvider** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _QueryContinue_|Required|**IQueryContinue**|Provides a way to query the host application for permission to continue the verification operation.|
| _psigsetup_|Required|**SignatureSetup**|Specifies configuration information about a signature line.|
| _psiginfo_|Required|**SignatureInfo**|Specifies information captured from the signing ceremony.|
| _XmlDsigStream_|Required|**IStream**|Represents a steam of data containing XML, which represents an  **XMLDSIG** object.|

## Remarks

XMLDSIG is a standards-based signature format (http://www.w3.org/TR/xmldsig-core/), verifiable by third parties. This is the default format for signatures in the Microsoft Office. 


## Example

The following example, written in C#, shows the implementation of the  **SignXmlDsig** method in a custom signature provider project.


```
 public void SignXmlDsig(object queryContinue, SignatureSetup sigsetup, SignatureInfo siginfo, object xmldsigStream) 
 { 
 using (COMStream comstream = new COMStream(xmldsigStream)) 
 { 
 XmlDocument xmldsig = new XmlDocument(); 
 xmldsig.PreserveWhitespace = true; 
 xmldsig.Load(comstream); 
 
 XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmldsig.NameTable); 
 nsmgr.AddNamespace("ds", "http://www.w3.org/2000/09/xmldsig#"); 
 
 XmlElement signature = xmldsig.DocumentElement; 
 SignedXml signedXml = new SignedXml(); 
 signedXml.LoadXml(signature); 
 
 // Cert 
 X509Certificate2 cert = TestSignatureProvider.GetSigningCertificate(siginfo); 
 KeyInfo keyInfo = new KeyInfo(); 
 if (cert.PrivateKey is RSA) 
 keyInfo.AddClause(new RSAKeyValue((RSA) cert.PrivateKey)); 
 else if (cert.PrivateKey is DSA) 
 keyInfo.AddClause(new DSAKeyValue((DSA) cert.PrivateKey)); 
 keyInfo.AddClause(new KeyInfoX509Data(cert)); 
 signedXml.SigningKey = cert.PrivateKey; 
 signedXml.KeyInfo = keyInfo; 
 
 // Compute signature 
 signedXml.ComputeSignature(); 
 
 // Copy data from signed signature 
 // REVIEW: Cleaner way to do this? 
 string[] xpathsToCopy = new string[] 
 { 
 "./ds:SignedInfo", 
 "./ds:SignatureValue", 
 "./ds:KeyInfo", 
 }; 
 XmlElement signedSignature = signedXml.GetXml(); 
 foreach (string xpathToCopy in xpathsToCopy) 
 { 
 signature.ReplaceChild( 
 xmldsig.ImportNode(signedSignature.SelectSingleNode(xpathToCopy, nsmgr), true), 
 signature.SelectSingleNode(xpathToCopy, nsmgr)); 
 } 
 
 // Save signature back to stream 
 comstream.SetLength(0); 
 comstream.Position = 0; 
 xmldsig.Save(new XmlTextWriter(comstream, new UTF8Encoding(false))); 
 } 
 }
```


 **Note**  Signature providers are implemented exclusively in custom COM add-ins created in managed and unmanaged code and cannot be implemented in Microsoft Visual Basic for Applications (VBA). 


## See also


#### Concepts


[SignatureProvider Object](signatureprovider-object-office.md)
#### Other resources


[SignatureProvider Object Members](signatureprovider-members-office.md)

