---
title: SignatureProvider.ShowSignatureDetails Method (Office)
keywords: vbaof11.chm287007
f1_keywords:
- vbaof11.chm287007
ms.prod: office
api_name:
- Office.SignatureProvider.ShowSignatureDetails
ms.assetid: ea334547-af85-6d80-75dc-ddee3ad3a2c7
ms.date: 06/08/2017
---


# SignatureProvider.ShowSignatureDetails Method (Office)

Provides a signature povider add-in the opportunity to display details about a signed signature line and display additional stored information such as a secure time-stamp.


## Syntax

 _expression_. **ShowSignatureDetails**( **_ParentWindow_**, **_psigsetup_**, **_psiginfo_**, **_XmlDsigStream_**, **_pcontverres_**, **_pcertverres_** )

 _expression_ An expression that returns a **SignatureProvider** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ParentWindow_|Required|**IOleWindow**|Contains the handle to the window containing the signature details.|
| _psigsetup_|Required|**SignatureSetup**|Specifies initial settings of the signature provider.|
| _psiginfo_|Required|**SignatureInfo**|Specifies information about the signed signature line.|
| _XmlDsigStream_|Required|**IStream**|Represents a steam of data or binary large object of XML.|
| _pcontverres_|Required|**ContentVerificationResults**|Contains a value representing the results of verificating the signature content.|
| _pcertverres_|Required|**CertificateVerificationResults**|Contains a value representing the results of verificating the signing certification.|

## Example

The following example, written in C#, shows the implementation of the  **ShowSignatureDetails** method in a custom signature provider project.


```
 public void ShowSignatureDetails(object parentWindow, SignatureSetup sigsetup, SignatureInfo siginfo, object xmldsigStream, ref ContentVerificationResults contverresults, ref CertificateVerificationResults certverresults) 
 { 
 using (Win32WindowFromOleWindow window = new Win32WindowFromOleWindow(parentWindow)) 
 { 
 using (SigningCeremonyForm signForm = new SigningCeremonyForm(sigsetup, siginfo)) 
 { 
 signForm.ShowDialog(window); 
 } 
 } 
 } 
 

```


 **Note**  Signature providers are implemented exclusively in custom COM add-ins and cannot be implemented in Microsoft Visual Basic for Applications (VBA). 


## See also


#### Concepts


[SignatureProvider Object](signatureprovider-object-office.md)
#### Other resources


[SignatureProvider Object Members](signatureprovider-members-office.md)

