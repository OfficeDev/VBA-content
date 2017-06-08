---
title: SignatureSetup Object (Office)
keywords: vbaof11.chm285000
f1_keywords:
- vbaof11.chm285000
ms.prod: office
api_name:
- Office.SignatureSetup
ms.assetid: e76b87c9-3163-654c-ab52-559dfdf43c90
ms.date: 06/08/2017
---


# SignatureSetup Object (Office)

Represents the information used to set up a signature packet.


## Example

The following example sets various properties of the  **SignatureSetup** object for a signature packet.


```
Dim objSigSetup As SignatureSetup 
With objSigSetup 
.AllowComments = True 
.ShowSignDate = True 
.SigningInstructions = "Please sign this document." 
.SuggestedSignerEmail = "jdow@example.com" 
Next
```


## Properties



|**Name**|
|:-----|
|[AdditionalXml](signaturesetup-additionalxml-property-office.md)|
|[AllowComments](signaturesetup-allowcomments-property-office.md)|
|[Application](signaturesetup-application-property-office.md)|
|[Creator](signaturesetup-creator-property-office.md)|
|[Id](signaturesetup-id-property-office.md)|
|[ReadOnly](signaturesetup-readonly-property-office.md)|
|[ShowSignDate](signaturesetup-showsigndate-property-office.md)|
|[SignatureProvider](signaturesetup-signatureprovider-property-office.md)|
|[SigningInstructions](signaturesetup-signinginstructions-property-office.md)|
|[SuggestedSigner](signaturesetup-suggestedsigner-property-office.md)|
|[SuggestedSignerEmail](signaturesetup-suggestedsigneremail-property-office.md)|
|[SuggestedSignerLine2](signaturesetup-suggestedsignerline2-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
