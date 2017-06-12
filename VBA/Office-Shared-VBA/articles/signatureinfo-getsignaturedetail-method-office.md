---
title: SignatureInfo.GetSignatureDetail Method (Office)
keywords: vbaof11.chm286006
f1_keywords:
- vbaof11.chm286006
ms.prod: office
api_name:
- Office.SignatureInfo.GetSignatureDetail
ms.assetid: 77a5a835-cc8a-0341-8e5d-6ddb603f9517
ms.date: 06/08/2017
---


# SignatureInfo.GetSignatureDetail Method (Office)

Displays a specified detail related to a signature.


## Syntax

 _expression_. **GetSignatureDetail**( **_sigdet_** )

 _expression_ An expression that returns a **SignatureInfo** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _sigdet_|Required|**SignatureDetail**|An enumerated value specifying which signature detail to display.|

### Return Value

Variant


## Example

The following example gets information on the suggested signer of the document.


```
Sub GetSigDetails() 
Dim objSignatureInfo As SignatureInfo 
Dim varDetail As Variant 
 
strDetail = objSignatureInfo.GetSignatureDetail(sigdetDelSuggSigner) 
 
End Sub
```


## See also


#### Concepts


[SignatureInfo Object](signatureinfo-object-office.md)
#### Other resources


[SignatureInfo Object Members](signatureinfo-members-office.md)

