---
title: SignatureSet.AddSignatureLine Method (Office)
keywords: vbaof11.chm247008
f1_keywords:
- vbaof11.chm247008
ms.prod: office
api_name:
- Office.SignatureSet.AddSignatureLine
ms.assetid: e887431f-8a01-99d7-6c9b-21aaf3d9198d
ms.date: 06/08/2017
---


# SignatureSet.AddSignatureLine Method (Office)

Adds lines to a document where signatures are collected.


## Syntax

 _expression_. **AddSignatureLine**( **_varSigProv_** )

 _expression_ An expression that returns a **SignatureSet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _varSigProv_|Optional|**Variant**|Represents the ID of the signature provider.|

### Return Value

Signature


## Remarks

After the line is added, the author of the document can add the necessary information so that each signature line shows the name and (optionally) the title of the person who should sign. When a user opens the document, Microsoft Office recognizes that one or more signature lines are present, but blank. Office then alerts the user that they need to sign this document and helps them find where in the document the requested signatures are located.


## Example

The procedure in the following example receives the ID of a signature provider and, as long as the document is not read-only, adds a signature line.


```
Function InsertSignatureLines(ByVal SignProviderID As Variant) As Signature 
Dim objSignature As Signature 
 
If CanAddSignatureLine Then 
 objSignature = AddSignatureLine(SignProviderID) 
End If 
 
InsertSignatureLines = objSignature 
 
End Function
```


## See also


#### Concepts


[SignatureSet Object](signatureset-object-office.md)
#### Other resources


[SignatureSet Object Members](signatureset-members-office.md)

