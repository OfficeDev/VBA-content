---
title: SignatureSetup Object (Office)
keywords: vbaof11.chm285000
f1_keywords:
- vbaof11.chm285000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SignatureSetup
ms.assetid: e76b87c9-3163-654c-ab52-559dfdf43c90
---


# SignatureSetup Object (Office)

Represents the information used to set up a signature packet.


## Example

The following example sets various properties of the  **SignatureSetup** object for a signature packet.


```vb
Dim objSigSetup As SignatureSetup 
With objSigSetup 
.AllowComments = True 
.ShowSignDate = True 
.SigningInstructions = "Please sign this document." 
.SuggestedSignerEmail = "jdow@example.com" 
Next
```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

