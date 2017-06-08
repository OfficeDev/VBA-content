---
title: Presentation.Signatures Property (PowerPoint)
keywords: vbapp10.chm583067
f1_keywords:
- vbapp10.chm583067
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.Signatures
ms.assetid: 978e39bb-298b-d820-63cb-2924bf0770b1
ms.date: 06/08/2017
---


# Presentation.Signatures Property (PowerPoint)

Returns a  **SignatureSet** object that represents a collection of digital signatures. Read-only.


## Syntax

 _expression_. **Signatures**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

SignatureSet


## Example

The following line of code displays the number of digital signatures.


```vb
Sub DisplayNumberOfSignatures
    MsgBox "Number of digital signatures: " &; _
        ActivePresentation.Signatures.Count
End Sub
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

