---
title: Signature.ShowDetails Method (Office)
keywords: vbaof11.chm248014
f1_keywords:
- vbaof11.chm248014
ms.prod: office
api_name:
- Office.Signature.ShowDetails
ms.assetid: 278b84b3-c500-6357-310b-537355ad20fd
ms.date: 06/08/2017
---


# Signature.ShowDetails Method (Office)

Displays details related to a signature packet.


## Syntax

 _expression_. **ShowDetails**

 _expression_ An expression that returns a **Signature** object.


## Example

The following example calls the  **ShowDetails** method to show details of the **Signature** object.


```
Sub getSignatureDetails(ByVal objSignature As Signature) 
If objSignature.IsSigned then 
 Msgbox(The document has been signed with the following details: " &amp; objSignature.ShowDetails) 
Else 
 Msgbox("The document has not been signed.") 
End If 
End Sub 
```


## See also


#### Concepts


[Signature Object](signature-object-office.md)
#### Other resources


[Signature Object Members](signature-members-office.md)

