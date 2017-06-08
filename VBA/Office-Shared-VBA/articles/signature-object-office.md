---
title: Signature Object (Office)
keywords: vbaof11.chm248000
f1_keywords:
- vbaof11.chm248000
ms.prod: office
api_name:
- Office.Signature
ms.assetid: 574d246b-95cd-e4da-081b-4540387662a0
ms.date: 06/08/2017
---


# Signature Object (Office)

Represents a digital signature attached to a document.  **Signature** objects are contained in the **SignatureSet** collection of the **Document** object.


## Remarks

You can add a  **Signature** object to a **SignatureSet** collection using the **Add** method and you can return an existing member using the **Item** method. To remove a **Signature** from a **SignatureSet** collection, use the **Delete** method of the **Signature** object.


## Example

The following example prompts the user to select a digital signature with which to sign the active document in Microsoft Word. To use this example, open a document in Word and pass this function the name of a certificate issuer and the name of a certificate signer that match the  **Issued By** and **Issued To** fields of a digital certificate in the **Digital Certificates** dialog box. This example will test to make sure that the digital signature that the user selects meets certain criteria, such as not having expired, before the new signature is committed to the disk.


```
Function AddSignature(ByVal strIssuer As String, _ 
 strSigner As String) As Boolean 
 
 On Error GoTo Error_Handler 
 
 Dim sig As Signature 
 
 'Display the dialog box that lets the 
 'user select a digital signature. 
 'If the user selects a signature, then 
 'it is added to the Signatures 
 'collection. If the user does not, then 
 'an error is returned. 
 Set sig = ActiveDocument.Signatures.Add 
 
 'Test several properties before commiting the Signature object to disk. 
 If sig.Issuer = strIssuer And _ 
 sig.Signer = strSigner And _ 
 sig.IsCertificateExpired = False And _ 
 sig.IsCertificateRevoked = False And _ 
 sig.IsValid = True Then 
 
 MsgBox "Signed" 
 AddSignature = True 
 'Otherwise, remove the Signature object from the SignatureSet collection. 
 Else 
 sig.Delete 
 MsgBox "Not signed" 
 AddSignature = False 
 End If 
 
 'Commit all signatures in the SignatureSet collection to the disk. 
 ActiveDocument.Signatures.Commit 
 
 Exit Function 
Error_Handler: 
 AddSignature = False 
 MsgBox "Action canceled." 
End Function
```


## See also


#### Concepts


[Object Model Reference](reference-object-library-reference-for-office.md)
#### Other resources


[Signature Object Members](signature-members-office.md)

