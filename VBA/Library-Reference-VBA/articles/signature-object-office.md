---
title: Signature Object (Office)
keywords: vbaof11.chm248000
f1_keywords:
- vbaof11.chm248000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.Signature
ms.assetid: 574d246b-95cd-e4da-081b-4540387662a0
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


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/c75a2200-081a-7f5c-ae02-ab7be845c003%28Office.15%29.aspx)|
|[ShowDetails](http://msdn.microsoft.com/library/278b84b3-c500-6357-310b-537355ad20fd%28Office.15%29.aspx)|
|[Sign](http://msdn.microsoft.com/library/37ba202a-da6d-9978-c8af-986a8218e004%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/fc445340-37a7-f9df-49a6-1489ac49b5f6%28Office.15%29.aspx)|
|[CanSetup](http://msdn.microsoft.com/library/6c4903e9-2fd0-3947-aeb1-c0bc9c437fe7%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/f0b89afe-1aee-d0bb-8756-40396e662b6d%28Office.15%29.aspx)|
|[Details](http://msdn.microsoft.com/library/c5de710a-876f-8eb4-ec46-21359b8d4bf4%28Office.15%29.aspx)|
|[IsSignatureLine](http://msdn.microsoft.com/library/88ed582d-ee3c-7aaa-cb46-90098f6968a9%28Office.15%29.aspx)|
|[IsSigned](http://msdn.microsoft.com/library/ddaa2ad6-26ce-35d7-ed69-9faef04b7a31%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/0b30078a-8f03-48b6-2b72-b3f2ddfaa76f%28Office.15%29.aspx)|
|[Setup](http://msdn.microsoft.com/library/9ccfd72f-af1c-a0d5-3a8f-97ee58bda211%28Office.15%29.aspx)|
|[SignatureLineShape](http://msdn.microsoft.com/library/8ba372b9-40f9-bc9c-03de-97827b0c257d%28Office.15%29.aspx)|
|[SortHint](http://msdn.microsoft.com/library/9554cf10-85ab-508c-a13e-08b9504bdd1a%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
[Signature Object Members](http://msdn.microsoft.com/library/1054db23-fe1c-f81f-e44b-d8c2c82ca7fa%28Office.15%29.aspx)
