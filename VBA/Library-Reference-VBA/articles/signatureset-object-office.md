---
title: SignatureSet Object (Office)
keywords: vbaof11.chm247000
f1_keywords:
- vbaof11.chm247000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SignatureSet
ms.assetid: 574cba16-c632-ab66-f014-58172ff1c091
---


# SignatureSet Object (Office)

A collection of  **Signature** objects that correspond to the digital signature attached to a document.


## Remarks

Use the  **Signatures** property of the **Document** object to return a **SignatureSet** collection; for example:


```
Set sigs = ActiveDocument.Signatures
```

You can add a  **Signature** object to a **SignatureSet** collection using the **Add** method and you can return an existing member using the **Item** method. The **AddSignatureLine** method also adds a **Signature** object to the collection. Also see the **Subset** property, which acts as a filter for whether certain **Signature** objects appear in the collection. To remove a **Signature** from a **SignatureSet** collection, use the **Delete** method of the **Signature** object.


## Example

The following example prompts the user to select a digital signature with which to sign the active document in Microsoft Word. To use this example, open a document in Word and pass this function the name of a certificate issuer and the name of a certificate signer that match the  **Issued By** and **Issued To** fields of a digital certificate in the **Digital Certificates** dialog box. This example will test to make sure that the digital signature that the user selects meets certain criteria, such as not having expired, before the new signature is committed to the disk.


```
Function AddSignature(ByVal strIssuer As String, _ 
 strSigner As String) As Boolean 
 
 Dim sig As Signature 
 
 'Display the dialog box that lets the 
 'user select a digital signature. 
 'If the user selects a signature, then 
 'it is added to the Signatures 
 'collection. If the user doesn't, then 
 'an error is returned. 
 Set sig = ActiveDocument.Signatures.Add 
 
 'Test several properties before committing the Signature object to disk. 
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
 
End Function
```


## Methods



|**Name**|
|:-----|
|[AddNonVisibleSignature](http://msdn.microsoft.com/library/f8d3a749-9507-628f-2192-552bd4cbb00c%28Office.15%29.aspx)|
|[AddSignatureLine](http://msdn.microsoft.com/library/e887431f-8a01-99d7-6c9b-21aaf3d9198d%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/55eb69e8-f7d0-ed4c-ef9f-91e374b4f658%28Office.15%29.aspx)|
|[CanAddSignatureLine](http://msdn.microsoft.com/library/e5b54883-4ac5-b239-b17c-efbdcd4bc849%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/6918bb9c-775e-241d-c126-6e4a3a63c654%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/0fc2f22f-57b8-0dc9-1e31-48b5a66b01bf%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/f55249e6-22e1-84bd-175f-e615533a37cd%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/88fd9392-e2f3-e84e-9f7c-c2fce32de296%28Office.15%29.aspx)|
|[ShowSignaturesPane](http://msdn.microsoft.com/library/1aa332cd-5b4e-06e8-2ebb-3c64128ded04%28Office.15%29.aspx)|
|[Subset](http://msdn.microsoft.com/library/0ce176cb-9869-19ed-a3bc-e17b04c59255%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
[SignatureSet Object Members](http://msdn.microsoft.com/library/abe810a3-ffe4-ee26-8df7-d68cfbf3bf1e%28Office.15%29.aspx)
