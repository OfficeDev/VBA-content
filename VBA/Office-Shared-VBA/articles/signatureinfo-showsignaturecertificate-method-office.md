---
title: SignatureInfo.ShowSignatureCertificate Method (Office)
keywords: vbaof11.chm286014
f1_keywords:
- vbaof11.chm286014
ms.prod: office
api_name:
- Office.SignatureInfo.ShowSignatureCertificate
ms.assetid: 8fef7299-e110-b0a2-7a0c-552e9068e001
ms.date: 06/08/2017
---


# SignatureInfo.ShowSignatureCertificate Method (Office)

Displays the selected or default digital certificate. 


## Syntax

 _expression_. **ShowSignatureCertificate**( **_ParentWindow_** )

 _expression_ An expression that returns a **SignatureInfo** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ParentWindow_|Required|**IOleWindow**|Contains the handle to the window that contains the  **Certificate** dialog box.|

## Example

The following example displays a digital certificate in the window specified by the  _Hwnd_ argument.


```
Sub DisplayCertificate(ByVal intHwnd As Long) 
Dim objSignatureInfo As SignatureInfo 
Dim objDialog As Object 
 
objDialog = objSignatureInfo.ShowSignatureCertificate(intHwnd) 
 
End Sub
```


## See also


#### Concepts


[SignatureInfo Object](signatureinfo-object-office.md)
#### Other resources


[SignatureInfo Object Members](signatureinfo-members-office.md)

