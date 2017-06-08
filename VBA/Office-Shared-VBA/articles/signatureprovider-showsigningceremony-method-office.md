---
title: SignatureProvider.ShowSigningCeremony Method (Office)
keywords: vbaof11.chm287003
f1_keywords:
- vbaof11.chm287003
ms.prod: office
api_name:
- Office.SignatureProvider.ShowSigningCeremony
ms.assetid: d098e755-2f64-4801-6b5c-ef36d721ee9c
ms.date: 06/08/2017
---


# SignatureProvider.ShowSigningCeremony Method (Office)

Provides a signature provider add-in the opportunity to display the  **Signature** dialog box to users, allowing them to specify their identity and then be authenticated.


## Syntax

 _expression_. **ShowSigningCeremony**( **_ParentWindow_**, **_psigsetup_**, **_psiginfo_** )

 _expression_ An expression that returns a **SignatureProvider** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ParentWindow_|Required|**IOleWindow**|Contains the handle to the window containing the  **Signature** dialog box.|
| _psigsetup_|Required|**SignatureSetup**|Specifies initial settings of the signature provider.|
| _psiginfo_|Required|**SignatureInfo**|Specifies information about the signature provider.|

## Remarks

This method is internally called by the Microsoft Office application when the user is attempting to sign a signature line, or if the add-in has called the  **Sign** method in the Office application's object model on a **SignatureLine** object.


## Example

The following example, written in C#, shows the implementation of the  **ShowSigningCeremony** method in a custom signature provider project.


```
 public void ShowSigningCeremony(object parentWindow, SignatureSetup sigsetup, SignatureInfo siginfo) 
 { 
 using (Win32WindowFromOleWindow window = new Win32WindowFromOleWindow(parentWindow)) 
 { 
 if (!((bool) siginfo.GetCertificateDetail(CertificateDetail.certdetAvailable))) 
 { 
 MessageBox.Show(window, "You need a digital certificate to sign this document", "Signing Ceremony", MessageBoxButtons.OK); 
 throw new System.Runtime.InteropServices.COMException("Canceled", -2147467260 /*E_ABORT*/); 
 } 
 
 using (SigningCeremonyForm signForm = new SigningCeremonyForm(sigsetup, siginfo)) 
 { 
 signForm.ShowDialog(window); 
 if (!signForm.success) 
 throw new System.Runtime.InteropServices.COMException("Cancelled", -2147467260 /*E_ABORT*/); 
 } 
 } 
 } 

```


 **Note**  Signature providers are implemented exclusively in custom COM add-ins and cannot be implemented in Microsoft Visual Basic for Applications (VBA). 


## See also


#### Concepts


[SignatureProvider Object](signatureprovider-object-office.md)
#### Other resources


[SignatureProvider Object Members](signatureprovider-members-office.md)

