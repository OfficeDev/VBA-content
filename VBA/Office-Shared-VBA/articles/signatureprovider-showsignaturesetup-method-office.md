---
title: SignatureProvider.ShowSignatureSetup Method (Office)
keywords: vbaof11.chm287002
f1_keywords:
- vbaof11.chm287002
ms.prod: office
api_name:
- Office.SignatureProvider.ShowSignatureSetup
ms.assetid: 458efe65-acb8-f329-7ca4-b0a316869c13
ms.date: 06/08/2017
---


# SignatureProvider.ShowSignatureSetup Method (Office)

Provides a signature provider add-in the opportunity to display the  **Signature Setup** dialog box to the user.


## Syntax

 _expression_. **ShowSignatureSetup**( **_ParentWindow_**, **_psigsetup_** )

 _expression_ An expression that returns a **SignatureProvider** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ParentWindow_|Required|**IOleWindow**|Contains the handle to the window containing the  **Signature Setup** dialog box.|
| _psigsetup_|Required|**SignatureSetup**|Specifies initial settings of the signature provider.|

## Remarks

This method is used for both the insertion time configuration process and if a user later wants to re-configure the signature line. You display the  **Signature Setup** dialog box during this callback and wait for the user to select **OK** or **Cancel**. It is not necessary to display a dialog box for signature setup unless you specifically need information from the author about the signature line. If you can provide all of the necessary details back to Microsoft Office without user input, then no dialog is necessary.


## Example

The following example, written in C#, shows the implementation of the  **ShowSignatureSetup** method in a custom signature provider project.


```
 public void ShowSignatureSetup(object parentWindow, SignatureSetup sigsetup) 
 { 
 bool firstInit = string.IsNullOrEmpty(sigsetup.AdditionalXml); 
 if (sigsetup != null &amp;&amp; !sigsetup.ReadOnly &amp;&amp; firstInit) 
 { 
 sigsetup.SigningInstructions = "Please sign this document."; 
 sigsetup.ShowSignDate = true; 
 sigsetup.AdditionalXml = "<TestSignatureData />"; 
 } 
 
 using (Win32WindowFromOleWindow window = new Win32WindowFromOleWindow(parentWindow)) 
 { 
 using (SignatureSetupForm sigsetupForm = new SignatureSetupForm(sigsetup)) 
 { 
 sigsetupForm.ShowDialog(window); 
 if (!sigsetupForm.success &amp;&amp; firstInit) 
 throw new System.Runtime.InteropServices.COMException("Canceled", -2147467260 /*E_ABORT*/); 
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

