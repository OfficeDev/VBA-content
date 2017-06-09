---
title: SignatureProvider.GenerateSignatureLineImage Method (Office)
keywords: vbaof11.chm287001
f1_keywords:
- vbaof11.chm287001
ms.prod: office
api_name:
- Office.SignatureProvider.GenerateSignatureLineImage
ms.assetid: 36430574-939a-4050-c7b1-ce738be334a2
ms.date: 06/08/2017
---


# SignatureProvider.GenerateSignatureLineImage Method (Office)

Gets a signature line image.


## Syntax

 _expression_. **GenerateSignatureLineImage**( **_siglnimg_**, **_psigsetup_**, **_psiginfo_** )

 _expression_ An expression that returns a **SignatureProvider** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _siglnimg_|Required|**SignatureLineImage**|Contains the name if the signature line graphic.|
| _psigsetup_|Required|**SignatureSetup**|Specifies initial settings of the signature provider add-in.|
| _psiginfo_|Required|**SignatureInfo**|Specifies information about the signature provider add-in.|

## Remarks

The  **SignatureProvider** object is used exclusively in custom signature provider add-ins. This method is called for each of the images that is displayed in the document's content. The method can be called asynchronously. For example, the method may be called for the "Unsigned" image and the "No-software" image directly after signature setup. The method may then be called after signing for the "Signed" image. The four images used are:


-  **siglnimgSoftwareRequired**: This image is displayed to the user if the signature provider add-in is not installed on the user's computer. If the user tries to sign or view a signature line, they are redirected to a provider-supplied hyperlink specified in the **GetProviderDetail** method.
    
-  **siglnimgUnsigned**: This image is displayed for an unsigned signature image. Basically, when a document loads with an unsigned signature line, the signature provider prompts for an up-to-date signature image and shows this image.
    
-  **siglnimgSignedValid**: This is the image that is displayed when a signature line is signed and valid (or, to be more specific, signed and the signature does not register as invalid). When the document opens, it is assumed that all signed signatures are valid until the verification process is complete, at which point a "Signed/invalid" image is displayed for the invalid signatures. Because signature verification is time-intensive, the signature verification runs in parallel with Office on a background thread. Because your add-in implements signature verification, your code runs parallel with Office and should not attempt to display UI during signing or verifying.
    
-  **siglnimgSignedInvalid**: This is the image we show when a signature line is signed but there is a problem with the signature, such as the document was modified or the user's certificate is revoked. Because your add-in implements signature verification, you can decide how and when a signature is invalid.
    



## Example

The following example, written in C#, shows the implementation of the  **GenerateSignatureLineImage** method in a custom signature provider project.


```
 public IPictureDisp GenerateSignatureLineImage(SignatureLineImage siglnimg, SignatureSetup sigsetup, SignatureInfo siginfo, object xmldsigStream) 
 { 
 IPictureDisp sigline = null; 
 
 System.Drawing.Bitmap draw = new System.Drawing.Bitmap(200, 100); 
 Graphics g = Graphics.FromImage(draw); 
 g.DrawRectangle(new Pen(Color.Gray, 2), 0, 0, 200, 100); 
 
 if (siglnimg == SignatureLineImage.siglnimgUnsigned) 
 { 
 g.FillRectangle(new SolidBrush(Color.LightSlateGray), 2, 2, 196, 96); 
 g.DrawString("Requested Signature", new System.Drawing.Font("Verdana", 10), new SolidBrush(Color.Yellow), new PointF(20, 20)); 
 g.DrawString(sigsetup.SuggestedSigner, new System.Drawing.Font("Courier", 8), new SolidBrush(Color.Yellow), new PointF(20, 50)); 
 } 
 else if (siglnimg == SignatureLineImage.siglnimgSignedValid) 
 { 
 g.FillRectangle(new SolidBrush(Color.LightSlateGray), 2, 2, 196, 96); 
 g.DrawString("Valid Signature", new System.Drawing.Font("Verdana", 10), new SolidBrush(Color.LimeGreen), new PointF(20, 20)); 
 g.DrawString(sigsetup.SuggestedSigner, new System.Drawing.Font("Courier", 8), new SolidBrush(Color.LimeGreen), new PointF(20, 50)); 
 } 
 else if (siglnimg == SignatureLineImage.siglnimgSignedInvalid) 
 { 
 g.FillRectangle(new SolidBrush(Color.LightSlateGray), 2, 2, 196, 96); 
 g.DrawString("Invalid Signature", new System.Drawing.Font("Verdana", 10), new SolidBrush(Color.Red), new PointF(20, 20)); 
 g.DrawString(sigsetup.SuggestedSigner, new System.Drawing.Font("Courier", 8), new SolidBrush(Color.Red), new PointF(20, 50)); 
 } 
 else if (siglnimg == SignatureLineImage.siglnimgSoftwareRequired) 
 { 
 g.FillRectangle(new SolidBrush(Color.LightSlateGray), 2, 2, 196, 96); 
 g.DrawString("Software Required", new System.Drawing.Font("Verdana", 10), new SolidBrush(Color.AliceBlue), new PointF(20, 20)); 
 } 
 else 
 { 
 throw new NotImplementedException(); 
 } 
 
 System.IntPtr hbitmap = draw.GetHbitmap(Color.Green); 
 Image img = Image.FromHbitmap(hbitmap); 
 
 sigline = (IPictureDisp)AxHost2.GetIPictureDispFromPicture(img); 
 
 return sigline; 
 
 }
```


 **Note**  Signature providers are implemented exclusively in custom COM add-ins and cannot be implemented in Microsoft Visual Basic for Applications (VBA). 


## See also


#### Concepts


[SignatureProvider Object](signatureprovider-object-office.md)
#### Other resources


[SignatureProvider Object Members](signatureprovider-members-office.md)

