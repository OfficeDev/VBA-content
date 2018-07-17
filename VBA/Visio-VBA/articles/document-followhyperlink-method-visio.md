---
title: Document.FollowHyperlink Method (Visio)
keywords: vis_sdr.chm10516295
f1_keywords:
- vis_sdr.chm10516295
ms.prod: visio
api_name:
- Visio.Document.FollowHyperlink
ms.assetid: 555e607d-7e4a-d3c8-9c78-1733b112200c
ms.date: 06/08/2017
---


# Document.FollowHyperlink Method (Visio)

Navigates to an arbitrary hyperlink.


## Syntax

 _expression_ . **FollowHyperlink**( **_Address_** , **_SubAddress_** , **_ExtraInfo_** , **_Frame_** , **_NewWindow_** , **_res1_** , **_res2_** , **_res3_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Address_|Required| **String**|The address to which you want to navigate. If you pass an incorrect or non-existent path or filename for  _Address_, Visio displays an error message.|
| _SubAddress_|Required| **String**|The secondary address to which you want to navigate; if you don't need this information, pass an empty string. If  _Address_ is the full path of a Visio document that contains multiple pages, for example, you can use _SubAddress_ to specify the page.|
| _ExtraInfo_|Optional| **Variant**|Extra URL request information to use in resolving the URL.|
| _Frame_|Optional| **Variant**|The HTML frame to which to navigate.|
| _NewWindow_|Optional| **Variant**|Specifies if a new window is to be opened. Passing any non-zero number or  **True** opens the linked page in a new window.|
| _res1_|Optional| **Variant**|Unused.|
| _res2_|Optional| **Variant**|Unused.|
| _res3_|Optional| **Variant**|Unused.|

### Return Value

Nothing


## Remarks

If you don't need to pass any information for one or more optional arguments, from Microsoft Visual Basic or Visual Basic for Applications, do not pass a value. From C or C++, pass an empty variant.

Visio 4.5 provided an undocumented  **Hyperlink** method for a **Document** object that had the following signature:




```
HRESULT FollowHyperlink[in] (BSTR Target, [in] BSTR Location);
```

Visio 5.0 and later still support this method but it has been renamed  **FollowHyperlink45** :




```
HRESULT FollowHyperlink45[in] (BSTR Target, [in] BSTR Location);
```

Use of  **FollowHyperlink45** is discouraged, however?unless you are using version 4.5, use **FollowHyperlink** instead.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **FollowHyperlink** method to navigate to a site on the Internet and view the resulting Web page in a new browser window. It also shows how to navigate to the second page of the current document and to the first page of another document on your computer. Before running the macro, substitute the path and file name of a valid Visio document on your computer for _&lt;path\filename&gt;_. 


```vb
Public Sub FollowHyperlink_Example() 
 
 'Navigate to the Microsoft Web site and view the page in a new browser window. 
 ThisDocument.FollowHyperlink "http://www.microsoft.com", "", , , 1 
 
 'Navigate to the second page of the current document. 
 ThisDocument.FollowHyperlink "", "Page-2" 
 
 'Navigate to the first page of another document 
 ThisDocument.FollowHyperlink "<path\filename> ", "Page-1" 
 
End Sub
```


