---
title: Hyperlink.CreateURL Method (Visio)
keywords: vis_sdr.chm15016155
f1_keywords:
- vis_sdr.chm15016155
ms.prod: visio
api_name:
- Visio.Hyperlink.CreateURL
ms.assetid: 3a9cdcb3-19cd-fe03-51a7-24b916b870cc
ms.date: 06/08/2017
---


# Hyperlink.CreateURL Method (Visio)

Returns a fully qualified and optionally canonicalized representation of the hyperlink's absolute address.


## Syntax

 _expression_ . **CreateURL**( **_CanonicalForm_** )

 _expression_ A variable that represents a **Hyperlink** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CanonicalForm_|Required| **Integer**| **True** (non-zero) if canonical form; otherwise, **False** (0).|

### Return Value

String


## Remarks

The  **CreateURL** method of the **Hyperlink** object can be used to resolve relative URLs against a hyperlink's base address.

When you use the canonical form, the  **CreateURL** method applies URL canonicalization rules to the hyperlink. Only spaces are URL-encoded during canonicalization. Port 80 is assumed for HTTP URLs and is removed during canonicalization. For example, the URL "http://www.microsoft.com:80/" is returned as "http://www.microsoft.com/", whereas http://www.microsoft.com:1000/" is unchanged.


## Example

Here are some examples of results of the  **CreateURL** method:


```
Address = "http://www.microsoft.com/" 
CreateURL(False) returns "http://www.microsoft.com/" 
 
Address = "C:\My Documents\Spreadsheet.XLS" 
CreateURL(False) returns "file://C:\My Documents\Spreadsheet.XLS" 
CreateURL(True) returns "file://C:\My%20Documents\Spreadsheet.XLS" 
 

```

Relative path example:




```
Assume : Document.HyperlinkBase = "http://www.microsoft.com/widgets/" 
Address = "../file.htm" 
CreateURL(False) returns "http://www.microsoft.com/file.htm" 
 

```



The following example shows how to use the  **CreateURL** method to resolve relative URLs against the base address of a hyperlink. Before running this macro, replace _drive\folder\subfolder_ with a valid file path on your computer, replace _address_ with a valid Internet or intranet address, and replace _drawing.vsd_ with a valid file on your computer.




```vb
 
Sub CreateURL_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim vsoHyperlink As Visio.Hyperlink 
 
 'Draw a rectangle shape on the active page 
 Set vsoShape = ActivePage.DrawRectangle(1, 2, 2, 1) 
 
 'Add a hyperlink to a shape 
 Set vsoHyperlink = vsoShape.AddHyperlink 
 
 'Allow relative hyperlink addresses 
 ActiveDocument.HyperlinkBase = "drive :\folder \subfolder " 
 
 'Return a relative address 
 vsoHyperlink.Address = "..\drawing.vsd " 
 
 'Print the resulting URLs to the Debug window 
 'to show how the relative path is derived 
 'from the base path and the difference 
 'between canonical and non-canonical forms 
 Debug.Print vsoHyperlink.CreateURL(False) 
 Debug.Print vsoHyperlink.CreateURL(True) 
 
 'Return an absolute address 
 vsoHyperlink.Address = "http://address " 
 
 'Print the resulting URL to the Debug window 
 Debug.Print vsoHyperlink.CreateURL(False) 
 
End Sub
```


