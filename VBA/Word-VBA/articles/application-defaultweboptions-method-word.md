---
title: Application.DefaultWebOptions Method (Word)
keywords: vbawd10.chm158335381
f1_keywords:
- vbawd10.chm158335381
ms.prod: word
api_name:
- Word.Application.DefaultWebOptions
ms.assetid: ee683d3c-b331-cccd-27ec-b3258b42961e
ms.date: 06/08/2017
---


# Application.DefaultWebOptions Method (Word)

Returns the  **[DefaultWebOptions](defaultweboptions-object-word.md)** object that contains global application-level attributes used by Microsoft Word whenever you save a document as a Web page or open a Web page.


## Syntax

 _expression_ . **DefaultWebOptions**

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Return Value

DefaultWebOptions


## Example

This example checks to see whether the default setting for document encoding is Western, and then it sets the string strDocEncoding accordingly.


```vb
Dim strDocEncoding As String 
 
If Application.DefaultWebOptions.Encoding _ 
 = msoEncodingWestern Then 
 strDocEncoding = "Western" 
Else 
 strDocEncoding = "Other" 
End If
```


## See also


#### Concepts


[Application Object](application-object-word.md)

