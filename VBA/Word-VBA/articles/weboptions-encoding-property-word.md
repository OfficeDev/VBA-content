---
title: WebOptions.Encoding Property (Word)
keywords: vbawd10.chm165937162
f1_keywords:
- vbawd10.chm165937162
ms.prod: word
api_name:
- Word.WebOptions.Encoding
ms.assetid: 4156a3cc-744f-5a62-5961-a26e0e155567
ms.date: 06/08/2017
---


# WebOptions.Encoding Property (Word)

Returns or sets the document encoding (code page or character set) to be used by the Web browser when you view the saved document. Read/write  **MsoEncoding** .


## Syntax

 _expression_ . **Encoding**

 _expression_ Required. A variable that represents a **[WebOptions](weboptions-object-word.md)** collection.


## Example

This example checks to see whether the default document encoding is Western, and then it sets the string strDocEncoding accordingly.


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


[WebOptions Object](weboptions-object-word.md)

