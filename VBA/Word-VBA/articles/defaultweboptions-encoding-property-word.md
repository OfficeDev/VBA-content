---
title: DefaultWebOptions.Encoding Property (Word)
keywords: vbawd10.chm165871629
f1_keywords:
- vbawd10.chm165871629
ms.prod: word
api_name:
- Word.DefaultWebOptions.Encoding
ms.assetid: 2876e36d-927d-c9aa-6df4-9f2995a3a3d1
ms.date: 06/08/2017
---


# DefaultWebOptions.Encoding Property (Word)

Returns or sets the document encoding (code page or character set) to be used by the Web browser when you view the saved document. Read/write  **MsoEncoding** .


## Syntax

 _expression_ . **Encoding**

 _expression_ Required. A variable that represents a **[DefaultWebOptions](defaultweboptions-object-word.md)** collection.


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


[DefaultWebOptions Object](defaultweboptions-object-word.md)

