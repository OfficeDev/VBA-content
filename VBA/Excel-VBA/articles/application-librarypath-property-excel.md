---
title: Application.LibraryPath Property (Excel)
keywords: vbaxl10.chm133155
f1_keywords:
- vbaxl10.chm133155
ms.prod: excel
api_name:
- Excel.Application.LibraryPath
ms.assetid: 783efa4a-640b-ab78-2831-da2ecd05558a
ms.date: 06/08/2017
---


# Application.LibraryPath Property (Excel)

Returns the path to the Library folder, but without the final separator. Read-only  **String** .


## Syntax

 _expression_ . **LibraryPath**

 _expression_ A variable that represents an **Application** object.


## Example

This example opens the file Oscar.xla in the Library folder.


```
pathSep = Application.PathSeparator 
f = Application.LibraryPath &; pathSep &; "Oscar.Xla" 
Workbooks.Open filename:=f
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

