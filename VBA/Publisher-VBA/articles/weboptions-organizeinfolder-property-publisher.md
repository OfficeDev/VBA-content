---
title: WebOptions.OrganizeInFolder Property (Publisher)
keywords: vbapb10.chm8257542
f1_keywords:
- vbapb10.chm8257542
ms.prod: publisher
api_name:
- Publisher.WebOptions.OrganizeInFolder
ms.assetid: f09ac701-d8d8-a58f-965c-bd5e4b69820c
ms.date: 06/08/2017
---


# WebOptions.OrganizeInFolder Property (Publisher)

Returns or sets a  **Boolean** value that specifies whether a Web publication will be saved in a flat structure or hierarchical structure. If **False**, all files in the Web publication will be saved in a flat structure within the root folder. If  **True**, the files will be saved in a hierarchical structure within the root folder. The default value is  **True**. Read/write.


## Syntax

 _expression_. **OrganizeInFolder**

 _expression_A variable that represents an  **WebOptions** object.


### Return Value

Boolean


## Example

The following example specifies that all files in the Web publication should be saved in a flat structure within the root folder.


```vb
Dim theWO As WebOptions 
 
Set theWO = Application.WebOptions 
 
With theWO 
 .OrganizeInFolder = False 
End With
```


