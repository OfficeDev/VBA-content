---
title: Application.ProductCode Property (PowerPoint)
keywords: vbapp10.chm502037
f1_keywords:
- vbapp10.chm502037
ms.prod: powerpoint
api_name:
- PowerPoint.Application.ProductCode
ms.assetid: 27376e9f-47c6-7373-af34-4ce71723e6a6
ms.date: 06/08/2017
---


# Application.ProductCode Property (PowerPoint)

Returns the Microsoft PowerPoint globally unique identifier (GUID). Read-only.


## Syntax

 _expression_. **ProductCode**

 _expression_ A variable that represents a **Application** object.


### Return Value

String


## Remarks

You might use the GUID, for example, when making program calls to an Application Programming Interface (API). 


## Example

This example returns the PowerPoint GUID to the variable  `pptGUID`.


```vb
Dim pptGUID As String

pptGUID = Application.ProductCode
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

