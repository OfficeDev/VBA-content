---
title: Application.ProductCode Property (Publisher)
keywords: vbapb10.chm131105
f1_keywords:
- vbapb10.chm131105
ms.prod: publisher
api_name:
- Publisher.Application.ProductCode
ms.assetid: aacd5ff6-dad1-af86-f4e0-af9012ae93f8
ms.date: 06/08/2017
---


# Application.ProductCode Property (Publisher)

Returns a  **String** indicating the Microsoft Publisher globally unique identifier (GUID). Read-only.


## Syntax

 _expression_. **ProductCode**

 _expression_A variable that represents a  **Application** object.


### Return Value

String


## Example

The following example displays the product code for Publisher.


```vb
MsgBox "The product code for Microsoft Publisher is " _ 
 &; ProductCode
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

