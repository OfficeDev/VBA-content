---
title: Document.PublicationType Property (Publisher)
keywords: vbapb10.chm196736
f1_keywords:
- vbapb10.chm196736
ms.prod: publisher
api_name:
- Publisher.Document.PublicationType
ms.assetid: 264c2769-2452-0009-4853-84a6a426db38
ms.date: 06/08/2017
---


# Document.PublicationType Property (Publisher)

Returns a  **PbPublicationType** constant that represents the type of the specified publication. Read-only.


## Syntax

 _expression_. **PublicationType**

 _expression_A variable that represents a  **Document** object.


### Return Value

PbPublicationType


## Remarks

The  **PublicationType** property value can be one of these **PbPublicationType** constants.



| **pbTypePrint**|
| **pbTypeWeb**|

## Example

The following example determines if the active publication is a print publication. If it is, the publication is converted to a Web publication.


```vb
Sub ChangePublicationType() 
 With ActiveDocument 
 If .PublicationType = pbTypePrint Then 
 .ConvertPublicationType (pbTypeWeb) 
 End If 
 End With 
End Sub
```


