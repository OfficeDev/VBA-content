---
title: Document.SolutionXMLElementExists Property (Visio)
keywords: vis_sdr.chm10550865
f1_keywords:
- vis_sdr.chm10550865
ms.prod: visio
api_name:
- Visio.Document.SolutionXMLElementExists
ms.assetid: d4a0bd9b-a3ea-de0a-5c33-ccad4d4398eb
ms.date: 06/08/2017
---


# Document.SolutionXMLElementExists Property (Visio)

Indicates whether a named SolutionXML element exists in the document. Read-only.


## Syntax

 _expression_ . **SolutionXMLElementExists**( **_ElementName_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ElementName_|Required| **String**|The case-sensitive name of the SolutionXML element.|

### Return Value

Boolean


## Remarks

Because the  **SolutionXMLElement** property can overwrite existing XML data, always use the **SolutionXMLElementExists** property to verify whether _ElementName_ already exists in the document.


