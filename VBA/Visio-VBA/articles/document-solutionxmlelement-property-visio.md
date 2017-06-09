---
title: Document.SolutionXMLElement Property (Visio)
keywords: vis_sdr.chm10550855
f1_keywords:
- vis_sdr.chm10550855
ms.prod: visio
api_name:
- Visio.Document.SolutionXMLElement
ms.assetid: 44e9daa6-96dc-3041-ed50-dd4670298b6a
ms.date: 06/08/2017
---


# Document.SolutionXMLElement Property (Visio)

Contains solution-specific, well-formed XML data stored with a document. Read/write.


## Syntax

 _expression_ . **SolutionXMLElement**( **_ElementName_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ElementName_|Required| **String**|The case-sensitive name of the SolutionXML data element.|

### Return Value

String


## Remarks

The value of  _ElementName_ must match the value of the SolutionXML element's Name attribute. For example, if a solution's XML data began with the statement <SolutionXML Name='somename'>, use the _ElementName_ "somename" to retrieve that data.




- If  _ElementName_ already exists, the **SolutionXMLElement** property overwrites existing XML data. Use the **SolutionXMLElementExists** property before writing XML data to avoid losing data unintentionally.
    
- If  _ElementName_ does not exist, the **SolutionXMLElement** property creates an element by that name.
    


Because your XML data is validated when you write it, you will typically perform this operation during a  **DocumentSaved** event for performance reasons.

At the document level, if the XML data you pass to the  **SolutionXMLElement** property is well formed and contains a valid schema and namespace declaration, it is saved as nested XML within the Microsoft Visio VDX file format. If you pass invalid XML data, Visio converts this data to an XML comment so that the data is not lost. However, if you subsequently load the saved VDX file containing the comment into Visio, the XML comment will be ignored, and consequently the data will be lost.

If you put invalid or non-well-formed XML data into a cell, Visio saves it as a string in the cell so that it is not lost and can perhaps later be fixed.


