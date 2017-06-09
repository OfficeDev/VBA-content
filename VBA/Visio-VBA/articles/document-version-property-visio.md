---
title: Document.Version Property (Visio)
keywords: vis_sdr.chm10551135
f1_keywords:
- vis_sdr.chm10551135
ms.prod: visio
api_name:
- Visio.Document.Version
ms.assetid: 336b6825-3d1c-9589-e916-f8d7821f6383
ms.date: 06/08/2017
---


# Document.Version Property (Visio)

Gets the version of a saved document or sets the version in which to save a document. Read/write.


## Syntax

 _expression_ . **Version**

 _expression_ A variable that represents a **Document** object.


### Return Value

VisDocVersions


## Remarks

Setting the  **Version** property of a document tells Microsoft Visio which file format version to save the document in the next time the document is saved. The Visio type library declares constants for file format versions in **[VisDocVersions](visdocversions-enumeration-visio.md)** .

Microsoft Visio can save a document in the following file format versions. Note that there are two constants,  **visVersion120** and **visVersion110** , that have identical values.



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visVersion60**|&;H60000|Visio 2000 or Visio 2002 document|
| **visVersion110**|&;HB0000| Visio 2003 or Visio 2007 document|
| **visVersion120**|&;HB0000|Visio 2003 or Visio 2007 document|
When Visio opens a document that was saved in an earlier version format, it converts the document's in-memory representation to the current version. However, when closing the document, Visio recognizes that the document was saved in an earlier version format and allows the user to choose the version in which to save the document.


