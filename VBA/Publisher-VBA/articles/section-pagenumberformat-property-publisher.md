---
title: Section.PageNumberFormat Property (Publisher)
keywords: vbapb10.chm7405573
f1_keywords:
- vbapb10.chm7405573
ms.prod: publisher
api_name:
- Publisher.Section.PageNumberFormat
ms.assetid: 5b64a352-2fd8-9e19-3425-a7984dd67edd
ms.date: 06/08/2017
---


# Section.PageNumberFormat Property (Publisher)

Sets or returns a  **PbPageNumberFormat** constant that reperesents the formatting of the page numbering. Read/write.


## Syntax

 _expression_. **PageNumberFormat**

 _expression_A variable that represents a  **Section** object.


### Return Value

PbPageNumberFormat


## Remarks

The  **PageNumberFormat** property value can be one of the **[PbPageNumberFormat](pbpagenumberformat-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.

Not all of the  **PbPageNumberFormat** constants are available, depending on the languages that are enabled or installed.


## Example

This example adds a new section to the active document, sets the page number format to lowercase roman, and then sets the starting page number to 1.


```vb
Dim objSection As Section 
Set objSection = ActiveDocument.Sections.Add(2) 
With objSection 
 .PageNumberFormat = pbPageNumberFormatLCRoman 
 .PageNumberStart = 1 
End With 

```


