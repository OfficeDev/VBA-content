---
title: Section.ContinueNumbersFromPreviousSection Property (Publisher)
keywords: vbapb10.chm7405575
f1_keywords:
- vbapb10.chm7405575
ms.prod: publisher
api_name:
- Publisher.Section.ContinueNumbersFromPreviousSection
ms.assetid: a3d64f14-dc65-4fb1-5079-0fdf2e3f8f38
ms.date: 06/08/2017
---


# Section.ContinueNumbersFromPreviousSection Property (Publisher)

 **True** if the specified section continues the numbering from the prvious section. Read/write **Boolean**.


## Syntax

 _expression_. **ContinueNumbersFromPreviousSection**

 _expression_A variable that represents a  **Section** object.


### Return Value

Boolean


## Example

The following example adds three pages to the publication, adds a new section after the first page, and then sets the  **ContinueNumbersFromPreviousSection** to **False** for the new section.


```vb
Dim objSection As Section 
ActiveDocument.Pages.Add Count:=3, After:=1 
Set objSection = ActiveDocument.Sections.Add(StartPageIndex:=2) 
objSection.ContinueNumbersFromPreviousSection = False 
 
 

```


