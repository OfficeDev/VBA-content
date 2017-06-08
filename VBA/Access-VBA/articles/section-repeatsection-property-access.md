---
title: Section.RepeatSection Property (Access)
keywords: vbaac10.chm12199
f1_keywords:
- vbaac10.chm12199
ms.prod: access
api_name:
- Access.Section.RepeatSection
ms.assetid: 8995af8f-f3c2-456c-dbd8-721e37ced40f
ms.date: 06/08/2017
---


# Section.RepeatSection Property (Access)

You can use the  **RepeatSection** property to specify whether a group header is repeated on the next page or column when a group spans more than one page or column. Read/write **Boolean**.


## Syntax

 _expression_. **RepeatSection**

 _expression_ A variable that represents a **Section** object.


## Remarks

The  **RepeatSection** property only applies to group headers on a report.

The  **RepeatSection** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|**Yes**|**True**|The group header is repeated.|
|**No**|**False**|(Default) The group header isn't repeated.|
When printing a report that contains a subreport, the subreport's  **RepeatSection** property will determine if the subreport group headers are repeated across pages or columns.


## Example

The following example prints the group header "GroupHeader0" at the top of each page.


```vb
Reports("Purchase Order").Section("GroupHeader0").RepeatSection = True
```


## See also


#### Concepts


[Section Object](section-object-access.md)

