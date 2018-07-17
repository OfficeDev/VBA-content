---
title: TablesOfContents.MarkEntry Method (Word)
keywords: vbawd10.chm152305765
f1_keywords:
- vbawd10.chm152305765
ms.prod: word
api_name:
- Word.TablesOfContents.MarkEntry
ms.assetid: ef8e1d14-82b0-d1f8-8aaf-e2e1b4079c2b
ms.date: 06/08/2017
---


# TablesOfContents.MarkEntry Method (Word)

Inserts a TC (Table of Contents Entry) field after the specified range. The method returns a  **Field** object representing the TC field.


## Syntax

 _expression_ . **MarkEntry**( **_Range_** , **_Entry_** , **_EntryAutoText_** , **_TableID_** , **_Level_** )

 _expression_ Required. A variable that represents a **[TablesOfContents](tablesofcontents-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|The location of the entry. The TC field is inserted after Range.|
| _Entry_|Optional| **Variant**|The text that appears in the table of contents or table of figures. To indicate a subentry, include the main entry text and the subentry text, separated by a colon (:) (for example, "Introduction:The Product").|
| _EntryAutoText_|Optional| **Variant**|The AutoText entry name that includes text for the index, table of figures, or table of contents (Entry is ignored).|
| _TableID_|Optional| **Variant**|A one-letter identifier for the table of figures or table of contents item (for example, "i" for an "illustration").|
| _Level_|Optional| **Variant**|A level for the entry in the table of contents or table of figures.|

### Return Value

Field


## See also


#### Concepts


[TablesOfContents Collection Object](tablesofcontents-object-word.md)

