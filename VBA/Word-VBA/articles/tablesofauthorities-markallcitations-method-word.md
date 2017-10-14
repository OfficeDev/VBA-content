---
title: TablesOfAuthorities.MarkAllCitations Method (Word)
keywords: vbawd10.chm152174694
f1_keywords:
- vbawd10.chm152174694
ms.prod: word
api_name:
- Word.TablesOfAuthorities.MarkAllCitations
ms.assetid: 5f07956b-2e51-f88e-f758-a2ee055d7a36
ms.date: 06/08/2017
---


# TablesOfAuthorities.MarkAllCitations Method (Word)

Inserts a TA (Table of Authorities Entry) field after all instances of the  **ShortCitation** text.


## Syntax

 _expression_ . **MarkAllCitations**( **_ShortCitation_** , **_LongCitation_** , **_LongCitationAutoText_** , **_Category_** )

 _expression_ Required. A variable that represents a **[TablesOfAuthorities](tablesofauthorities-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ShortCitation_|Required| **String**|The short citation for the entry as it will appear in the  **Mark Citation** dialog box ( **Insert** menu, **Index and Tables** command).|
| _LongCitation_|Optional| **Variant**|The long citation string for the entry as it will appear in the table of authorities.|
| _LongCitationAutoText_|Optional| **Variant**|The AutoText entry name that contains the text of the long citation as it will appear in the table of authorities.|
| _Category_|Optional| **Variant**|The category number to be associated with the entry: 1 corresponds to the first category in the  **Category** box in the **Mark Citation** dialog box, 2 corresponds to the second category, and so on.|

## Example

This example marks all instances of "Forrester v. Craddock" in the active document with a TA field that references the "Forrester v. Craddock, 51 Wn. 2d 315 (1957)" citation.


```vb
ActiveDocument.TablesOfAuthorities.MarkAllCitations _ 
 ShortCitation:="Forrester v. Craddock", Category:=1, _ 
 LongCitation:="Forrester v. Craddock, 51 Wn. 2d 315 (1957)"
```


## See also


#### Concepts


[TablesOfAuthorities Collection Object](tablesofauthorities-object-word.md)

