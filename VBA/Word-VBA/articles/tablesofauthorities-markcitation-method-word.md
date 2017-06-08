---
title: TablesOfAuthorities.MarkCitation Method (Word)
keywords: vbawd10.chm152174693
f1_keywords:
- vbawd10.chm152174693
ms.prod: word
api_name:
- Word.TablesOfAuthorities.MarkCitation
ms.assetid: 6dbbd99e-11c2-803a-fb31-e486ba530585
ms.date: 06/08/2017
---


# TablesOfAuthorities.MarkCitation Method (Word)

Inserts a TA (Table of Authorities Entry) field and returns the field as a  **Field** object.


## Syntax

 _expression_ . **MarkCitation**( **_Range_** , **_ShortCitation_** , **_LongCitation_** , **_LongCitationAutoText_** , **_Category_** )

 _expression_ Required. A variable that represents a **[TablesOfAuthorities](tablesofauthorities-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|The location of the table of authorities entry. The TA field is inserted after Range.|
| _ShortCitation_|Required| **String**|The short citation for the entry as it will appear in the  **Mark Citation** dialog box ( **Insert** menu, **Index and Tables** command).|
| _LongCitation_|Optional| **Variant**|The long citation for the entry as it will appear in the table of authorities.|
| _LongCitationAutoText_|Optional| **Variant**|The name of the AutoText entry that contains the text of the long citation as it will appear in the table of authorities.|
| _Category_|Optional| **Variant**|The category number to be associated with the entry: 1 corresponds to the first category in the  **Category** box in the **Mark Citation** dialog box, 2 corresponds to the second category, and so on.|

## Example

This example inserts a table of authorities entry (a TA field) that references the selected text. The long citation text is set to "Forrester v. Craddock" and the category is set to Cases.


```vb
ActiveDocument.TablesOfAuthorities.MarkCitation _ 
 Range:=Selection.Range, ShortCitation:=Selection.Range.Text, _ 
 LongCitation:="Forrester v. Craddock", Category:=1
```

This example inserts a table of authorities entry that references the selected text. The entry text that appears in the table of authorities is the text typed into the input box and the category is set to Other Authorities.




```vb
Dim strCitation As String 
 
strCitation = InputBox("Type citation text") 
ActiveDocument.TablesOfAuthorities.MarkCitation _ 
 Range:=Selection.Range, ShortCitation:=Selection.Range.Text, _ 
 LongCitation:=strCitation, Category:=3
```


## See also


#### Concepts


[TablesOfAuthorities Collection Object](tablesofauthorities-object-word.md)

