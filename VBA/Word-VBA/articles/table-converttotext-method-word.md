---
title: Table.ConvertToText Method (Word)
keywords: vbawd10.chm156303379
f1_keywords:
- vbawd10.chm156303379
ms.prod: word
api_name:
- Word.Table.ConvertToText
ms.assetid: 750db54e-faca-f1eb-8eb8-3a5c0dbb2c25
ms.date: 06/08/2017
---


# Table.ConvertToText Method (Word)

Converts a table to text and returns a  **Range** object that represents the delimited text.


## Syntax

 _expression_ . **ConvertToText**( **_Separator_** , **_NestedTables_** )

 _expression_ Required. A variable that represents a **[Table](table-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Separator_|Optional| **Variant**|The character that delimits the converted columns (paragraph marks delimit the converted rows). Can be any  **WdTableFieldSeparator** constants.|
| _NestedTables_|Optional| **Variant**| **True** if nested tables are converted to text. This argument is ignored if Separator is not **wdSeparateByParagraphs** . The default value is **True** .|

## Remarks

When you apply the  **ConvertToText** method to a **Table** object, the object is deleted. To maintain a reference to the converted contents of the table, you must assign the **Range** object returned by the **ConvertToText** method to a new object variable. In the following example, the first table in the active document is converted to text and then formatted as a bulleted list.


```vb
Dim tableTemp As Table 
Dim rngTemp As Range 
 
Set tableTemp = ActiveDocument.Tables(1) 
Set rngTemp = _ 
 tableTemp.ConvertToText(Separator:=wdSeparateByParagraphs) 
 
rngTemp.ListFormat.ApplyListTemplate _ 
 ListTemplate:=ListGalleries(wdBulletGallery).ListTemplates(1)
```


## Example

This example creates a table and then converts it to text by using tabs as separator characters.


```vb
Dim docNew As Document 
Dim tableNew As Table 
Dim intTemp As Integer 
Dim cellLoop As Cell 
Dim rngTemp As Range 
 
Set docNew = Documents.Add 
Set tableNew = docNew.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=3, NumColumns:=3) 
 
intTemp = 1 
 
For Each cellLoop In tableNew.Range.Cells 
 cellLoop.Range.InsertAfter "Cell " &; intTemp 
 intTemp = intTemp + 1 
Next cellLoop 
 
MsgBox "Click OK to convert table to text." 
Set rngTemp = _ 
 tableNew.ConvertToText(Separator:=wdSeparateByTabs)
```

This example converts the table that contains the selection to text, with spaces between the columns.




```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Tables(1).ConvertToText Separator:=" " 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


#### Concepts


[Table Object](table-object-word.md)

