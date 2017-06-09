---
title: HeadingStyles.Add Method (Word)
keywords: vbawd10.chm160039012
f1_keywords:
- vbawd10.chm160039012
ms.prod: word
api_name:
- Word.HeadingStyles.Add
ms.assetid: 1ad89871-cd73-4159-e85f-e0cdbe3633af
ms.date: 06/08/2017
---


# HeadingStyles.Add Method (Word)

Returns a  **HeadingStyle** object that represents a new heading style added to a document. The new heading style will be included whenever you compile a table of contents or table of figures.


## Syntax

 _expression_ . **Add**( **_Style_** , **_Level_** )

 _expression_ Required. A variable that represents a **[HeadingStyles](headingstyles-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **Variant**|The style you want to add. You can specify this argument by using either the string name for the style or a  **Style** object.|
| _Level_|Required| **Integer**|A number that represents the level of the heading.|

### Return Value

HeadingStyle


## Example

This example adds a table of contents at the beginning of the active document and then adds the Title style to the list of styles used to build a table of contents.


```vb
Set myToc = ActiveDocument.TablesOfContents _ 
 .Add(Range:=ActiveDocument.Range(0, 0), _ 
 UseHeadingStyles:=True, UpperHeadingLevel:=1, _ 
 LowerHeadingLevel:=3) 
myToc.HeadingStyles.Add Style:="Title", Level:=2
```


## See also


#### Concepts


[HeadingStyles Collection Object](headingstyles-object-word.md)

