---
title: Style.LinkToListTemplate Method (Word)
keywords: vbawd10.chm153878629
f1_keywords:
- vbawd10.chm153878629
ms.prod: word
api_name:
- Word.Style.LinkToListTemplate
ms.assetid: 1b938b1b-aa8f-655b-123e-fb6f00229e23
ms.date: 06/08/2017
---


# Style.LinkToListTemplate Method (Word)

Links the specified style to a list template so that the style's formatting can be applied to lists.


## Syntax

 _expression_ . **LinkToListTemplate**( **_ListTemplate_** , **_ListLevelNumber_** )

 _expression_ Required. A variable that represents a **[Style](style-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ListTemplate_|Required| **ListTemplate object**|The list template that the style is to be linked to.|
| _ListLevelNumber_|Optional| **Variant**|An integer corresponding to the list level that the style is to be linked to. If this argument is omitted, then the level of the style is used.|

## Example

This example creates a new list template and then links heading styles 1 through 9 to levels 1 through 9. The new list template is then applied to the document. Any paragraphs formatted as heading styles will assume the numbering from the list template.


```vb
Dim ltTemp As ListTemplate 
Dim intLoop As Integer 
 
Set ltTemp = _ 
 ActiveDocument.ListTemplates.Add(OutlineNumbered:=True) 
 
For intLoop = 1 To 9 
 With ltTemp.ListLevels(intLoop) 
 .NumberStyle = wdListNumberStyleArabic 
 .NumberPosition = InchesToPoints(0.25 * (intLoop - 1)) 
 .TextPosition = InchesToPoints(0.25 * intLoop) 
 .NumberFormat = "%" &; intLoop &; "." 
 End With 
 With ActiveDocument.Styles("Heading " &; intLoop) 
 .LinkToListTemplate ListTemplate:=ltTemp 
 End With 
Next intLoop 
 
ActiveDocument.Content.ListFormat.ApplyListTemplate _ 
 ListTemplate:=ltTemp
```


## See also


#### Concepts


[Style Object](style-object-word.md)

