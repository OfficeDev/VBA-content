---
title: StyleSheets.Add Method (Word)
keywords: vbawd10.chm209584130
f1_keywords:
- vbawd10.chm209584130
ms.prod: word
api_name:
- Word.StyleSheets.Add
ms.assetid: 82659cfc-6681-93c8-299c-f570f23016b2
ms.date: 06/08/2017
---


# StyleSheets.Add Method (Word)

Returns a  **StyleSheet** object that represents a new style sheet added to a Web document.


## Syntax

 _expression_ . **Add**( **_FileName_** , **_LinkType_** , **_Title_** , **_Precedence_** )

 _expression_ Required. A variable that represents a **[StyleSheets](stylesheets-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The path and file name of the cascading style sheet.|
| _LinkType_|Required| **WdStyleSheetLinkType**|Indicates whether the style sheet should be added as a link or imported into the Web document.|
| _Title_|Required| **String**|The name of the style sheet.|
| _Precedence_|Required| **WdStyleSheetPrecedence**|Indicates the level of importance compared with other cascading style sheets attached to the Web document.|

### Return Value

StyleSheet


## Example

This example adds a style sheet to the active document and places it highest in the list of style sheets attached to the document. This example assumes that you have a style sheet document named Website.css located on your drive C.


```vb
Sub NewStylesheet() 
 ActiveDocument.StyleSheets.Add _ 
 FileName:="c:\WebSite.css", _ 
 Precedence:=wdStyleSheetPrecedenceHighest, _ 
 LinkType:=wdStyleSheetLinkTypeLinked, _ 
 Title:="Test Stylesheet" 
End Sub
```


## See also


#### Concepts


[StyleSheets Collection](stylesheets-object-word.md)

