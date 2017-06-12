---
title: Font.GetScriptName Method (Publisher)
keywords: vbapb10.chm5374000
f1_keywords:
- vbapb10.chm5374000
ms.prod: publisher
api_name:
- Publisher.Font.GetScriptName
ms.assetid: 332860de-33fa-7d6a-ac42-28c39856cff7
ms.date: 06/08/2017
---


# Font.GetScriptName Method (Publisher)

Returns a  **String** that represents the name of the font script being used in a text range.


## Syntax

 _expression_. **GetScriptName**( **_Script_**)

 _expression_A variable that represents a  **Font** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Script|Required| **PbFontScriptType**|The script name.|

### Return Value

String


## Remarks

The Script parameter can be one of the  **[PbFontScriptType](pbfontscripttype-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

This example verifies that the default font script in use for the specified text range is Tahoma and, if not, sets it as the default font script.


```vb
Sub GetScript() 
 With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Font 
 If .GetScriptName(Script:=pbFontScriptDefault) <> "Tahoma" Then 
 .SetScriptName Script:=pbFontScriptDefault, _ 
 FontName:="Tahoma" 
 End If 
 End With 
End Sub
```


