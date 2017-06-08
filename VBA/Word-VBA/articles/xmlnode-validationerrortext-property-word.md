---
title: XMLNode.ValidationErrorText Property (Word)
keywords: vbawd10.chm37748760
f1_keywords:
- vbawd10.chm37748760
ms.prod: word
api_name:
- Word.XMLNode.ValidationErrorText
ms.assetid: 85816e71-2629-0f5c-3775-e42f7fb7f9a5
ms.date: 06/08/2017
---


# XMLNode.ValidationErrorText Property (Word)

Returns a  **String** that represents the description for a validation error on an **XMLNode** object.


## Syntax

 _expression_ . **ValidationErrorText**( **_Advanced_** )

 _expression_ A variable that represents a **[XMLNode](xmlnode-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Advanced_|Optional| **Boolean**|Indicates that the error text displayed is the advanced version of the validation error description, which comes from the MSXML 5.0 component included with Microsoft Word.|

## Example

The following example checks each element in the active document and displays a message containing the elements and attributes that do not validate according to the schema and a description of why.


```vb
Dim objNode As XMLNode 
Dim strValid As String 
 
For Each objNode In ActiveDocument.XMLNodes 
 objNode.Validate 
 If objNode.ValidationStatus <> wdXMLValidationStatusOK Then 
 strValid = strValid &; objNode.BaseName &; vbTab &; _ 
 objNode.ValidationErrorText &; vbCrLf 
 End If 
Next 
 
MsgBox "The following elements do not validate against " &; _ 
 "the schema." &; vbCrLf &; vbCrLf &; strValid &; vbCrLf &; _ 
 "You should fix these elements before continuing."
```


## See also


#### Concepts


[XMLNode Object](xmlnode-object-word.md)

