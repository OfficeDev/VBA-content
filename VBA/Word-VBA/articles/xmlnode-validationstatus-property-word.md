---
title: XMLNode.ValidationStatus Property (Word)
keywords: vbawd10.chm37748758
f1_keywords:
- vbawd10.chm37748758
ms.prod: word
api_name:
- Word.XMLNode.ValidationStatus
ms.assetid: 795114a1-09d3-f2c6-3572-4a8929ca062c
ms.date: 06/08/2017
---


# XMLNode.ValidationStatus Property (Word)

 Returns a **WdXMLValidationStatus** constant that represents whether an element or attribute is valid according to the attached schema.


## Syntax

 _expression_ . **ValidationStatus**

 _expression_ Required. A variable that represents a **[XMLNode](xmlnode-object-word.md)** object.


## Remarks

This property can return either of the two following  **WdXMLValidationStatus** constants.



| **wdXMLValidationStatusCustom**|Indicates that the  **SetValidationError** method was used to set **ValidationErrorText** property to a custom text string.|
| **wdXMLValidationStatusOK**|Indicates an XML element or attribute is valid according to the attached schema.|
While these are the only two named constants the  **ValidationStatus** property allows, there are many more unnamed values that come from the MSXML 5.0 component included with Microsoft Word. For a more complete list of possible values and their corresponding meaning, refer to the Microsoft Word XML schema reference on the Microsoft Developer Network (MSDN) Web site.


## Example

The following example checks each element in the active document and displays a message containing the elements that do not validate according to the schema and a description of why.


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

