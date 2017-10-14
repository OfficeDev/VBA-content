---
title: XMLNode.Validate Method (Word)
keywords: vbawd10.chm37748840
f1_keywords:
- vbawd10.chm37748840
ms.prod: word
api_name:
- Word.XMLNode.Validate
ms.assetid: 1a520e28-6b4c-dd95-ba74-cde60e36ad32
ms.date: 06/08/2017
---


# XMLNode.Validate Method (Word)

Validates an individual XML element against the XML schemas that are attached to a document.


## Syntax

 _expression_ . **Validate**

 _expression_ An expression that returns an **XMLNode** object.


### Return Value

Nothing


## Remarks

Use the  **Validate** method with the **[ValidationStatus](xmlnode-validationstatus-property-word.md)** and **[ValidationErrorText](xmlnode-validationerrortext-property-word.md)** properties to determine if an XML element is valid against the applied schema and what error text to display to the user. Use the **[SetValidationError](xmlnode-setvalidationerror-method-word.md)** method to override the schema violations with custom validation errors.

When you run the  **Validate** method, Microsoft Word populates the **[XMLSchemaViolations](http://msdn.microsoft.com/library/9bed9233-4b6b-fe11-d681-8c9f72f99449%28Office.15%29.aspx)** property of the **[Document](document-object-word.md)** object with a collection of the XML nodes that have validation errors.


## Example

The following example checks each element and attribute in the active document and displays a message containing the elements and attributes that do not pass validation, according to the schema, and a description of why.


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

