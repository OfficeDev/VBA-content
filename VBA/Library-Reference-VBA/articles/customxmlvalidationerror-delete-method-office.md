---
title: CustomXMLValidationError.Delete Method (Office)
keywords: vbaof11.chm307006
f1_keywords:
- vbaof11.chm307006
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CustomXMLValidationError.Delete
ms.assetid: d425c0f8-6eb1-9e1d-5246-3ba77bbf3cd3
---


# CustomXMLValidationError.Delete Method (Office)

Deletes the  **CustomXMLValidationError** object representing a data validation error.


## Syntax

 _expression_. **Delete**

 _expression_ An expression that returns a **CustomXMLValidationError** object.


## Example

The following example deletes the validation error containing specific text.


```vb
Dim objCustomXMLValidationError as CustomXMLValidationError 
 
' Deletes the specified error message. 
objCustomXMLValidationError.Text("To add content to this data stream, it must be valid, well-formed XML." ).Delete
```


## See also


#### Concepts


[CustomXMLValidationError Object](customxmlvalidationerror-object-office.md)

