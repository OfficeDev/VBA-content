---
title: ShapeRange.AlternativeText Property (Publisher)
keywords: vbapb10.chm2293856
f1_keywords:
- vbapb10.chm2293856
ms.prod: publisher
api_name:
- Publisher.ShapeRange.AlternativeText
ms.assetid: 94cbb99b-3b35-76bb-e269-db8295b84f2f
ms.date: 06/08/2017
---


# ShapeRange.AlternativeText Property (Publisher)

Returns or sets a  **String** representing the text displayed by a Web browser in place of the **Shape** object while the **Shape** object is being downloaded or when graphics are turned off. Read/write.


## Syntax

 _expression_. **AlternativeText**

 _expression_A variable that represents a  **ShapeRange** object.


## Remarks

The maximum length of the  **AlternativeText** property is 254 characters. Microsoft Publisher returns an error if the text length exceeds this number.


## Example

This example sets the alternative text for the selected shape in the active document. This example assumes that you have a publication that the selected shape is a picture of a duck.


```vb
Public Sub Alternative_Text() 
 
 ' The picture of a duck must be selected. 
 Publisher.ActiveDocument.Selection.ShapeRange _ 
 .AlternativeText = "This is a mallard duck." 
 
End Sub
```


