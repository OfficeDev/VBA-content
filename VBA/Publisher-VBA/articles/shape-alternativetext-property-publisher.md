---
title: Shape.AlternativeText Property (Publisher)
keywords: vbapb10.chm2228320
f1_keywords:
- vbapb10.chm2228320
ms.prod: publisher
api_name:
- Publisher.Shape.AlternativeText
ms.assetid: 13bc57af-7067-d60c-5096-a68b1f821d58
ms.date: 06/08/2017
---


# Shape.AlternativeText Property (Publisher)

Returns or sets a  **String** representing the text displayed by a Web browser in place of the **Shape** object while the **Shape** object is being downloaded or when graphics are turned off. Read/write.


## Syntax

 _expression_. **AlternativeText**

 _expression_A variable that represents a  **Shape** object.


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


