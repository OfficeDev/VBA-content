---
title: Shapes.Range Method (Word)
keywords: vbawd10.chm161415189
f1_keywords:
- vbawd10.chm161415189
ms.prod: word
api_name:
- Word.Shapes.Range
ms.assetid: 277b5e9c-1bec-9e0c-b022-32cef4c5e38e
ms.date: 06/08/2017
---


# Shapes.Range Method (Word)

Returns a  **ShapeRange** object that represents the shapes within a range.


## Syntax

 _expression_ . **Range**( **_Index_** )

 _expression_ Required. A variable that represents a **[Shapes](shapes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|Specifies which shapes are to be included in the specified range. Can be an integer that specifies the index number of a shape within the  **Shapes** collection, a string that specifies the name of a shape, or a array that contains integers or strings.|

### Return Value

ShapeRange


## Remarks

A  **Shape** object always appears on the same page as the range it is anchored to.


 **Note**  Most operations that you can do with a  **Shape** object you can also do with a **ShapeRange** object that contains a single shape. Some operations, when performed on a **ShapeRange** object that contains multiple shapes, produce an error.


## Example

This example sets the fill foreground color to purple for the first shape in the active document.


```vb
Sub ShRange() 
 With ActiveDocument.Shapes.Range(1).Fill 
 .ForeColor.RGB = RGB(255, 0, 255) 
 .Visible = msoTrue 
 End With 
End Sub
```

This example applies a shadow to a variable shape in the active document.




```vb
Sub ShpRange2(strShpName As String) 
 ActiveDocument.Shapes.Range(strShpName).Shadow.Type = msoShadow6 
End Sub
```

To call the preceding subroutine, enter the following code into a standard code module.




```vb
Sub CallShpRange2() 
 Dim shpArrow As Shape 
 Dim strName As String 
 
 Set shpArrow = ActiveDocument.Shapes.AddShape(Type:=msoShapeLeftArrow, _ 
 Left:=200, Top:=400, Width:=50, Height:=75) 
 
 shpArrow.Name = "myShape" 
 strName = shpArrow.Name 
 ShpRange2 strShpName:=strName 
End Sub
```

This example selects shapes one and three in the active document.




```vb
Sub SelectShapeRange() 
 ActiveDocument.Shapes.Range(Array(1, 3)).Select 
End Sub
```

This example selects and deletes the shapes in the first shape in the active document. This example assumes that the first shape is a canvas shape.




```vb
Sub CanvasShapeRange() 
 Dim rngCanvasShapes As Range 
 Set rngCanvasShapes = ActiveDocument.Shapes(1).CanvasItems.Range(1) 
 rngCanvasShapes.Select 
 rngCanvasShapes.Delete 
End Sub
```


## See also


#### Concepts


[Shapes Collection Object](shapes-object-word.md)

