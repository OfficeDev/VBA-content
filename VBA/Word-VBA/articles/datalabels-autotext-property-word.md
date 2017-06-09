---
title: DataLabels.AutoText Property (Word)
keywords: vbawd10.chm207487111
f1_keywords:
- vbawd10.chm207487111
ms.prod: word
api_name:
- Word.DataLabels.AutoText
ms.assetid: fa26ac03-bf5f-579f-12b5-d7888aa9de9b
ms.date: 06/08/2017
---


# DataLabels.AutoText Property (Word)

 **True** if all objects in the collection automatically generate appropriate text based on context. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoText**

 _expression_ A variable that represents a **[DataLabels](datalabels-object-word.md)** object.


## Remarks

Setting the value of this property sets the  **[AutoText](datalabel-autotext-property-word.md)** property of all **[DataLabel](datalabel-object-word.md)** objects contained by the collection. This property returns **True** only when the **AutoText** property for all **DataLabel** objects contained in the collection is set to **True** ; otherwise, this property returns **False** .


## Example

The following example sets the data labels for series one of the first chart in the active document to automatically generate appropriate text.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1). _ 
 DataLabels.AutoText = True 
 End If 
End With
```


## See also


#### Concepts


[DataLabels Object](datalabels-object-word.md)

