---
title: TextEffectFormat.KernedPairs Property (Publisher)
keywords: vbapb10.chm3735813
f1_keywords:
- vbapb10.chm3735813
ms.prod: publisher
api_name:
- Publisher.TextEffectFormat.KernedPairs
ms.assetid: 1382ae7a-250f-ca08-a57f-f7132078e3f2
ms.date: 06/08/2017
---


# TextEffectFormat.KernedPairs Property (Publisher)

Sets or returns an  **MsoTriState** constant that indicates whether character pairs in a WordArt object have been kerned. Read/write.


## Syntax

 _expression_. **KernedPairs**

 _expression_A variable that represents a  **TextEffectFormat** object.


### Return Value

MsoTriState


## Remarks

The  **KernedPairs** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**| Character pairs in the specified WordArt object have not been kerned.|
| **msoTriStateToggle**|Switches between  **msoTrue** and **msoFalse**.|
| **msoTrue**|Character pairs in the specified WordArt object have been kerned.|

## Example

This example turns on character pair kerning for all WordArt objects in the active publication.


```vb
Sub KernedWordArt() 
 Dim pagPage As Page 
 Dim shpShape As Shape 
 For Each pagPage In ActiveDocument.Pages 
 For Each shpShape In pagPage.Shapes 
 If shpShape.Type = msoTextEffect Then 
 shpShape.TextEffect.KernedPairs = msoTrue 
 End If 
 Next 
 Next 
End Sub
```


