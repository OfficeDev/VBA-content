---
title: ParagraphFormat.WidowControl Property (Publisher)
keywords: vbapb10.chm5439536
f1_keywords:
- vbapb10.chm5439536
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.WidowControl
ms.assetid: af1f1106-60e3-3987-3710-30fae7cf3940
ms.date: 06/08/2017
---


# ParagraphFormat.WidowControl Property (Publisher)

Sets or returns an  **MsoTriState** that represents whether or not the first or last line of the specified paragraph can appear by itself in a text box. Read/write.


## Syntax

 _expression_. **WidowControl**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

MsoTriState


## Remarks

This option ensures that the first or last line of the specified paragraph will not appear by itself in a text frame. For example, if the last line in a specified paragraph is the first line of a widow controlled paragraph, a second line will be moved to the next text frame with it.

The  **WidowControl** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The first or last line may appear by itself in a text box.|
| **msoTrue**|The first or last line will not appear by itself in a text box.|
The default setting for this property is  **msoFalse**.


## Example

This example sets the  **WidowControl** property to **msoTrue** for the specified **ParagraphFormat** object.


```vb
Dim objParaForm As ParagraphFormat 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Paragraphs(1).ParagraphFormat 
objParaForm.WidowControl = msoTrue 

```


