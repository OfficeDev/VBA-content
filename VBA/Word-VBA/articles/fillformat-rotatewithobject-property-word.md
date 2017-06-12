---
title: FillFormat.RotateWithObject Property (Word)
keywords: vbawd10.chm164102265
f1_keywords:
- vbawd10.chm164102265
ms.prod: word
api_name:
- Word.FillFormat.RotateWithObject
ms.assetid: 96a0a7e9-be99-fb36-b245-8850297fa765
ms.date: 06/08/2017
---


# FillFormat.RotateWithObject Property (Word)

Returns or sets whether the fill rotates with the specified shape. Read/write.


## Syntax

 _expression_ . **RotateWithObject**

 _expression_ An expression that returns a **FillFormat** object.


## Remarks

The value returned by the  **RotateWithObject** property can be one of the[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx) constants listed in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The fill does not rotate with the shape.|
| **msoTrue**|The fill rotates with the shape.|
The setting of the  **RotateWithObject** property corresponds to the setting of the **Rotate with shape** box on the **Fill** pane of the **Format Picture** dialog box in the Word user interface (under **Drawing Tools**, on the  **Format** Tab, in the **Shape Styles** group, click **Format Shape**.)


 **Note**  The  **Rotate with shape** box only appears if you have either the **Gradient fill** or **Picture or texture fill** option buttons selected on the **Fill** pane of the **Format Shape** dialog box.


## See also


#### Concepts


[FillFormat Object](fillformat-object-word.md)

