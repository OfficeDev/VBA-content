---
title: ParagraphFormat.LockToBaseLine Property (Publisher)
keywords: vbapb10.chm5439540
f1_keywords:
- vbapb10.chm5439540
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.LockToBaseLine
ms.assetid: 4430bab6-a338-e61d-681c-6063d4a5c3b3
ms.date: 06/08/2017
---


# ParagraphFormat.LockToBaseLine Property (Publisher)

Returns an  **MsoTristate** that represents whether or not text will be positioned along baseline guides. Read/write.


## Syntax

 _expression_. **LockToBaseLine**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

MsoTristate


## Remarks

The  **LockToBaseLine** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**| The text is not aligned to baselines.|
| **msoTriStateMixed**|The specified paragraphs contain both text that is aligned to baselines and text that is not aligned to baselines.|
| **msoTrue**|The text is aligned to baselines.|

## Example

The following example sets the  **LockToBaseLine** property to **True**.


```vb
Dim objParaForm As ParagraphFormat 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.ParagraphFormat 
objParaForm.LockToBaseLine = msoTrue 

```


