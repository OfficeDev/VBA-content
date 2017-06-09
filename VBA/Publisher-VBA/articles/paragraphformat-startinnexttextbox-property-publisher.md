---
title: ParagraphFormat.StartInNextTextBox Property (Publisher)
keywords: vbapb10.chm5439539
f1_keywords:
- vbapb10.chm5439539
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.StartInNextTextBox
ms.assetid: 96b34fa8-04ef-e472-16f0-15f82e7912ba
ms.date: 06/08/2017
---


# ParagraphFormat.StartInNextTextBox Property (Publisher)

Returns or sets an  **MsoTriState** constant that represents whether to always start the selected paragraph in the next linked text box. Read/write.


## Syntax

 _expression_. **StartInNextTextBox**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

MsoTriState


## Remarks

The  **StartInNextTextBox** property value can be one of the ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.

If text is added to the previous text box, causing text to overflow into the text box containing the specified text, the specified text (and any text following it) is moved to the top of the next available text box. If no linked text box is available, the specified text (and any text following it) is placed into the text overflow buffer. It will remain in the buffer until either another linked text box is added to the publication, or the  **StartInNextTextBox** property is changed.

This property corresponds to the  **Start in next text box** control in the **Paragraph** dialog box.


