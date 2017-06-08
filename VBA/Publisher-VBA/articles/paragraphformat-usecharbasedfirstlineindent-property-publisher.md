---
title: ParagraphFormat.UseCharBasedFirstLineIndent Property (Publisher)
keywords: vbapb10.chm5439529
f1_keywords:
- vbapb10.chm5439529
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.UseCharBasedFirstLineIndent
ms.assetid: c2ac44ab-6671-5851-ac62-7449fd646cc5
ms.date: 06/08/2017
---


# ParagraphFormat.UseCharBasedFirstLineIndent Property (Publisher)

Returns or sets an  **MsoTriState** constant that specifies whether a paragraph is indented using East Asian character width. Read/write.


## Syntax

 _expression_. **UseCharBasedFirstLineIndent**

 _expression_A variable that represents an  **ParagraphFormat** object.


### Return Value

MsoTriState


## Remarks

The  **UseCharBasedFirstLineIndent** property value can be one of the ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.

The value of  **UseCharBasedFirstLineIndent** can be set only if East Asian languages are enabled on the client computer, whereas the value can be returned regardless of whether East Asian languages are enabled. Note that **UseCharBasedFirstLineIndent** must be set before the **[CharBasedFirstLineIndent](paragraphformat-charbasedfirstlineindent-property-publisher.md)** property can be returned or set. A run-time "permission denied" error is returned if **UseCharBasedFirstLineIndent** is not set first.

If  **UseCharBasedFirstLineIndent** is **msoTrue**, the paragraph is indented using East Asian character width, and if it is  **msoFalse** it is not. The default value is **msoFalse**.


## Example

The following example creates a text box on the fourth page of the active publication. After the  **UseCharBasedFirstLineIndent** property is set to **True**, the width of the first line indent is set to 15 points by using the  **CharBasedFirstLineIndent** property. Font properties are then set, and text is inserted into the paragraph.


```vb
Dim theTextBox As Shape 
 
Set theTextBox = ActiveDocument.Pages(4).Shapes _ 
 .AddShape(msoShapeRectangle, 100, 100, 300, 200) 
 
With theTextBox 
 .TextFrame.TextRange.ParagraphFormat _ 
 .UseCharBasedFirstLineIndent = msoTrue 
 .TextFrame.TextRange.ParagraphFormat _ 
 .CharBasedFirstLineIndent = 15 
 .TextFrame.TextRange.Font.Name = "Verdana" 
 .TextFrame.TextRange.Font.Size = 12 
 .TextFrame.TextRange.Text = "This is a test sentence." _ 
 &; Chr(13) &; "This is another test sentence." 
End With
```


