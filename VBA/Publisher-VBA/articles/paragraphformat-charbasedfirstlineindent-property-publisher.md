---
title: ParagraphFormat.CharBasedFirstLineIndent Property (Publisher)
keywords: vbapb10.chm5439528
f1_keywords:
- vbapb10.chm5439528
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.CharBasedFirstLineIndent
ms.assetid: d0432be6-2e6a-39fa-9e9a-0300a0437f35
ms.date: 06/08/2017
---


# ParagraphFormat.CharBasedFirstLineIndent Property (Publisher)

Returns or sets the value of the first line indent (in East Asian character width). Read/write  **Long**.


## Syntax

 _expression_. **CharBasedFirstLineIndent**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

Long


## Remarks

The value of  **CharBasedFirstLineIndent** can be returned or set only after the **[UseCharBasedFirstLineIndent](paragraphformat-usecharbasedfirstlineindent-property-publisher.md)** has been set. A run-time "permission denied" error is returned if **UseCharBasedFirstLineIndent** is not set first. Note, however, that **UseCharBasedFirstLineIndent** can be set only if East Asian languages are enabled on the client computer (the value can be returned regardless of whether East Asian languages are enabled). This effectively means that **CharBasedFirstLineIndent** cannot be used unless East Asian languages are enabled on the client computer.

The value of  **CharBasedFirstLineIndent** can range from 0 (zero) to 80.


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


