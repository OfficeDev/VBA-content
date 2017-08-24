---
title: TextFrame Object (Publisher)
keywords: vbapb10.chm3932159
f1_keywords:
- vbapb10.chm3932159
ms.prod: publisher
api_name:
- Publisher.TextFrame
ms.assetid: 95e88f5a-b3dc-272e-7c1d-5282c97ae11e
ms.date: 06/08/2017
---


# TextFrame Object (Publisher)

Represents the text frame in a  **[Shape](http://msdn.microsoft.com/library/666cb7f0-62a8-f419-9838-007ef29506ee%28Office.15%29.aspx)** object. Contains the text in the text frame and the properties that control the margins and orientation of the text frame.


## Example

Use the  **[TextFrame](http://msdn.microsoft.com/library/fc654905-d56b-9a6c-28fa-4b54bf2a8686%28Office.15%29.aspx)** property to return the **TextFrame** object for a shape. The **[TextRange](http://msdn.microsoft.com/library/44a8395e-81dc-7d06-f068-89f77a889f5e%28Office.15%29.aspx)** property returns a **[TextRange](textrange-object-publisher.md)** object that represents the range of text inside the specified text frame. The following example adds text to the text frame of shape one in the active publication, and then formats the new text.


```
Sub AddTextToTextFrame() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange 
 .Text = "My Text" 
 With .Font 
 .Bold = msoTrue 
 .Size = 25 
 .Name = "Arial" 
 End With 
 End With 
End Sub
```


 **Note**  Some shapes do not support attached text (lines, freeforms, pictures, and OLE objects, for example). If you attempt to return or set properties that control text in a text frame for those objects, an error occurs.

Use the  **[HasTextFrame](http://msdn.microsoft.com/library/faf9514a-438b-ad12-a830-ed34cea8ba03%28Office.15%29.aspx)** property to determine whether the shape has a text frame and use the **[HasText](http://msdn.microsoft.com/library/f8d1c660-c3f1-e835-adc3-114e6611de98%28Office.15%29.aspx)** property to determine whether the text frame contains text, as shown in the following example.




```
Sub GetTextFromTextFrame() 
 Dim shpText As Shape 
 
 For Each shpText In ActiveDocument.Pages(1).Shapes 
 If shpText.HasTextFrame = msoTrue Then 
 With shpText.TextFrame 
 If .HasText Then MsgBox .TextRange.Text 
 End With 
 End If 
 Next 
End Sub
```

Text frames can be linked together so that the text flows from the text frame of one shape into the text frame of another shape. Use the  **[NextLinkedTextFrame](http://msdn.microsoft.com/library/5ba08ab5-8515-4efe-59a3-79a11f6a7c4e%28Office.15%29.aspx)** and **[PreviousLinkedTextFrame](http://msdn.microsoft.com/library/00947ec3-fcff-4451-491b-5b7748ccb74e%28Office.15%29.aspx)** properties to link text frames. The following example creates a text box (a rectangle with a text frame) and adds some text to it. It then creates another text box and links the two text frames together so that the text flows from the first text frame into the second one.




```
Sub LinkTextBoxes() 
 Dim shpTextBox1 As Shape 
 Dim shpTextBox2 As Shape 
 
 Set shpTextBox1 = ActiveDocument.Pages(1).Shapes.AddTextbox _ 
 (msoTextOrientationHorizontal, 72, 72, 72, 36) 
 shpTextBox1.TextFrame.TextRange.Text = _ 
 "This is some text. This is some more text." 
 
 Set shpTextBox2 = ActiveDocument.Pages(1).Shapes.AddTextbox _ 
 (msoTextOrientationHorizontal, 72, 144, 72, 36) 
 shpTextBox1.TextFrame.NextLinkedTextFrame = shpTextBox2 _ 
 .TextFrame 
End Sub
```


## Methods



|**Name**|
|:-----|
|[BreakForwardLink](http://msdn.microsoft.com/library/60a7a798-ebd3-e00d-032d-685dd0d5a042%28Office.15%29.aspx)|
|[ValidLinkTarget](http://msdn.microsoft.com/library/ee946f58-669f-7150-0f40-2dd3b857e274%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/14b41c64-cdd3-f1ab-202c-49f18d72d035%28Office.15%29.aspx)|
|[AutoFitText](http://msdn.microsoft.com/library/468a9d3e-cb9d-8147-60ea-eb839d691e7a%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/b025f208-3ca4-c0f1-e01e-023931c4c545%28Office.15%29.aspx)|
|[ColumnSpacing](http://msdn.microsoft.com/library/3b650d29-3716-e9b1-eaf0-92bdc0b77c5f%28Office.15%29.aspx)|
|[HasNextLink](http://msdn.microsoft.com/library/907ec470-e283-906a-e25f-f5a8548a18a4%28Office.15%29.aspx)|
|[HasPreviousLink](http://msdn.microsoft.com/library/85e0b497-55c9-d49f-2b65-e199361c121a%28Office.15%29.aspx)|
|[HasText](http://msdn.microsoft.com/library/f8d1c660-c3f1-e835-adc3-114e6611de98%28Office.15%29.aspx)|
|[IncludeContinuedFromPage](http://msdn.microsoft.com/library/7c129bf2-60da-4170-1410-94961ccf3345%28Office.15%29.aspx)|
|[IncludeContinuedOnPage](http://msdn.microsoft.com/library/defa0bd7-abe7-ac2a-97a1-de5c5f0df790%28Office.15%29.aspx)|
|[MarginBottom](http://msdn.microsoft.com/library/55858bba-1103-48ba-64d6-5cc5ab677867%28Office.15%29.aspx)|
|[MarginLeft](http://msdn.microsoft.com/library/4e784b9f-9467-5a14-c211-589e69c3b8bc%28Office.15%29.aspx)|
|[MarginRight](http://msdn.microsoft.com/library/bdbde217-6a51-7823-ac93-8bbffa583544%28Office.15%29.aspx)|
|[MarginTop](http://msdn.microsoft.com/library/9709eefe-0857-f228-aa56-780c4789a413%28Office.15%29.aspx)|
|[NextLinkedTextFrame](http://msdn.microsoft.com/library/5ba08ab5-8515-4efe-59a3-79a11f6a7c4e%28Office.15%29.aspx)|
|[Orientation](http://msdn.microsoft.com/library/f510e624-6322-4054-5e7f-8688c5ea817a%28Office.15%29.aspx)|
|[Overflowing](http://msdn.microsoft.com/library/5a0f053b-519a-1637-0d73-992c56cdd7f0%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/c4d2d0bd-7a6b-201c-4b1b-416490ab8023%28Office.15%29.aspx)|
|[PreviousLinkedTextFrame](http://msdn.microsoft.com/library/00947ec3-fcff-4451-491b-5b7748ccb74e%28Office.15%29.aspx)|
|[Story](http://msdn.microsoft.com/library/7bbe0967-83aa-745b-ad13-8a7dfe61811c%28Office.15%29.aspx)|
|[TextRange](http://msdn.microsoft.com/library/44a8395e-81dc-7d06-f068-89f77a889f5e%28Office.15%29.aspx)|
|[VerticalTextAlignment](http://msdn.microsoft.com/library/cd809f00-b092-c483-fe99-2aa8043fb684%28Office.15%29.aspx)|

