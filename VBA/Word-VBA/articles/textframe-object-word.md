---
title: TextFrame Object (Word)
keywords: vbawd10.chm2482
f1_keywords:
- vbawd10.chm2482
ms.prod: word
api_name:
- Word.TextFrame
ms.assetid: 46f7e410-80d9-9fe9-2224-488b623f8592
ms.date: 06/08/2017
---


# TextFrame Object (Word)

Represents the text frame in a  **Shape** object. The **TextFrame** object contains the text in the text frame and the properties that control the margins and orientation of the text frame.


## Remarks

Use the  **TextFrame** property to return the **TextFrame** object for a shape. The **TextRange** property returns a **[Range](range-object-word.md)** object that represents the range of text inside the specified text frame. The following example adds text to the text frame of shape one in the active document.


```
ActiveDocument.Shapes(1).TextFrame.TextRange.Text = "My Text"
```


 **Note**  Some shapes do not support attached text (lines, freeforms, pictures, and OLE objects, for example). If you attempt to return or set properties that control text in a text frame for those objects, an error occurs.

Use the  **HasText** property to determine whether the text frame contains text, as shown in the following example.




```
For Each s In ActiveDocument.Shapes 
 With s.TextFrame 
 If .HasText Then MsgBox .TextRange.Text 
 End With 
Next
```

Text frames can be linked together so that the text flows from the text frame of one shape into the text frame of another shape. Use the  **Next** and **Previous** properties to link text frames. The following example creates a text box (a rectangle with a text frame) and adds some text to it. It then creates another text box and links the two text frames together so that the text flows from the first text frame into the second one.




```
Set myTB1 = ActiveDocument.Shapes.AddTextbox _ 
 (msoTextOrientationHorizontal, 72, 72, 72, 36) 
myTB1.TextFrame.TextRange = _ 
 "This is some text. This is some more text." 
Set myTB2 = ActiveDocument.Shapes.AddTextbox _ 
 (msoTextOrientationHorizontal, 72, 144, 72, 36) 
myTB1.TextFrame.Next = myTB2.TextFrame
```

Use the  **ContainingRange** property to return a **Range** object that represents the entire story that flows between linked text frames. The following example checks the spelling of the text in TextBox 3 and of any other text that is linked to TextBox 3.




```
Set myStory = ActiveDocument.Shapes("TextBox 3") _ 
 .TextFrame.ContainingRange 
myStory.CheckSpelling
```


## Methods



|**Name**|
|:-----|
|[BreakForwardLink](textframe-breakforwardlink-method-word.md)|
|[DeleteText](textframe-deletetext-method-word.md)|
|[ValidLinkTarget](textframe-validlinktarget-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](textframe-application-property-word.md)|
|[AutoSize](textframe-autosize-property-word.md)|
|[Column](textframe-column-property-word.md)|
|[ContainingRange](textframe-containingrange-property-word.md)|
|[Creator](textframe-creator-property-word.md)|
|[HasText](textframe-hastext-property-word.md)|
|[HorizontalAnchor](textframe-horizontalanchor-property-word.md)|
|[MarginBottom](textframe-marginbottom-property-word.md)|
|[MarginLeft](textframe-marginleft-property-word.md)|
|[MarginRight](textframe-marginright-property-word.md)|
|[MarginTop](textframe-margintop-property-word.md)|
|[Next](textframe-next-property-word.md)|
|[NoTextRotation](textframe-notextrotation-property-word.md)|
|[Orientation](textframe-orientation-property-word.md)|
|[Overflowing](textframe-overflowing-property-word.md)|
|[Parent](textframe-parent-property-word.md)|
|[PathFormat](textframe-pathformat-property-word.md)|
|[Previous](textframe-previous-property-word.md)|
|[TextRange](textframe-textrange-property-word.md)|
|[ThreeD](textframe-threed-property-word.md)|
|[VerticalAnchor](textframe-verticalanchor-property-word.md)|
|[WarpFormat](textframe-warpformat-property-word.md)|
|[WordWrap](textframe-wordwrap-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
