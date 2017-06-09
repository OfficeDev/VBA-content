---
title: Hyperlink.TextToDisplay Property (PowerPoint)
keywords: vbapp10.chm526009
f1_keywords:
- vbapp10.chm526009
ms.prod: powerpoint
api_name:
- PowerPoint.Hyperlink.TextToDisplay
ms.assetid: 5f30033e-ddb8-8814-9e55-e0137ff6fa48
ms.date: 06/08/2017
---


# Hyperlink.TextToDisplay Property (PowerPoint)

Returns or sets the display text for a hyperlink not associated with a graphic. Read/write.


## Syntax

 _expression_. **TextToDisplay**

 _expression_ A variable that represents a **Hyperlink** object.


### Return Value

String


## Remarks

This property will cause a run-time error if used with a hyperlink that is not associated with a text range. You can use code similar to the following to test whether or not a given hyperlink, represented here by  `myHyperlink`, is associated with a text range.

 `If TypeName(myHyperlink.Parent.Parent) = "TextRange" Then strTRtest = "True" End If`


## Example

This example creates an associated hyperlink with the text in shape two on slide one. It then sets the display text to "Microsoft Home Page" and sets the hyperlink address to the correct URL.


```vb
With ActivePresentation.Slides(1).Shapes(2) _
        .TextFrame.TextRange
    With .ActionSettings(ppMouseClick)
        .Action = ppActionHyperlink
        .Hyperlink.TextToDisplay = "Microsoft Home Page"
        .Hyperlink.Address = "http://www.microsoft.com/"
    End With
End With
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-powerpoint.md)

