---
title: DropCap Object (Publisher)
keywords: vbapb10.chm5570559
f1_keywords:
- vbapb10.chm5570559
ms.prod: publisher
api_name:
- Publisher.DropCap
ms.assetid: 7c6aeffe-cf25-a834-52de-5966df5e21d2
ms.date: 06/08/2017
---


# DropCap Object (Publisher)

Represents a dropped capital letter at the beginning of a paragraph.
 


## Example

Use the  **[DropCap](textrange-dropcap-property-publisher.md)** property to return a **DropCap** object. The following example sets a dropped capital letter for the first letter of each paragraph in the first shape on the first page of the active publication. This example assumes that the specified shape is a text box and not another type of shape.
 

 

```
Sub ApplyDropCap() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .DropCap.ApplyCustomDropCap Size:=3, Span:=3, Bold:=True 
End Sub
```


## Methods



|**Name**|
|:-----|
|[ApplyCustomDropCap](dropcap-applycustomdropcap-method-publisher.md)|
|[Clear](dropcap-clear-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](dropcap-application-property-publisher.md)|
|[FontBold](dropcap-fontbold-property-publisher.md)|
|[FontColor](dropcap-fontcolor-property-publisher.md)|
|[FontItalic](dropcap-fontitalic-property-publisher.md)|
|[FontName](dropcap-fontname-property-publisher.md)|
|[LinesUp](dropcap-linesup-property-publisher.md)|
|[Parent](dropcap-parent-property-publisher.md)|
|[Size](dropcap-size-property-publisher.md)|
|[Span](dropcap-span-property-publisher.md)|

