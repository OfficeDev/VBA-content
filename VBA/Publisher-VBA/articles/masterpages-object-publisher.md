---
title: MasterPages Object (Publisher)
keywords: vbapb10.chm655359
f1_keywords:
- vbapb10.chm655359
ms.prod: publisher
api_name:
- Publisher.MasterPages
ms.assetid: 3a7e6021-cbe4-4700-018c-c91d2f7d908a
ms.date: 06/08/2017
---


# MasterPages Object (Publisher)

Represents the page master for a publication after which all pages in the publication will be designed. The  **MasterPages** object is a collection of **[Page](page-object-publisher.md)** objects.
 


## Example

Use the  **[MasterPages](document-masterpages-property-publisher.md)** property to return a **MasterPages** object. The following example adds two ruler guides to the master page so that each page in the active publication is divided into quarters.
 

 

```
Sub ChangeMasterPage() 
 Dim intWidth As Integer 
 Dim intHeight As Integer 
 
 With ActiveDocument 
 intWidth = .PageSetup.PageWidth 
 intWidth = intWidth / 2 
 intHeight = .PageSetup.PageHeight 
 intHeight = intHeight / 2 
 With .MasterPages(1).RulerGuides 
 .Add Position:=intWidth, _ 
 Type:=pbRulerGuideTypeVertical 
 .Add Position:=intHeight, _ 
 Type:=pbRulerGuideTypeHorizontal 
 End With 
 End With 
End Sub
```

Use the  **[Shapes](page-shapes-property-publisher.md)** property to work with AutoShapes and text boxes on the master page. This example adds a small red heart shape to the upper left corner of the master page that will appear on each page in the active publication.
 

 



```
Sub AddShapeToMasterPage() 
 ActiveDocument.MasterPages(1).Shapes.AddShape(Type:=msoShapeHeart, _ 
 Left:=36, Top:=36, Width:=36, Height:=36).Fill _ 
 .ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Add](masterpages-add-method-publisher.md)|
|[FindByPageID](masterpages-findbypageid-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](masterpages-application-property-publisher.md)|
|[Count](masterpages-count-property-publisher.md)|
|[Item](masterpages-item-property-publisher.md)|
|[Parent](masterpages-parent-property-publisher.md)|

