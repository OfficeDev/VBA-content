---
title: Page.Header Property (Publisher)
keywords: vbapb10.chm393247
f1_keywords:
- vbapb10.chm393247
ms.prod: publisher
api_name:
- Publisher.Page.Header
ms.assetid: f10806eb-972a-d482-935c-95d5ccbbbb36
ms.date: 06/08/2017
---


# Page.Header Property (Publisher)

Returns a  **HeaderFooter** object representing the header of the specified **Page** object. Read-only.


## Syntax

 _expression_. **Header**

 _expression_A variable that represents a  **Page** object.


### Return Value

HeaderFooter


## Remarks

This property is only for master pages. A "This feature is only for master pages" error is returned if the header property is accessed from a  **Page** object that is returned form the **Pages** collection. A new **HeaderFooter** object is created for the specified master page by accessing this property.


## Example

The following example creates a  **HeaderFooter** object and sets it to the header of the first master page.


```vb
Dim objHeader As HeaderFooter 
Set objHeader = ActiveDocument.MasterPages(1).Header
```

The  **HeaderFooter** object returned by the **Header** property can be used to manipulate the header content. The following example sets some properties of the **HeaderFooter** object of the first master page.




```vb
With ActiveDocument.masterPages(1) 
 With .Header 
 .TextRange.Text = "Windows" &; Chr(13) &; "Office" &; Chr(13) &; "Internet Explorer" 
 With .TextRange.ParagraphFormat 
 .SetListType Value:=pbListTypeBullet, BulletText:="*" 
 .Alignment = pbParagraphAlignmentLeft 
 End With 
 End With 
 With .Footer 
 .TextRange.Hyperlinks.Add Text:=.TextRange, _ 
 Address:="http://www.tailspintoys.com", _ 
 TextToDisplay:="Tailspin" 
 End With 
End With
```


