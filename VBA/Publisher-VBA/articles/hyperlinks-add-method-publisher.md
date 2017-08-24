---
title: Hyperlinks.Add Method (Publisher)
keywords: vbapb10.chm6881284
f1_keywords:
- vbapb10.chm6881284
ms.prod: publisher
api_name:
- Publisher.Hyperlinks.Add
ms.assetid: f5a8cc01-a571-623d-bfab-fe48e43a21b1
ms.date: 06/08/2017
---


# Hyperlinks.Add Method (Publisher)

Adds a new  **Hyperlink** object to the specified **Hyperlinks** collection and returns the new **Hyperlink** object.


## Syntax

 _expression_. **Add**( **_Text_**,  **_Address_**,  **_RelativePage_**,  **_PageID_**,  **_TextToDisplay_**)

 _expression_A variable that represents a  **Hyperlinks** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Text|Required| **TextRange**| **TextRange** object. The text range to be converted into a hyperlink.|
|Address|Optional| **String**|The address of the new hyperlink. If RelativePage is  **pbHlinkTargetTypeURL** (default) or **pbHlinkTargetTypeEmail**, Address must be specified or an error occurs.|
|RelativePage|Optional| **PbHlinkTargetType**| The type of hyperlink to add.|
|PageID|Optional| **Long**|The page ID of the destination page for the new hyperlink. If RelativePage is  **pbHlinkTargetTypePageID**, PageID must be specified or an error occurs. The page ID corresponds to the  **[PageID](hyperlink-pageid-property-publisher.md)** property of the destination page.|
|TextToDisplay|Optional| **String**|The display text of the new hyperlink. If specified,  **TextToDisplay** replaces the text range specified by the **Text** argument.|

### Return Value

Hyperlink


## Remarks

RelativePage can be one of these  **PbHlinkTargetType** constants. The default is **pbHlinkTargetTypeURL**.



| **pbHlinkTargetTypeEmail**|
| **pbHlinkTargetTypeFirstPage**|
| **pbHlinkTargetTypeLastPage**|
| **pbHlinkTargetTypeNextPage**|
| **pbHlinkTargetTypePageID**|
| **pbHlinkTargetTypePreviousPage**|
| **pbHlinkTargetTypeURL**|

## Example

The following example adds hyperlinks to shape one and shape two on page one of the active publication. The first hyperlink points to an external Web site, and the second link points to the fourth page in the publication. Shape one and shape two must be text boxes and there must be at least four pages in the publication for this example to work.


```vb
Dim hypNew As Hyperlink 
Dim lngPageID As Long 
Dim strPage As String 
 
With ActiveDocument.Pages(1).Shapes(1).TextFrame 
 Set hypNew = .TextRange.Hyperlinks.Add(Text:=.TextRange, _ 
 Address:="http://www.tailspintoys.com/", _ 
 TextToDisplay:="Tailspin") 
End With 
 
lngPageID = ActiveDocument.Pages(4).PageID 
strPage = "Go to page " _ 
 &; Str(ActiveDocument.Pages(4).PageNumber) 
 
With ActiveDocument.Pages(1).Shapes(2).TextFrame 
 Set hypNew = .TextRange.Hyperlinks.Add(Text:=.TextRange, _ 
 RelativePage:=pbHlinkTargetTypePageID, _ 
 PageID:=lngPageID, _ 
 TextToDisplay:=strPage) 
End With
```


