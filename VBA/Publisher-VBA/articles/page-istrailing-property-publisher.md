---
title: Page.IsTrailing Property (Publisher)
keywords: vbapb10.chm131101
f1_keywords:
- vbapb10.chm131101
ms.prod: publisher
api_name:
- Publisher.Page.IsTrailing
ms.assetid: e0ed15dc-d2e8-d6b7-913d-4e72b2817e88
ms.date: 06/08/2017
---


# Page.IsTrailing Property (Publisher)

 **True** if the specified **Page** object is a trailing page of a two-page spread. Read-only **Boolean**.


## Syntax

 _expression_. **IsTrailing**

 _expression_A variable that represents an  **Page** object.


### Return Value

Boolean


## Example

The following example diplays for each page whether the page is a trailing or leading page in the publication.


```vb
Dim objPage As Page 
Dim strPageInfo As String 
For Each objPage In ActiveDocument.Pages 
 strPageInfo = "Page number " &; objPage.PageNumber 
 If objPage.IsLeading Then 
 strPageInfo = strPageInfo &; " is a leading page." &; Chr(13) 
 ElseIf objPage.IsTrailing Then 
 strPageInfo = strPageInfo &; " is a trailing page." &; Chr(13) 
 End If 
 MsgBox strPageInfo 
Next objPage
```


