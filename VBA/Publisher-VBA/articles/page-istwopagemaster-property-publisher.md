---
title: Page.IsTwoPageMaster Property (Publisher)
keywords: vbapb10.chm131100
f1_keywords:
- vbapb10.chm131100
ms.prod: publisher
api_name:
- Publisher.Page.IsTwoPageMaster
ms.assetid: dbfc3c21-0070-3f0a-c0b0-746d83c46765
ms.date: 06/08/2017
---


# Page.IsTwoPageMaster Property (Publisher)

 **True** if the specified **Page** object is a two-page master. Read/write **Boolean**.


## Syntax

 _expression_. **IsTwoPageMaster**

 _expression_A variable that represents an  **Page** object.


### Return Value

Boolean


## Remarks

This method works for master pages only. Returns a  **This feature is only for master pages** error when attempting to access this property from a publication page object.


## Example

The following example adds text to each header of a two-page master page specifying the master page PageNumber and its place in the spread: 1 or 2.


```vb
Dim objMasterPage As Page 
Dim pageCount As Long 
Dim i As Long 
pageCount = ActiveDocument.MasterPages.Count 
For i = 1 To pageCount 
 Set objMasterPage = ActiveDocument.MasterPages(i) 
 If objMasterPage.IsTwoPageMaster Then 
 objMasterPage.Header.TextRange.Text = "MasterPage " &; _ 
 objMasterPage.PageNumber &; ", Page 1 of 2" 
 i = i + 1 
 Set objMasterPage = ActiveDocument.MasterPages(i) 
 objMasterPage.Header.TextRange.Text = "MasterPage " &; _ 
 objMasterPage.PageNumber &; ", Page 2 of 2" 
 End If 
Next i 

```


