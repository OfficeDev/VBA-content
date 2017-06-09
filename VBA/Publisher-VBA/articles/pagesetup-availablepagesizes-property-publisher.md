---
title: PageSetup.AvailablePageSizes Property (Publisher)
keywords: vbapb10.chm6946849
f1_keywords:
- vbapb10.chm6946849
ms.prod: publisher
api_name:
- Publisher.PageSetup.AvailablePageSizes
ms.assetid: 5ad79ee6-6d32-6c46-c02e-a9ab252267cf
ms.date: 06/08/2017
---


# PageSetup.AvailablePageSizes Property (Publisher)

Returns the  **PageSizes** collection that contains all the **[PageSize](pagesize-object-publisher.md)** objects available in the current publication.


## Syntax

 _expression_. **AvailablePageSizes**

 _expression_A variable that represents a  **PageSetup** object.


### Return Value

PageSizes


## Remarks

 **PageSize** objects correspond to the icons displayed under **Blank Page Sizes** in the **Page Setup** dialog box in the Microsoft Publisher user interface.

Page sizes returned by the  **AvailablePageSizes** property include not only the page sizes provided by Microsoft Publisher, but also custom page sizes that you create or download, if any.

As you add or remove custom page sizes, the index number for all existing page sizes may change. 


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to create a text file that contains the list of all page sizes available in the current publication and their corresponding index numbers. It saves the text file to the Documents (in Windows Vista) or My Documents (in Windows XP) folder of the current user.


```vb
Public Sub AvailablePageSizes_Example() 
 
 Dim pubPageSize As Publisher.PageSize 
 Dim pubPageSizes As Publisher.PageSizes 
 Dim intCount As Integer 
 Dim lngPageSizeFile As Long 
 
 intCount = 1 
 
 Set pubPageSizes = ThisDocument.PageSetup.AvailablePageSizes 
 
 lngPageSizeFile = FreeFile 
 Open Environ("USERPROFILE") + "\My Documents\PageSizeList.txt" For Output Access Write As lngPageSizeFile 
 
 For Each pubPageSize In pubPageSizes 
 Write #lngPageSizeFile, pubPageSize.Name, intCount 
 intCount = intCount + 1 
 Next 
 
 Close lngPageSizeFile 
 
End Sub
```


