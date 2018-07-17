---
title: PageSizes Object (Publisher)
keywords: vbapb10.chm8847359
f1_keywords:
- vbapb10.chm8847359
ms.prod: publisher
api_name:
- Publisher.PageSizes
ms.assetid: f31b08cc-2c76-e2d6-d1ae-6dcf2ac5824c
ms.date: 06/08/2017
---


# PageSizes Object (Publisher)

Represents the collection of all  **PageSize** objects in the parent **Document** object, where each **PageSize** object represents one of the page sizes available in the current Microsoft Publisher document.


## Remarks

The page sizes represented by the  **PageSizes** collection correspond to the icons displayed under **Blank Page Sizes** in the **Page Setup** dialog box in the Publisher user interface.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **PageSizes** collection to get all the page sizes available in the current document and print the list in the **Immediate** window.


```
Public Sub PageSizes_Example() 
 
 Dim pubPageSizes As Publisher.PageSizes 
 Dim pubPageSize As Publisher.PageSize 
 
 Set pubPageSizes = ThisDocument.PageSetup.AvailablePageSizes 
 For Each pubPageSize In pubPageSizes 
 Debug.Print pubPageSize.Name 
 Next 
 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/bce8ec2c-1a05-1e0b-8d75-7e4dd7084a19%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/10770705-e8b3-903c-bcfa-84ba26dc1478%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/7fc17907-7e3b-8415-23dc-7f1a64db731c%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/622d2bee-a7b7-6f5f-cb7c-39d69f432b27%28Office.15%29.aspx)|

