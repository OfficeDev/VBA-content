---
title: Report.PicturePages Property (Access)
keywords: vbaac10.chm13709,vbaac10.chm4464
f1_keywords:
- vbaac10.chm13709,vbaac10.chm4464
ms.prod: access
api_name:
- Access.Report.PicturePages
ms.assetid: a1266a43-3e1c-33f3-ae18-a7306723cc11
ms.date: 06/08/2017
---


# Report.PicturePages Property (Access)

You can use the  **PicturePages** property to specify on which page or pages of a report a picture will be displayed. Read/write **Byte**.


## Syntax

 _expression_. **PicturePages**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **PicturePages** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|All Pages|0|(Default) The picture appears on all pages of the report.|
|First Page|1|The picture appears only on the first page of the report.|
|No Pages|2|The picture doesn't appear on the report.|

## Example

The following example prints a stretched version of the picture "Logo.gif" on only the first page of the "Purchase Order" report.


```vb
With Reports("Purchase Order") 
 .Picture = "C:\Picture Files\Logo.gif" 
 .PictureSizeMode = 1 
 .PicturePages = 1 
End With
```


## See also


#### Concepts


[Report Object](report-object-access.md)

