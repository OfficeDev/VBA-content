---
title: Sections Object (Publisher)
keywords: vbapb10.chm7405567
f1_keywords:
- vbapb10.chm7405567
ms.prod: publisher
api_name:
- Publisher.Sections
ms.assetid: 429c03b8-b574-86db-c39d-551a4c753b04
ms.date: 06/08/2017
---


# Sections Object (Publisher)

A collection of all the  **Section** objects in the document.
 


## Example

Use  **Sections**.Item(index) where index is the index number, to return a single **Section** object. The following example sets the number format and the starting number for the first section of the active document.
 

 

```
With ActiveDocument.Sections.Item(1) 
 .PageNumberFormat = pbPageNumberFormatArabic 
 .PageNumberStart = 1 
End With
```

Using  **Sections** (index) where index is the index number, will also return a single **Section** object. The following example sets continues the numbering from the previous section for the second section in the active document.
 

 



```
ActiveDocument.Sections(2).ContinueNumbersFromPreviousSection=True
```

Use  **Sections**.Count to return the number of sections in the publication. The following example display the number of sections in the first open document.
 

 



```
MsgBox Documents(1).Sections.Count
```

Use  **Sections**.Add(StartPageIndex) where StartPageIndex is the index number of the page, to reutrn a new section added to a document. A "Permission denied." error will be returned if the page already contains a section head. The following example adds a new section to the second page of the active document.
 

 



```
Dim objSection As Section 
Set objSection = ActiveDocument.Sections.Add(StartPageIndex:=2)
```

Use  **Sections** (index).Delete where index is the index number, to delete the specified section from the document. A "Permission denied" error will be returned if an attempt is made to delete the first section. The following example deletes all of the sections of the active document except the first one.
 

 

 **Note**  The iteration is from the last to the first to avoid a "Subscript out of range." error when accessing a deleted section in the  **Sections** collection.
 




```
Dim i As Long 
For i = ActiveDocument.Sections.Count To 1 Step -1 
 If i = 1 Then Exit For 
 ActiveDocument.Sections(i).Delete 
Next i
```


## Methods



|**Name**|
|:-----|
|[Add](sections-add-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](sections-application-property-publisher.md)|
|[Count](sections-count-property-publisher.md)|
|[Item](sections-item-property-publisher.md)|
|[Parent](sections-parent-property-publisher.md)|

