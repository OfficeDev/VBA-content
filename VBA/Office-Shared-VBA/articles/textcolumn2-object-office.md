---
title: TextColumn2 Object (Office)
ms.prod: office
api_name:
- Office.TextColumn2
ms.assetid: 631387c1-2b7a-6c98-d05f-c054434c8b9d
ms.date: 06/08/2017
---


# TextColumn2 Object (Office)

Represents a single text column. The  **TextColumn2** object is a member of the **TextColumns2** collection.


## Remarks

Use  **TextColumns2(Index)**, where _Index_ is the index number, to return a single **TextColumn2** object. The index number represents the position of the column in the **TextColumns2** collection (counting from left to right).


## Example

Use the  **Add** method to add a column to the collection of columns. By default, there's one text column in the **TextColumns2** collection. The following example adds a 2.5-inch-widecolumn to the active Microsoft Word document.


```
ActiveDocument.PageSetup.TextColumns2.Add _ 
 Width:=InchesToPoints(2.5), _ 
 Spacing:=InchesToPoints(0.5), EvenlySpaced:=False 

```


## Properties



|**Name**|
|:-----|
|[Application](textcolumn2-application-property-office.md)|
|[Creator](textcolumn2-creator-property-office.md)|
|[Number](textcolumn2-number-property-office.md)|
|[Spacing](textcolumn2-spacing-property-office.md)|
|[TextDirection](textcolumn2-textdirection-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
