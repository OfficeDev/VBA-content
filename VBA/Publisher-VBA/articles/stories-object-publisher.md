---
title: Stories Object (Publisher)
keywords: vbapb10.chm5767167
f1_keywords:
- vbapb10.chm5767167
ms.prod: publisher
api_name:
- Publisher.Stories
ms.assetid: 694a0376-fa41-3097-180b-40b8a005ddf6
ms.date: 06/08/2017
---


# Stories Object (Publisher)

Represents all the text in a publication.
 


## Example

Use the  **Stories** property of a **Document** object to return a **Stories** collection. Use the **Item** method of the **Stories** collection to access individual **Story** objects.
 

 

 

 
The  **Stories** collection enables efficient access to text in a publication. A simple loop through the **Stories** collection can scan all text in text frames or tables without the need to search each shape on every page of a publication.
 

 

 

 
The  **Stories** collection contains one **Story** object for each unlinked text frame, each chain of linked text frames, and each table in a publication. Text in WordArt frames, OLE objects, and pictures are not included in the **Stories** collection.
 

 

 

 
This example assigns the first story in the active publication to an object variable.
 

 



```
Dim stFirst As Story 
 
stFirst = Application.ActiveDocument.Stories(1)
```


## Methods



|**Name**|
|:-----|
|[Item](stories-item-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](stories-application-property-publisher.md)|
|[Count](stories-count-property-publisher.md)|
|[Parent](stories-parent-property-publisher.md)|

