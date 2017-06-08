---
title: Bookmark Object (Word)
keywords: vbawd10.chm2408
f1_keywords:
- vbawd10.chm2408
ms.prod: word
api_name:
- Word.Bookmark
ms.assetid: be6b0c7b-60ca-97e7-ef19-6de335da3197
ms.date: 06/08/2017
---


# Bookmark Object (Word)

Represents a single bookmark in a document, selection, or range. The  **Bookmark** object is a member of the **[Bookmarks](bookmarks-object-word.md)** collection. The **Bookmarks** collection includes all the bookmarks listed in the **Bookmark** dialog box ( **Insert** menu).


## Remarks

Using the Bookmark Object

Use  **Bookmarks** (index), where index is the bookmark name or index number, to return a single **Bookmark** object. You must exactly match the spelling (but not necessarily the capitalization) of the bookmark name. The following example selects the bookmark named "temp" in the active document.




```
ActiveDocument.Bookmarks("temp").Select
```

The index number represents the position of the bookmark in the  **[Selection](selection-object-word.md)** or **[Range](range-object-word.md)** object. For the **[Document](document-object-word.md)** object, the index number represents the position of the bookmark in the alphabetical list of bookmarks in the **Bookmarks** dialog box (click **Name** to sort the list of bookmarks alphabetically). The following example displays the name of the second bookmark in the **Bookmarks** collection.




```
MsgBox ActiveDocument.Bookmarks(2).Name
```

Use the  **[Add](bookmarks-add-method-word.md)** method to add a bookmark to a document range. The following example marks the selection by adding a bookmark named "temp."




```
ActiveDocument.Bookmarks.Add Name:="temp", Range:=Selection.Range
```

Remarks

Use the  **BookmarkID** property with a range or selection object to return the index number of a **Bookmark** object in the **Bookmarks** collection. The following example displays the index number of the bookmark named "temp" in the active document.




```
MsgBox ActiveDocument.Bookmarks("temp").Range.BookmarkID
```

You can use [predefined bookmarks](http://msdn.microsoft.com/library/aa1c6d85-fe70-8f73-5682-ae6ada65be7c%28Office.15%29.aspx)with the  **Bookmarks** property. The following example sets the bookmark named "currpara" to the location marked by the predefined bookmark named "\Para".




```
ActiveDocument.Bookmarks("\Para").Copy "currpara"
```

Use the  **[Exists](bookmarks-exists-method-word.md)** method to determine whether a bookmark already exists in the selection, range, or document. The following example ensures that the bookmark named "temp" exists in the active document before selecting the bookmark.




```
If ActiveDocument.Bookmarks.Exists("temp") = True Then 
 ActiveDocument.Bookmarks("temp").Select 
End If
```


## Methods



|**Name**|
|:-----|
|[Copy](bookmark-copy-method-word.md)|
|[Delete](bookmark-delete-method-word.md)|
|[Select](bookmark-select-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](bookmark-application-property-word.md)|
|[Column](bookmark-column-property-word.md)|
|[Creator](bookmark-creator-property-word.md)|
|[Empty](bookmark-empty-property-word.md)|
|[End](bookmark-end-property-word.md)|
|[Name](bookmark-name-property-word.md)|
|[Parent](bookmark-parent-property-word.md)|
|[Range](bookmark-range-property-word.md)|
|[Start](bookmark-start-property-word.md)|
|[StoryType](bookmark-storytype-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
