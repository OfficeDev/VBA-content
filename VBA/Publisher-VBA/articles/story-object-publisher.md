---
title: Story Object (Publisher)
keywords: vbapb10.chm5898239
f1_keywords:
- vbapb10.chm5898239
ms.prod: publisher
api_name:
- Publisher.Story
ms.assetid: 0385b4be-0046-9198-a186-0d992601780e
ms.date: 06/08/2017
---


# Story Object (Publisher)

Represents the text in an unlinked text frame, text flowing between linked text frames, or text in a table cell. The  **Story** object is a member of the **TextFrame** and **TextRange** objects and the **Stories** collection.


## Example

Use the  **Story** property to return the **Story** object in a text range or text frame. This example returns the story in the selected text range and, if it is in a text frame, inserts text into the text range.


```
Sub AddTextToStory() 
 With Selection.TextRange.Story 
 If .HasTextFrame Then .TextRange _ 
 .InsertAfter NewText:=vbLf &amp; "This is a test." 
 End With 
End Sub
```

Use  **Stories** (index), where index is the number of the story, to return an individual **Story** object. This example determines if the first story in the active publication has a text frame and, if it does, formats the paragraphs in the story with a half inch first line indent and a six-point spacing before each paragraph.




```
Sub StoryParagraphFirstLineIndent() 
 With ActiveDocument.Stories(1) 
 If .HasTextFrame Then 
 With .TextFrame.TextRange.ParagraphFormat 
 .FirstLineIndent = InchesToPoints(0.5) 
 .SpaceBefore = 6 
 End With 
 End If 
 End With 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/26c38a3a-e30b-1f2d-d535-57bb978bc4f7%28Office.15%29.aspx)|
|[HasTable](http://msdn.microsoft.com/library/bc4912e2-f521-c6b5-b5a6-a49952014966%28Office.15%29.aspx)|
|[HasTextFrame](http://msdn.microsoft.com/library/10c3a002-05ae-1167-784c-d62066de802d%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/fbcc74f6-a7ba-df22-0b75-a7b365883d89%28Office.15%29.aspx)|
|[Table](http://msdn.microsoft.com/library/e9da80d3-ea3c-b47c-d434-498c72955c14%28Office.15%29.aspx)|
|[TextFrame](http://msdn.microsoft.com/library/bb6ce510-068c-27c2-9df0-a709ab46db2e%28Office.15%29.aspx)|
|[TextRange](http://msdn.microsoft.com/library/c948da79-ea67-0c8c-1df3-2b32499ea9b3%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/71e6548d-f54a-b4df-d878-d86a85c1332b%28Office.15%29.aspx)|

