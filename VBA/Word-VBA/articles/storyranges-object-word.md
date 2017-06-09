---
title: StoryRanges Object (Word)
keywords: vbawd10.chm2444
f1_keywords:
- vbawd10.chm2444
ms.prod: word
ms.assetid: 131b04b0-e4a8-8969-0a4b-e5b3793af03d
ms.date: 06/08/2017
---


# StoryRanges Object (Word)

A collection of  **Range** objects that represent stories in a document.


## Remarks

Use the  **StoryRanges** property to return the **StoryRanges** collection. The following example removes manual character formatting from the text in all stories other than the main text story in the active document.


```
For Each aStory In ActiveDocument.StoryRanges 
 If aStory.StoryType <> wdMainTextStory Then aStory.Font.Reset 
Next aStory
```

The  **Add** method is not available for the **StoryRanges** collection. The number of stories in the **StoryRanges** collection is finite.

Use  **StoryRanges** (Index), where Index is a **WdStoryType** constant, to return a single story as a **[Range](range-object-word.md)** object. The following example adds text to the primary header story and then displays the text.




```
ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range _ 
 .Text = "Header text" 
MsgBox ActiveDocument.StoryRanges(wdPrimaryHeaderStory).Text
```

The following example copies the text of the footnotes from the active document into a new document.




```
If ActiveDocument.Footnotes.Count >= 1 Then 
 ActiveDocument.StoryRanges(wdFootnotesStory).Copy 
 Documents.Add.Content.Paste 
End If
```

If you attempt to return a story that is not available in the specified document, an error occurs. The following example determines whether a footnote story is available in the active document.




```
On Error GoTo errhandler 
Set MyRange = ActiveDocument.StoryRanges(wdFootnotesStory) 
errhandler: 
If Err = 5941 Then MsgBox "The footnotes story is not available."
```

Use the  **NextStoryRange** property to loop through all stories in a document. The following example searches each story in the active document for the text "Microsoft Word." When the text is found, it is formatted as italic.




```
For Each myStoryRange In ActiveDocument.StoryRanges 
 myStoryRange.Find.Execute _ 
 FindText:="Microsoft Word", Forward:=True 
 While myStoryRange.Find.Found 
 myStoryRange.Italic = True 
 myStoryRange.Find.Execute _ 
 FindText:="Microsoft Word", Forward:=True 
 Wend 
 While Not (myStoryRange.NextStoryRange Is Nothing) 
 Set myStoryRange = myStoryRange.NextStoryRange 
 myStoryRange.Find.Execute _ 
 FindText:="Microsoft Word", Forward:=True 
 While myStoryRange.Find.Found 
 myStoryRange.Italic = True 
 myStoryRange.Find.Execute _ 
 FindText:="Microsoft Word", Forward:=True 
 Wend 
 Wend 
Next myStoryRange
```


## Methods



|**Name**|
|:-----|
|[Item](storyranges-item-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](storyranges-application-property-word.md)|
|[Count](storyranges-count-property-word.md)|
|[Creator](storyranges-creator-property-word.md)|
|[Parent](storyranges-parent-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
