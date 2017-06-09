---
title: Predefined Bookmarks
ms.prod: word
ms.assetid: aa1c6d85-fe70-8f73-5682-ae6ada65be7c
ms.date: 06/08/2017
---


# Predefined Bookmarks

Word sets and automatically updates a number of reserved bookmarks. You can use these predefined bookmarks just as you use bookmarks that you place in documents, except that you do not have to set them and they are not listed on the  **Go To** tab in the **Find and Replace** dialog box.

You can use predefined bookmarks with the  **[Bookmarks](bookmarks-defaultsorting-property-word.md)** property. The following example sets the bookmark named "currpara" to the location marked by the predefined bookmark named "\Para".



```vb
ActiveDocument.Bookmarks("\Para").Copy "currpara"
```

The following table describes the predefined bookmarks available in Word.


|**Bookmark**|**Description**|
|:-----|:-----|
|\Sel|Current selection or the insertion point.|
|\PrevSel1|Most recent selection where editing occurred; going to this bookmark is equivalent to running the  **[GoBack](application-goback-method-word.md)** method once.|
|\PrevSel2|Second most recent selection where editing occurred; going to this bookmark is equivalent to running the  **GoBack** method twice.|
|\StartOfSel|Start of the current selection.|
|\EndOfSel|End of the current selection.|
|\Line|Current line or the first line of the current selection. If the insertion point is at the end of a line that is not the last line in the paragraph, the bookmark includes the entire next line.|
|\Char|Current character, which is the character following the insertion point if there is no selection, or the first character of the selection.|
|\Para|Current paragraph, which is the paragraph containing the insertion point or, if more than one paragraph is selected, the first paragraph of the selection. Note that if the insertion point or selection is in the last paragraph of the document, the "\Para" bookmark does not include the paragraph mark.|
|\Section|Current section, including the break at the end of the section, if any. The current section contains the insertion point or selection. If the selection contains more than one section, the "\Section" bookmark is the first section in the selection.|
|\Doc|Entire contents of the active document, with the exception of the final paragraph mark.|
|\Page|Current page, including the break at the end of the page, if any. The current page contains the insertion point. If the current selection contains more than one page, the "\Page" bookmark is the first page of the selection. Note that if the insertion point or selection is in the last page of the document, the "\Page" bookmark does not include the final paragraph mark.|
|\StartOfDoc|Beginning of the document.|
|\EndOfDoc|End of the document.|
|\Cell|Current cell in a table, which is the cell containing the insertion point. If one or more cells of a table are included in the current selection, the "\Cell" bookmark is the first cell in the selection.|
|\Table|Current table, which is the table containing the insertion point or selection. If the selection includes more than one table, the "\Table" bookmark is the entire first table of the selection, even if the entire table is not selected.|
|\HeadingLevel|The heading that contains the insertion point or selection, plus any subordinate headings and text. If the current selection is body text, the "\HeadingLevel" bookmark includes the preceding heading, plus any headings and text subordinate to that heading.|

