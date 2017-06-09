---
title: Range.NextStoryRange Property (Word)
keywords: vbawd10.chm157155648
f1_keywords:
- vbawd10.chm157155648
ms.prod: word
api_name:
- Word.Range.NextStoryRange
ms.assetid: 392b17ff-335f-9b2b-7641-62ae44d7e919
ms.date: 06/08/2017
---


# Range.NextStoryRange Property (Word)

Returns a  **Range** object that refers to the next story. Read-only **Range** .


## Syntax

 _expression_ . **NextStoryRange**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

The following table describes the range returned depending on the type of story.



|**Story type**|**Item returned by the NextStoryRange method**|
|:-----|:-----|
| **wdMainTextStory** , **wdFootnotesStory** , **wdEndnotesStory** , and **wdCommentsStory**|Always returns  **Nothing**|
| **wdTextFrameStory**|The story of the next set of linked text boxes|
| **wdEvenPagesHeaderStory** , **wdPrimaryHeaderStory** , **wdEvenPagesFooterStory** , **wdPrimaryFooterStory** , **wdFirstPageHeaderStory** , **wdFirstPageFooterStory**|The next section's story of the same type|

## See also


#### Concepts


[Range Object](range-object-word.md)

