---
title: Replacement Object (Word)
keywords: vbawd10.chm2481
f1_keywords:
- vbawd10.chm2481
ms.prod: word
api_name:
- Word.Replacement
ms.assetid: 5d9615e4-f6ef-af5f-6e45-c382a88395c9
ms.date: 06/08/2017
---


# Replacement Object (Word)

Represents the replace criteria for a find-and-replace operation. The properties and methods of the  **Replacement** object correspond to the options in the **Find and Replace** dialog box.


## Remarks

Use the  **Replacement** property to return a **Replacement** object. The following example replaces the next occurrence of the word "hi" with the word "hello."


```
With Selection.Find 
 .Text = "hi" 
 .ClearFormatting 
 .Replacement.Text = "hello" 
 .Replacement.ClearFormatting 
 .Execute Replace:=wdReplaceOne, Forward:=True 
End With
```

To find and replace formatting, set both the find text and the replace text to empty strings ("") and set the Format argument of the  **Execute** method to **True**. The following example removes all the bold formatting in the active document. The **Bold** property is **True** for the **Find** object and **False** for the **Replacement** object.




```
With ActiveDocument.Content.Find 
 .ClearFormatting 
 .Font.Bold = True 
 .Text = "" 
 With .Replacement 
 .ClearFormatting 
 .Font.Bold = False 
 .Text = "" 
 End With 
 .Execute Format:=True, Replace:=wdReplaceAll 
End With
```


## Methods



|**Name**|
|:-----|
|[ClearFormatting](replacement-clearformatting-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](replacement-application-property-word.md)|
|[Creator](replacement-creator-property-word.md)|
|[Font](replacement-font-property-word.md)|
|[Frame](replacement-frame-property-word.md)|
|[Highlight](replacement-highlight-property-word.md)|
|[LanguageID](replacement-languageid-property-word.md)|
|[LanguageIDFarEast](replacement-languageidfareast-property-word.md)|
|[NoProofing](replacement-noproofing-property-word.md)|
|[ParagraphFormat](replacement-paragraphformat-property-word.md)|
|[Parent](replacement-parent-property-word.md)|
|[Style](replacement-style-property-word.md)|
|[Text](replacement-text-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
