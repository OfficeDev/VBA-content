---
title: Working with Range Objects
ms.prod: word
ms.assetid: 9e240aa7-8608-9d70-aee3-2e202687459e
ms.date: 06/08/2017
---


# Working with Range Objects

A common task when using Visual Basic is to specify an area in a document and then do something with it, such as insert text or apply formatting. For example, you may want to write a macro that locates a word or phrase within a portion of a document. The portion of the document can be represented by a  **[Range](range-object-word.md)** object. After the  **Range** object is identified, methods and properties of the **Range** object can be applied to modify the contents of the range.

A  **Range** object refers to a contiguous area in a document. Each **Range** object is defined by a starting and ending character position. Similar to the way bookmarks are used in a document, **Range** objects are used in Visual Basic procedures to identify specific portions of a document. A **Range** object can be as small as the insertion point or as large as the entire document. However, unlike a bookmark, a **Range** object exists only while the procedure that defined it is running.

The  **[Start](range-start-property-word.md)**,  **[End](range-end-property-word.md)**, and  **[StoryType](range-storytype-property-word.md)** properties uniquely identify a  **Range** object. The **Start** and **End** properties return or set the starting and ending character positions of the **Range** object. The character position at the beginning of the document is zero, the position after the first character is one, and so on. There are 11 different story types represented by the **WdStoryType** constants of the **StoryType** property.


 **Note**   **Range** objects are independent of the selection. That is, you can define and modify a range without changing the current selection. You can also define multiple ranges in a document, but there is only one selection per document pane.


## Using the Range method

Use the  **[Range](range-checksynonyms-method-word.md)** method of the  **[Document](document-object-word.md)** object to create a **Range** object that is located in the main story and has a given start and endpoint. The following example creates a **Range** object that starts at the beginning of the first character and extends through the tenth character.


```vb
Sub SetNewRange() 
 Dim rngDoc As Range 
 Set rngDoc = ActiveDocument.Range(Start:=0, End:=10) 
End Sub
```

You can see that the  **Range** object is created when you apply a property or method to the **Range** object. For example, the following applies bold formatting to the first 10 characters in the active document.




```vb
Sub SetBoldRange() 
 Dim rngDoc As Range 
 Set rngDoc = ActiveDocument.Range(Start:=0, End:=10) 
 rngDoc.Bold = True 
End Sub
```

When you need to refer to a  **Range** object multiple times, you can use the **Set** statement to set a variable equal to the **Range** object. However, if you only need to perform a single action on a **Range** object, you do not need to store the object in a variable. The same result can be achieved using just one instruction that identifies the range and changes the **[Bold](range-bold-property-word.md)** property.




```vb
Sub BoldRange() 
 ActiveDocument.Range(Start:=0, End:=10).Bold = True 
End Sub
```

Like a bookmark, a range can span a group of characters or mark a location in a document. The  **Range** object in the following example has the same starting and ending points. The range does not include any text. The following example inserts text at the beginning of the active document.




```vb
Sub InsertTextBeforeRange() 
 Dim rngDoc As Range 
 Set rngDoc = ActiveDocument.Range(Start:=0, End:=0) 
 rngDoc.InsertBefore "Hello " 
End Sub
```

You can define the beginning and endpoints of a range using the character position numbers, as shown above, or use the  **Start** and **End** properties with objects such as **[Selection](selection-object-word.md)**,  **[Bookmark](bookmark-object-word.md)**, or  **Range** objects. The following example creates a **Range** object beginning at the start of the second paragraph and ending after the third paragraph.




```vb
Sub NewRange() 
 Dim doc As Document 
 Dim rngDoc As Range 
 
 Set doc = ActiveDocument 
 Set rngDoc = doc.Range(Start:=doc.Paragraphs(2).Range.Start, _ 
 End:=doc.Paragraphs(3).Range.End) 
End Sub
```

For additional information and examples, see the  **[Range](range-checksynonyms-method-word.md)** method.


## Using the Range property

The  **Range** property appears on multiple objects—such as **[Paragraph](paragraph-object-word.md)**,  **[Bookmark](bookmark-object-word.md)**, and  ** [Cell](cell-object-word.md)**—and is used to return a  **Range** object. The following example returns a **Range** object that refers to the first paragraph in the active document.


```vb
Sub SetParagraphRange() 
 Dim rngParagraph As Range 
 Set rngParagraph = ActiveDocument.Paragraphs(1).Range 
End Sub
```

After you have a  **Range** object, you can use any of its properties or methods to modify the **Range** object. The following example selects the second paragraph in the active document and then centers the selection.




```vb
Sub FormatRange() 
 ActiveDocument.Paragraphs(2).Range.Select 
 Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter 
End Sub
```

If you need to apply numerous properties or methods to the same  **Range** object, you can use the **With…End With** structure. The following example formats the text in the first paragraph of the active document.




```vb
Sub FormatFirstParagraph() 
 Dim rngParagraph As Range 
 Set rngParagraph = ActiveDocument.Paragraphs(1).Range 
 With rngParagraph 
 .Bold = True 
 .ParagraphFormat.Alignment = wdAlignParagraphCenter 
 With .Font 
 .Name = "Stencil" 
 .Size = 15 
 End With 
 End With 
End Sub
```

For additional information and examples, see the  **[Range](range-case-property-word.md)** property topic.


## Redefining a Range object

Use the  **[SetRange](range-setrange-method-word.md)** method to redefine an existing  **Range** object. The following example defines a range as the current selection. The **SetRange** method then redefines the range so that it refers to the current selection plus the next 10 characters.


```vb
Sub ExpandRange() 
 Dim rngParagraph As Range 
 Set rngParagraph = Selection.Range 
 rngParagraph.SetRange Start:=rngParagraph.Start, _ 
 End:=rngParagraph.End + 10 
End Sub
```

For additional information and examples, see the  **[Range](range-checksynonyms-method-word.md)** method for the  **[Document](document-object-word.md)**.


 **Note**  When debugging your macros, you can use the  **Select**method to ensure that a  **Range** object is referring to the correct range of text. For example, the following selects a **Range** object that refers to the second and third paragraphs in the active document, and then formats the font of the selection.




```vb
Sub SelectRange() 
 Dim rngParagraph As Range 
 
 Set rngParagraph = ActiveDocument.Paragraphs(2).Range 
 
 rngParagraph.SetRange Start:=rngParagraph.Start, _ 
 End:=ActiveDocument.Paragraphs(3).Range.End 
 rngParagraph.Select 
 
 Selection.Font.Italic = True 
End Sub
```


