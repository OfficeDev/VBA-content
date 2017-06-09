---
title: Modifying a Portion of a Document
ms.prod: word
ms.assetid: e664871f-4499-421e-deb7-e064cdeba0f0
ms.date: 06/08/2017
---


# Modifying a Portion of a Document

Visual Basic includes objects that you can use to modify the following document elements: characters, words, sentences, paragraphs, and sections. The following table includes the properties that correspond to these document elements and the objects they return.



|**This expression**|**Returns this object**|
|:-----|:-----|
| **Words** ( _index_)| **[Range](range-object-word.md)**|
| **Characters** ( _index_)| **[Range](range-object-word.md)**|
| **Sentences** ( _index_)| **[Range](range-object-word.md)**|
| **Paragraphs** ( _index_)| **[Paragraph](paragraph-object-word.md)**|
| **Sections** ( _index_)| **[Section](section-object-word.md)**|

When these properties are used without an index, a collection object with the same name is returned. For example, the  **Paragraphs** property returns the **[Paragraphs](paragraphs-object-word.md)** collection object. However, if you identify an item within these collections by index, the object in the second column of the table is returned. For example, `Words(1)` returns a **Range** object. After you have a **Range** object, you can use any of the range properties or methods to modify the **Range** object. For example, the following instruction copies the first word in the selection to the Clipboard.




```vb
Sub CopyWord() 
    Selection.Words(1).Copy 
End Sub
```


 **Note**  The items in the  **[Paragraphs](paragraphs-object-word.md)** and **[Sections](sections-object-word.md)** collections are singular forms of the collection, specifically **Paragraph** objects and **Section** objects, rather than **Range** objects. In fact, most collections in the Word object model have singular form objects with which you can work. However, the **Range** property (which returns a **Range** object) is available from both the **Paragraph** object and the **Section** object, and from most other objects that are children of collections. For example, the following instruction copies the first paragraph in the active document to the Clipboard.




```vb
Sub CopyParagraph() 
    ActiveDocument.Paragraphs(1).Range.Copy 
End Sub
```

All of the document element properties in the preceding table are available from the  **Document**,  **Selection**, and  **Range** objects. The following examples demonstrate how you can drill down to these properties from **[Document](document-object-word.md)**,  **[Selection](selection-object-word.md)**, and  **Range** objects.
The following example sets the case of the first word in the active document.



```vb
Sub ChangeCase() 
    ActiveDocument.Words(1).Case = wdUpperCase 
End Sub
```

The following example sets the bottom margin of the current section to 0.5 inch.



```vb
Sub ChangeSectionMargin() 
    Selection.Sections(1).PageSetup.BottomMargin = InchesToPoints(0.5) 
End Sub
```

The following example double spaces the text in the active document (the  **[Content](document-content-property-word.md)** property returns a **Range** object).



```vb
Sub DoubleSpaceDocument() 
    ActiveDocument.Content.ParagraphFormat.Space2 
End Sub
```


## Modifying a group of document elements

To modify a range of text that consists of a group of document elements (characters, words, sentences, paragraphs, or sections), you need to create a  **Range** object. The **Range** method creates a **Range** object given a start and endpoint. For example, the following instruction creates a **Range** object that refers to the first ten characters in the active document.


```vb
Sub SetRangeForFirstTenCharacters() 
    Dim rngTenCharacters As Range 
    Set rngTenCharacters = ActiveDocument.Range(Start:=0, End:=10) 
End Sub
```

Using the  **[Start](range-start-property-word.md)** and **[End](range-end-property-word.md)** properties with a **Range** object, you can create a new **Range** object that refers to a group of document elements. For example, the following instruction creates a **Range** object ( `rngThreeWords`) that refers to the first three words in the active document.




```vb
Sub SetRangeForFirstThreeWords() 
    Dim docActive As Document 
    Dim rngThreeWords As Range 
    Set docActive = ActiveDocument 
    Set rngThreeWords = docActive.Range(Start:=docActive.Words(1).Start, _ 
        End:=docActive.Words(3).End) 
End Sub
```

The following example creates a  **Range** object ( `rngParagraphs`) beginning at the start of the second paragraph and ending after the third paragraph.




```vb
Sub SetParagraphRange() 
    Dim docActive As Document 
    Dim rngParagraphs As Range 
    Set docActive = ActiveDocument 
    Set rngParagraphs = docActive.Range(Start:=docActive.Paragraphs(2).Range.Start, _ 
        End:=docActive.Paragraphs(3).Range.End) 
End Sub
```

For more information about defining  **Range** objects, see [Working with Range objects](working-with-range-objects.md).


