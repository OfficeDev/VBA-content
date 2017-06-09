---
title: TextRange Object (Publisher)
keywords: vbapb10.chm5373951
f1_keywords:
- vbapb10.chm5373951
ms.prod: publisher
api_name:
- Publisher.TextRange
ms.assetid: 566f240b-d2a6-8cb3-9eb7-68328d6c28bd
ms.date: 06/08/2017
---


# TextRange Object (Publisher)

Contains the text that is attached to a shape, in addition to properties and methods for manipulating the text. This topic describes how to: 


- Return the text range in any shape you specify.
    
- Return a text range from the selection.
    
- Return particular characters, words, lines, sentences, or paragraphs from a text range.
    
- Insert text, the date and time, or the page number into a text range.
    

## Example

Use the  **[TextRange](http://msdn.microsoft.com/library/44a8395e-81dc-7d06-f068-89f77a889f5e%28Office.15%29.aspx)** property of the **[TextFrame](textframe-object-publisher.md)** object to return a **TextRange** object for any shape you specify. Use the **[Text](http://msdn.microsoft.com/library/13584812-307a-c32b-ca8f-27869728b64e%28Office.15%29.aspx)** property to return the string of text in the **TextRange** object. The following example adds a rectangle to the active publication and sets the text it contains.


```
Sub AddTextToShape() 
    With ActiveDocument.Pages(1).Shapes.AddShape(Type:=msoShapeRectangle, _ 
        Left:=72, Top:=72, Width:=250, Height:=140) 
        .TextFrame.TextRange.Text = "Here is some test text" 
    End With 
End Sub
```

Because the  **Text** property is the default property of the **TextRange** object, the following two statements are equivalent.




```
ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
    .TextRange.text = "Here is some test text" 
ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
    .TextRange = "Here is some test text"
```

Use the  **[HasTextFrame](http://msdn.microsoft.com/library/8a3b4f3b-3282-686b-f4fe-abf2d7677b3e%28Office.15%29.aspx)** property to determine whether a shape has a text frame, and use the **[HasText](http://msdn.microsoft.com/library/f8d1c660-c3f1-e835-adc3-114e6611de98%28Office.15%29.aspx)** property to determine whether the text frame contains text.

Use the  **TextRange** property of the **Selection** object to return the currently selected text. The following example copies the selection to the Clipboard.




```
Sub CopyAndPasteText() 
    With ActiveDocument 
        .Selection.TextRange.Copy 
        .Pages(1).Shapes(1).TextFrame.TextRange.Paste 
    End With 
End Sub
```

Use one of the following methods to return a portion of the text of a  **TextRange** object: **[Characters](http://msdn.microsoft.com/library/e851767e-12b2-ad77-071b-9d27bbf0d637%28Office.15%29.aspx)**, **[Lines](http://msdn.microsoft.com/library/56862090-b2ff-403b-d016-e37108d5ccc1%28Office.15%29.aspx)**, **[Paragraphs](http://msdn.microsoft.com/library/895c32cf-cdbe-74b0-ab47-6ae63d1bdea0%28Office.15%29.aspx)**, or **[Words](http://msdn.microsoft.com/library/df812db2-98ca-848b-7922-6905cb71124c%28Office.15%29.aspx)**. The following example formats the second word in the first shape on the first page of the active publication. For this example to work, the specified shape must contain text.




```
Sub FormatWords() 
    With ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
            .TextRange.Words(2).Font 
        .Bold = msoTrue 
        .Size = 15 
        .Name = "Text Name" 
    End With 
End Sub
```

Use one of the following methods to insert characters into a  **TextRange** object: **[InsertAfter](http://msdn.microsoft.com/library/f647be29-68c7-b221-adf1-fa233583e74e%28Office.15%29.aspx)**, **[InsertBefore](http://msdn.microsoft.com/library/b0e4355b-b1bc-ae78-08ad-000d577fd7db%28Office.15%29.aspx)**, **[InsertDateTime](http://msdn.microsoft.com/library/1d02471a-f22b-7dad-bcbb-40af3a04d198%28Office.15%29.aspx)**, **[InsertPageNumber](http://msdn.microsoft.com/library/f71d3b40-0263-93fa-d7e3-d815b90f71f7%28Office.15%29.aspx)**, or **[InsertSymbol](http://msdn.microsoft.com/library/607d12da-5a2d-4e0e-b45e-92275ce97bab%28Office.15%29.aspx)**. This example inserts a new line with text after any existing text in the first shape on the first page of the active publication.




```
Sub InsertNewText() 
    Dim intCount As Integer 
    With ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
            .TextRange 
        For intCount = 1 To 3 
            .InsertAfter vbLf &amp; "This is a test." 
        Next intCount 
    End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Characters](http://msdn.microsoft.com/library/e851767e-12b2-ad77-071b-9d27bbf0d637%28Office.15%29.aspx)|
|[Collapse](http://msdn.microsoft.com/library/ae177297-bf3b-ce0f-cf3a-29093b115996%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/e0d92492-fa0e-9424-471d-09866402702c%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/c9b8b896-26e7-ac58-0e1a-a66ef789f397%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/3062b5ea-fdb7-6632-0838-02e2c9c1c906%28Office.15%29.aspx)|
|[Expand](http://msdn.microsoft.com/library/66d8b1a3-5fc4-bed7-94d2-06be6203e1e9%28Office.15%29.aspx)|
|[InsertAfter](http://msdn.microsoft.com/library/f647be29-68c7-b221-adf1-fa233583e74e%28Office.15%29.aspx)|
|[InsertBarcode](http://msdn.microsoft.com/library/ad613ca7-f056-55b0-1a96-51167555ce6f%28Office.15%29.aspx)|
|[InsertBefore](http://msdn.microsoft.com/library/b0e4355b-b1bc-ae78-08ad-000d577fd7db%28Office.15%29.aspx)|
|[InsertDateTime](http://msdn.microsoft.com/library/1d02471a-f22b-7dad-bcbb-40af3a04d198%28Office.15%29.aspx)|
|[InsertMailMergeField](http://msdn.microsoft.com/library/97bce07d-b831-3ad6-2436-f85590c3bcd8%28Office.15%29.aspx)|
|[InsertPageNumber](http://msdn.microsoft.com/library/f71d3b40-0263-93fa-d7e3-d815b90f71f7%28Office.15%29.aspx)|
|[InsertSymbol](http://msdn.microsoft.com/library/607d12da-5a2d-4e0e-b45e-92275ce97bab%28Office.15%29.aspx)|
|[Lines](http://msdn.microsoft.com/library/56862090-b2ff-403b-d016-e37108d5ccc1%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/a51b4153-2ac5-2293-d2a0-d4a3786268d7%28Office.15%29.aspx)|
|[MoveEnd](http://msdn.microsoft.com/library/4fe27375-34e2-2ecc-33c8-a07230012b13%28Office.15%29.aspx)|
|[MoveStart](http://msdn.microsoft.com/library/5a9c480b-3cb7-0fd8-59c0-e2f93a925164%28Office.15%29.aspx)|
|[Paragraphs](http://msdn.microsoft.com/library/895c32cf-cdbe-74b0-ab47-6ae63d1bdea0%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/dd29c9ab-7f56-3604-3390-8f5a3b97821f%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/36097502-2b06-37ac-3148-43a82cca4411%28Office.15%29.aspx)|
|[Words](http://msdn.microsoft.com/library/df812db2-98ca-848b-7922-6905cb71124c%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/8c39c632-7c5b-6057-c4f7-2003b59b4644%28Office.15%29.aspx)|
|[BoundHeight](http://msdn.microsoft.com/library/010d3de9-5838-fbf7-fb75-b80a06aafac8%28Office.15%29.aspx)|
|[BoundLeft](http://msdn.microsoft.com/library/1ad36906-3dbf-9158-173b-b9047910f6d2%28Office.15%29.aspx)|
|[BoundTop](http://msdn.microsoft.com/library/f3c2cd42-8d2b-f757-bcbb-140f5e567a1e%28Office.15%29.aspx)|
|[BoundWidth](http://msdn.microsoft.com/library/bab5053f-958b-9264-9a1e-6f81b5a860b7%28Office.15%29.aspx)|
|[ContainingObject](http://msdn.microsoft.com/library/f15c81b5-d03f-0d83-323b-6ec6f57b4f26%28Office.15%29.aspx)|
|[DropCap](http://msdn.microsoft.com/library/a5c29dd4-62f4-39fb-4b76-390d62bd8e32%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/545dbfdb-4cd5-99b1-1ba3-b723e8d7b827%28Office.15%29.aspx)|
|[End](http://msdn.microsoft.com/library/594cc4b8-d7fb-4b81-4be7-2d416ae513e2%28Office.15%29.aspx)|
|[Fields](http://msdn.microsoft.com/library/01efbcae-b65b-68d9-20b0-6bbee31fd762%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/453e1507-a02d-a91b-730b-fb4a13396dbc%28Office.15%29.aspx)|
|[Font](http://msdn.microsoft.com/library/c5795f33-4e7b-f765-9ba8-f5b6706561d6%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/0cf1f043-532c-3ffc-67cf-389adc5ac02f%28Office.15%29.aspx)|
|[InlineShapes](http://msdn.microsoft.com/library/ffe2d8f2-e1d7-44ea-00fd-3c6523c9fe44%28Office.15%29.aspx)|
|[LanguageID](http://msdn.microsoft.com/library/1007c821-cafd-0cb3-94f4-4ac25decad30%28Office.15%29.aspx)|
|[Length](http://msdn.microsoft.com/library/003b4ad1-2c09-17c9-279b-b1cf2ebdb40a%28Office.15%29.aspx)|
|[LinesCount](http://msdn.microsoft.com/library/0764107c-422d-5c97-1fd5-feae43579759%28Office.15%29.aspx)|
|[MajorityFont](http://msdn.microsoft.com/library/b0007ebc-ed0b-aab8-49fe-76353efbc1d2%28Office.15%29.aspx)|
|[MajorityParagraphFormat](http://msdn.microsoft.com/library/d67e81fe-ab9b-8bfd-c31d-76feb1b6e15b%28Office.15%29.aspx)|
|[ParagraphFormat](http://msdn.microsoft.com/library/475da411-9292-a12d-addd-1bbe822ec09e%28Office.15%29.aspx)|
|[ParagraphsCount](http://msdn.microsoft.com/library/ba9cf774-b10f-3585-fc11-4b9ab6dc602d%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/1919f251-04ae-c521-34fa-aeff0d9177c1%28Office.15%29.aspx)|
|[Script](http://msdn.microsoft.com/library/54e5a19f-9cb0-0fbc-5ebe-cd4db6c0de8e%28Office.15%29.aspx)|
|[Start](http://msdn.microsoft.com/library/40604058-7c3e-b4c7-c793-bbf09091b4c1%28Office.15%29.aspx)|
|[Story](http://msdn.microsoft.com/library/833f9537-5c11-a4d5-907a-777eaecb89d2%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/13584812-307a-c32b-ca8f-27869728b64e%28Office.15%29.aspx)|
|[WordsCount](http://msdn.microsoft.com/library/93d13801-b126-7ec9-8f79-89260f8f0140%28Office.15%29.aspx)|

