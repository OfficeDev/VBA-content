---
title: TextRange Object (PowerPoint)
keywords: vbapp10.chm569000
f1_keywords:
- vbapp10.chm569000
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange
ms.assetid: 7c234107-c423-7ec9-e8bd-a82cc3b345de
ms.date: 06/08/2017
---


# TextRange Object (PowerPoint)

Contains the text that's attached to a shape, and properties and methods for manipulating the text.


## Remarks

The following examples describe how to:


- Return the text range in any shape you specify.
    
- Return a text range from the selection.
    
- Return particular characters, words, lines, sentences, or paragraphs from a text range.
    
- Find and replace text in a text range.
    
- Insert text, the date and time, or the slide number into a text range.
    
- Position the cursor wherever you want in a text range.
    

## Example

Use the [TextRange](http://msdn.microsoft.com/library/4a565e39-8bfa-7370-3ed6-57c442796144%28Office.15%29.aspx)property of the  **[TextFrame](textframe-object-powerpoint.md)** object to return a **TextRange** object for any shape you specify. Use the[Text](http://msdn.microsoft.com/library/c80c8b19-73e2-0820-abd6-c44f4b2644b2%28Office.15%29.aspx)property to return the string of text in the  **TextRange** object. The following example adds a rectangle to `myDocument` and sets the text it contains.


```
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140) _

    .TextFrame.TextRange.Text = "Here is some test text"
```

Because the  **Text** property is the default property of the **TextRange** object, the following two statements are equivalent.




```
ActivePresentation.Slides(1).Shapes(1).TextFrame _

    .TextRange.Text = "Here is some test text"

ActivePresentation.Slides(1).Shapes(1).TextFrame _

    .TextRange = "Here is some test text"
```

Use the [HasTextFrame](http://msdn.microsoft.com/library/ea1a53e4-32d8-e51f-9e60-9ef719c0d973%28Office.15%29.aspx)property to determine whether a shape has a text frame, and use the [HasText](http://msdn.microsoft.com/library/7bce3bae-38e7-d9d4-b67c-9454fafc620f%28Office.15%29.aspx)property to determine whether the text frame contains text.

Use the  **TextRange** property of the **Selection** object to return the currently selected text. The following example copies the selection to the Clipboard.




```
ActiveWindow.Selection.TextRange.Copy
```

Use one of the following methods to return a portion of the text of a  **TextRange** object: **[Characters](http://msdn.microsoft.com/library/019c15d3-349d-ab10-7448-70bf81176150%28Office.15%29.aspx)**, **[Lines](http://msdn.microsoft.com/library/8e9f344b-2e74-5a9d-06e8-3e6ff9ca6bd0%28Office.15%29.aspx)**, **[Paragraphs](http://msdn.microsoft.com/library/5062eccf-4db2-692f-501e-b7d214181171%28Office.15%29.aspx)**, **[Runs](http://msdn.microsoft.com/library/0bf2724a-0735-bd79-31e5-894d1320b9b2%28Office.15%29.aspx)**, **[Sentences](http://msdn.microsoft.com/library/c3640cb8-f78a-2934-bbe0-506cb8d2534c%28Office.15%29.aspx)**, or **[Words](http://msdn.microsoft.com/library/b8cd8dca-bf10-1041-dd9e-adc04b2df42d%28Office.15%29.aspx)**.

Use the [Find](http://msdn.microsoft.com/library/24186821-3a0a-efd5-c35a-8b553e00f92b%28Office.15%29.aspx)and [Replace](http://msdn.microsoft.com/library/046d1c3d-fd3e-7871-e31e-6529b77fcd60%28Office.15%29.aspx)methods to find and replace text in a text range.

Use one of the following methods to insert characters into a  **TextRange** object:[InsertAfter](http://msdn.microsoft.com/library/2af4e134-c205-fbe6-a006-3fc1ca8d6a50%28Office.15%29.aspx), [InsertBefore](http://msdn.microsoft.com/library/fbadcecd-a31b-8c8d-3281-63d70286bcff%28Office.15%29.aspx), [InsertDateTime](http://msdn.microsoft.com/library/b1f6c2db-2524-f76e-eee2-8f177b08dcde%28Office.15%29.aspx), [InsertSlideNumber](http://msdn.microsoft.com/library/07489db8-9db1-9721-845a-7895ad316aca%28Office.15%29.aspx), or [InsertSymbol](http://msdn.microsoft.com/library/a424e011-1bfe-f690-cbc0-604f89718831%28Office.15%29.aspx).


## Methods



|**Name**|
|:-----|
|[AddPeriods](http://msdn.microsoft.com/library/597592ba-6c26-7645-33b8-19ccb375a098%28Office.15%29.aspx)|
|[ChangeCase](http://msdn.microsoft.com/library/a14edb26-7ec3-5fb5-7590-cd67a75c1f03%28Office.15%29.aspx)|
|[Characters](http://msdn.microsoft.com/library/019c15d3-349d-ab10-7448-70bf81176150%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/c8d1edf7-68ef-aaa4-e2db-717263df8dd3%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/9be71668-1486-0466-f87b-47792d402102%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/2baac89b-d7b1-2209-b17f-b65b71b5aed4%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/24186821-3a0a-efd5-c35a-8b553e00f92b%28Office.15%29.aspx)|
|[InsertAfter](http://msdn.microsoft.com/library/2af4e134-c205-fbe6-a006-3fc1ca8d6a50%28Office.15%29.aspx)|
|[InsertBefore](http://msdn.microsoft.com/library/fbadcecd-a31b-8c8d-3281-63d70286bcff%28Office.15%29.aspx)|
|[InsertDateTime](http://msdn.microsoft.com/library/b1f6c2db-2524-f76e-eee2-8f177b08dcde%28Office.15%29.aspx)|
|[InsertSlideNumber](http://msdn.microsoft.com/library/07489db8-9db1-9721-845a-7895ad316aca%28Office.15%29.aspx)|
|[InsertSymbol](http://msdn.microsoft.com/library/a424e011-1bfe-f690-cbc0-604f89718831%28Office.15%29.aspx)|
|[Lines](http://msdn.microsoft.com/library/8e9f344b-2e74-5a9d-06e8-3e6ff9ca6bd0%28Office.15%29.aspx)|
|[LtrRun](http://msdn.microsoft.com/library/5c6787cc-d37c-8aec-b49e-12418291e006%28Office.15%29.aspx)|
|[Paragraphs](http://msdn.microsoft.com/library/5062eccf-4db2-692f-501e-b7d214181171%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/4bbaa1f3-206e-2009-11f0-5abde24517c6%28Office.15%29.aspx)|
|[PasteSpecial](http://msdn.microsoft.com/library/97bfd298-f8e8-32f0-b05c-6a93ed651954%28Office.15%29.aspx)|
|[RemovePeriods](http://msdn.microsoft.com/library/66562cc7-e25b-e110-000e-c01b62caf761%28Office.15%29.aspx)|
|[Replace](http://msdn.microsoft.com/library/046d1c3d-fd3e-7871-e31e-6529b77fcd60%28Office.15%29.aspx)|
|[RotatedBounds](http://msdn.microsoft.com/library/33a4393e-3b87-a4d3-0e16-8230a4beabe3%28Office.15%29.aspx)|
|[RtlRun](http://msdn.microsoft.com/library/eb474c9b-d789-f741-9ba9-0514f0a5b0be%28Office.15%29.aspx)|
|[Runs](http://msdn.microsoft.com/library/0bf2724a-0735-bd79-31e5-894d1320b9b2%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/cd6fb1ba-ac49-a7d8-2777-fda2ce2746a4%28Office.15%29.aspx)|
|[Sentences](http://msdn.microsoft.com/library/c3640cb8-f78a-2934-bbe0-506cb8d2534c%28Office.15%29.aspx)|
|[TrimText](http://msdn.microsoft.com/library/8566ed9d-c73a-d699-bcb7-edcd9a375afe%28Office.15%29.aspx)|
|[Words](http://msdn.microsoft.com/library/b8cd8dca-bf10-1041-dd9e-adc04b2df42d%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActionSettings](http://msdn.microsoft.com/library/7a66ca28-d6b9-2066-4c88-a04888d5ba1e%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/53b4f6fc-7e1b-7045-e59d-eec668a75d3e%28Office.15%29.aspx)|
|[BoundHeight](http://msdn.microsoft.com/library/8f3b9947-5ee3-260d-3d44-0ad2da422724%28Office.15%29.aspx)|
|[BoundLeft](http://msdn.microsoft.com/library/2641e084-6b6e-ff6e-c6a6-27cb84cbd4dd%28Office.15%29.aspx)|
|[BoundTop](http://msdn.microsoft.com/library/cfc3baec-06c4-da2f-a233-afcb5301302a%28Office.15%29.aspx)|
|[BoundWidth](http://msdn.microsoft.com/library/409d1c66-8956-cdd0-2328-f1cbe584f778%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/9c514376-18ef-1eac-661a-c1fc46514b32%28Office.15%29.aspx)|
|[Font](http://msdn.microsoft.com/library/234c8843-3c0d-a425-0173-02c3910ba400%28Office.15%29.aspx)|
|[IndentLevel](http://msdn.microsoft.com/library/3ba39fc4-6fc4-62ca-0e87-a7605d6c8bea%28Office.15%29.aspx)|
|[LanguageID](http://msdn.microsoft.com/library/f6744845-5125-239e-65d1-7db8dacdaecd%28Office.15%29.aspx)|
|[Length](http://msdn.microsoft.com/library/4eb64830-f8e4-5226-57c1-80df7f4bd39f%28Office.15%29.aspx)|
|[ParagraphFormat](http://msdn.microsoft.com/library/41d3f0f3-70e3-ad1a-efcb-de849d4a03d4%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/303cc0cf-8c1c-60af-648e-fea4d25abb36%28Office.15%29.aspx)|
|[Start](http://msdn.microsoft.com/library/1e37b589-842e-b03b-09eb-a19ce03f6a72%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/c80c8b19-73e2-0820-abd6-c44f4b2644b2%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
