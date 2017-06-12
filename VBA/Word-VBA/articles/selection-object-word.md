---
title: Selection Object (Word)
keywords: vbawd10.chm2421
f1_keywords:
- vbawd10.chm2421
ms.prod: word
api_name:
- Word.Selection
ms.assetid: 7b574a91-c33e-ecfd-6783-6b7528b2ed8f
ms.date: 06/08/2017
---


# Selection Object (Word)

Represents the current selection in a window or pane. A selection represents either a selected (or highlighted) area in the document, or it represents the insertion point if nothing in the document is selected. There can be only one  **Selection** object per document window pane, and only one **Selection** object in the entire application can be active.


## Remarks

Use the  **Selection** property to return the **Selection** object. If no object qualifier is used with the **Selection** property, Microsoft Word returns the selection from the active pane of the active document window. The following example copies the current selection from the active document.


```
Selection.Copy
```

The following example deletes the selection from the third document in the  **Documents** collection. The document does not have to be active to access its current selection.




```
Documents(3).ActiveWindow.Selection.Cut
```

The following example copies the selection from the first pane of the active document and pastes it into the second pane.




```
ActiveDocument.ActiveWindow.Panes(1).Selection.Copy 
ActiveDocument.ActiveWindow.Panes(2).Selection.Paste
```

The  **Text** property is the default property of the **Selection** object. Use this property to set or return the text in the current selection. The following example assigns the text in the current selection to the variable `strTemp`, removing the last character if it is a paragraph mark.




```
Dim strTemp as String 
 
strTemp = Selection.Text 
If Right(strTemp, 1) = vbCr Then _ 
 strTemp = Left(strTemp, Len(strTemp) - 1)
```

The  **Selection** object has various methods and properties with which you can collapse, expand, or otherwise change the current selection. The following example moves the insertion point to the end of the document and selects the last three lines.




```
Selection.EndOf Unit:=wdStory, Extend:=wdMove 
Selection.HomeKey Unit:=wdLine, Extend:=wdExtend 
Selection.MoveUp Unit:=wdLine, Count:=2, Extend:=wdExtend
```

The  **Selection** object has various methods and properties with which you can edit selected text in a document. The following example selects the first sentence in the active document and replaces it with a new paragraph.




```
Options.ReplaceSelection = True 
ActiveDocument.Sentences(1).Select 
Selection.TypeText "Material below is confidential." 
Selection.TypeParagraph
```

The following example deletes the last paragraph of the first document in the  **Documents** collection and pastes it at the beginning of the second document.




```
With Documents(1) 
 .Paragraphs.Last.Range.Select 
 .ActiveWindow.Selection.Cut 
End With 
 
With Documents(2).ActiveWindow.Selection 
 .StartOf Unit:=wdStory, Extend:=wdMove 
 .Paste 
End With
```

The  **Selection** object has various methods and properties with which you can change the formatting of the current selection. The following example changes the font of the current selection from Times New Roman to Tahoma.




```
If Selection.Font.Name = "Times New Roman" Then _ 
 Selection.Font.Name = "Tahoma"
```

Use properties like  **Flags**, **Information**, and **Type** to return information about the current selection. You can use the following example in a procedure to determine whether there is anything selected in the active document; if there is not, the rest of the procedure is skipped.




```
If Selection.Type = wdSelectionIP Then 
 MsgBox Prompt:="You have not selected any text! Exiting procedure..." 
 Exit Sub 
End If
```

Even when a selection is collapsed to an insertion point, it is not necessarily empty. For example, the  **Text** property will still return the character to the right of the insertion point; this character also appears in the **Characters** collection of the **Selection** object. However, calling methods like **Cut** or **Copy** from a collapsed selection causes an error.

It is possible for the user to select a region in a document that does not represent contiguous text (for example, when using the ALT key with the mouse). Because the behavior of such a selection can be unpredictable, you may want to include a step in your code that checks the  **Type** property of a selection before performing any operations on it ( `Selection.Type = wdSelectionBlock`). Similarly, selections that include table cells can also lead to unpredictable behavior. The  **Information** property will tell you if a selection is inside a table ( `Selection.Information(wdWithinTable) = True`). The following example determines if a selection is normal (for example, it is not a row or column in a table, it is not a vertical block of text); you could use it to test the current selection before performing any operations on it.




```
If Selection.Type <> wdSelectionNormal Then 
 MsgBox Prompt:="Not a valid selection! Exiting procedure..." 
 Exit Sub 
End If
```

Because  **Range** objects share many of the same methods and properties as **Selection** objects, using **Range** objects is preferable for manipulating a document when there is not a reason to physically change the current selection. For more information about **Selection** and **Range** objects, see [Working with the Selection object](http://msdn.microsoft.com/library/a1ef7e48-5a0f-d278-4b67-7b96f4e24052%28Office.15%29.aspx) and [Working with Range objects](http://msdn.microsoft.com/library/9e240aa7-8608-9d70-aee3-2e202687459e%28Office.15%29.aspx).


## Methods



|**Name**|
|:-----|
|[BoldRun](http://msdn.microsoft.com/library/0998afe2-dcd9-c1e4-9614-a1af4c6bbeaf%28Office.15%29.aspx)|
|[Calculate](http://msdn.microsoft.com/library/a4e7ef08-8442-0579-e738-e4f53ee62d62%28Office.15%29.aspx)|
|[ClearCharacterAllFormatting](http://msdn.microsoft.com/library/1d0dfb43-4855-1534-5ec2-475232a6a457%28Office.15%29.aspx)|
|[ClearCharacterDirectFormatting](http://msdn.microsoft.com/library/d2138876-c832-2407-a53e-5bd4af2421b7%28Office.15%29.aspx)|
|[ClearCharacterStyle](http://msdn.microsoft.com/library/ff9795f9-ea74-fa03-5d87-9c56152d179d%28Office.15%29.aspx)|
|[ClearFormatting](http://msdn.microsoft.com/library/66c2f088-5d35-f8b0-10e5-2faa0db14d7f%28Office.15%29.aspx)|
|[ClearParagraphAllFormatting](http://msdn.microsoft.com/library/b3a88322-933a-ff14-e788-e1934aba243d%28Office.15%29.aspx)|
|[ClearParagraphDirectFormatting](http://msdn.microsoft.com/library/66df2319-f02e-7cd9-4cef-fda6468dcd67%28Office.15%29.aspx)|
|[ClearParagraphStyle](http://msdn.microsoft.com/library/cfbafeac-99e1-5fae-a9a0-8cf8836add94%28Office.15%29.aspx)|
|[Collapse](http://msdn.microsoft.com/library/92ccd3dc-41ab-b3d4-5397-fca7d7f01635%28Office.15%29.aspx)|
|[ConvertToTable](http://msdn.microsoft.com/library/b2f487dd-7a10-5e5f-74f1-a2e9b5e9d958%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/5af32d69-5c0f-428a-44f3-35c75b5fb050%28Office.15%29.aspx)|
|[CopyAsPicture](http://msdn.microsoft.com/library/f5c73e30-1601-62a7-ec0e-2dc49c6f51fe%28Office.15%29.aspx)|
|[CopyFormat](http://msdn.microsoft.com/library/ef892e50-2ff1-3ab0-1112-cf6d268a1103%28Office.15%29.aspx)|
|[CreateAutoTextEntry](http://msdn.microsoft.com/library/def6f758-af70-eaf2-f15c-4a6a28c247b5%28Office.15%29.aspx)|
|[CreateTextbox](http://msdn.microsoft.com/library/e3c567ee-949f-5e87-43c2-633cdae334b0%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/1e5dec1a-c621-2b54-ab7f-78ce90c0936f%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/35bfdf19-62d3-5593-0b2f-dd6b642b4cc3%28Office.15%29.aspx)|
|[DetectLanguage](http://msdn.microsoft.com/library/cfbc0d54-bb00-2bd0-ad9a-e646fdcbfe46%28Office.15%29.aspx)|
|[EndKey](http://msdn.microsoft.com/library/4f27681c-1117-99c2-1aba-bd97082bb8ba%28Office.15%29.aspx)|
|[EndOf](http://msdn.microsoft.com/library/33aa094b-17f9-3572-f66f-59692c57dc01%28Office.15%29.aspx)|
|[EscapeKey](http://msdn.microsoft.com/library/a498cf00-a3dc-b084-79ae-c31d6f4e5e27%28Office.15%29.aspx)|
|[Expand](http://msdn.microsoft.com/library/8b716453-7656-e8b8-f6b0-0dc97ef2714d%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/0fc22f07-6a21-d04e-e90b-73e33f5e4f36%28Office.15%29.aspx)|
|[Extend](http://msdn.microsoft.com/library/7f9108a1-9b23-bc45-61f5-49aca9979932%28Office.15%29.aspx)|
|[GoTo](http://msdn.microsoft.com/library/7a69e581-4047-ae62-e112-97fe2c2633bb%28Office.15%29.aspx)|
|[GoToEditableRange](http://msdn.microsoft.com/library/01c287a4-9293-22c1-9439-4a069a1e7299%28Office.15%29.aspx)|
|[GoToNext](http://msdn.microsoft.com/library/af6a4e91-7ec1-929a-7577-4e457f5ce1bd%28Office.15%29.aspx)|
|[GoToPrevious](http://msdn.microsoft.com/library/da41b0b4-673e-5701-d31d-ab3314600e53%28Office.15%29.aspx)|
|[HomeKey](http://msdn.microsoft.com/library/24264193-d610-acbc-b393-de41fd55e976%28Office.15%29.aspx)|
|[InRange](http://msdn.microsoft.com/library/3759ad96-44b5-d63c-f4d5-844f937f4216%28Office.15%29.aspx)|
|[InsertAfter](http://msdn.microsoft.com/library/21286a89-5e4e-56ae-27a5-f581a337bfbb%28Office.15%29.aspx)|
|[InsertBefore](http://msdn.microsoft.com/library/05dfc75f-9bb3-e090-9b31-aeb48b6c2ed8%28Office.15%29.aspx)|
|[InsertBreak](http://msdn.microsoft.com/library/2c9d8cb8-1cc1-3d69-1e26-3a6878c0b1da%28Office.15%29.aspx)|
|[InsertCaption](http://msdn.microsoft.com/library/848c1686-ca8c-d022-68f1-74a2f3d46498%28Office.15%29.aspx)|
|[InsertCells](http://msdn.microsoft.com/library/461085a3-ae98-8028-5ad2-d5e22038c6db%28Office.15%29.aspx)|
|[InsertColumns](http://msdn.microsoft.com/library/d58691b4-afa5-959a-a6a8-f202723df9f1%28Office.15%29.aspx)|
|[InsertColumnsRight](http://msdn.microsoft.com/library/0367ae17-d5f0-90f6-7834-4856ff7a1530%28Office.15%29.aspx)|
|[InsertCrossReference](http://msdn.microsoft.com/library/3aa9261d-8e2a-6230-8f02-629f0a0104bf%28Office.15%29.aspx)|
|[InsertDateTime](http://msdn.microsoft.com/library/f9cfca41-e0f2-4656-5fa2-2463c50af1f5%28Office.15%29.aspx)|
|[InsertFile](http://msdn.microsoft.com/library/963a5987-e6f8-824a-47d6-9788f026cf10%28Office.15%29.aspx)|
|[InsertFormula](http://msdn.microsoft.com/library/a193c4ee-a667-04af-e22c-3a5b5bbc5c3b%28Office.15%29.aspx)|
|[InsertNewPage](http://msdn.microsoft.com/library/54c7e18a-6dfb-b8da-0e6d-3c53e71f42cd%28Office.15%29.aspx)|
|[InsertParagraph](http://msdn.microsoft.com/library/bceda293-7294-8769-75fe-4792199439c1%28Office.15%29.aspx)|
|[InsertParagraphAfter](http://msdn.microsoft.com/library/ae97fbab-417a-14e2-0154-f0361826f903%28Office.15%29.aspx)|
|[InsertParagraphBefore](http://msdn.microsoft.com/library/f4843e0b-0d0f-ef6f-6f7a-423b49dceb50%28Office.15%29.aspx)|
|[InsertRows](http://msdn.microsoft.com/library/326ad049-4d39-1ca6-a203-ddba0e77cba4%28Office.15%29.aspx)|
|[InsertRowsAbove](http://msdn.microsoft.com/library/f5387043-34d0-cd84-6550-bfd96bf661b8%28Office.15%29.aspx)|
|[InsertRowsBelow](http://msdn.microsoft.com/library/d36441d1-ff1f-b557-d0d0-1d12d4abab2d%28Office.15%29.aspx)|
|[InsertStyleSeparator](http://msdn.microsoft.com/library/cbfd7a55-4048-0e16-eeb2-e8d8d167a769%28Office.15%29.aspx)|
|[InsertSymbol](http://msdn.microsoft.com/library/13f18c60-89e7-3ba7-1c4c-928b28f5e72a%28Office.15%29.aspx)|
|[InsertXML](http://msdn.microsoft.com/library/7a9e52b5-9b05-f939-6fd0-33a923989f48%28Office.15%29.aspx)|
|[InStory](http://msdn.microsoft.com/library/29dae109-4361-f1ee-eb71-76f57ae186a3%28Office.15%29.aspx)|
|[IsEqual](http://msdn.microsoft.com/library/57ca55bc-17cf-054c-81dd-aa6d1e536cd8%28Office.15%29.aspx)|
|[ItalicRun](http://msdn.microsoft.com/library/0d36eff1-7308-7695-7058-be79455836ee%28Office.15%29.aspx)|
|[LtrPara](http://msdn.microsoft.com/library/992886b8-44e3-3b1f-cc6d-7b16e1c58aef%28Office.15%29.aspx)|
|[LtrRun](http://msdn.microsoft.com/library/e2b905f1-3ce1-ce51-bc9f-c5325fa0e9af%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/8bd36cf4-ca05-6412-2145-31d07367730e%28Office.15%29.aspx)|
|[MoveDown](http://msdn.microsoft.com/library/d3ea31e8-04a5-c342-24ca-c93ac1a1258e%28Office.15%29.aspx)|
|[MoveEnd](http://msdn.microsoft.com/library/11fbcd45-16e6-611b-d296-a88cc7d3ca50%28Office.15%29.aspx)|
|[MoveEndUntil](http://msdn.microsoft.com/library/e8f7532a-6a5a-3173-3e5e-db46aec44170%28Office.15%29.aspx)|
|[MoveEndWhile](http://msdn.microsoft.com/library/c1cd97cc-9836-a61f-3b7a-4e178bc3d1e0%28Office.15%29.aspx)|
|[MoveLeft](http://msdn.microsoft.com/library/23c22588-e774-f70f-28ea-81b1a54c0dd5%28Office.15%29.aspx)|
|[MoveRight](http://msdn.microsoft.com/library/fcac96c7-7189-87b2-d800-9d161edb1e09%28Office.15%29.aspx)|
|[MoveStart](http://msdn.microsoft.com/library/c58f4dd5-791b-ac0f-8445-29e0ade48d7f%28Office.15%29.aspx)|
|[MoveStartUntil](http://msdn.microsoft.com/library/a461cf49-1ed9-425b-5417-0a882c17d792%28Office.15%29.aspx)|
|[MoveStartWhile](http://msdn.microsoft.com/library/b6e33ffc-a07f-2ef9-0e35-55aaf256f098%28Office.15%29.aspx)|
|[MoveUntil](http://msdn.microsoft.com/library/888655d0-44ec-b589-bd0b-b3e193e413ef%28Office.15%29.aspx)|
|[MoveUp](http://msdn.microsoft.com/library/46993371-c916-06b5-a644-960f8a283536%28Office.15%29.aspx)|
|[MoveWhile](http://msdn.microsoft.com/library/ba35991c-2ae3-e78f-7538-c102149cf392%28Office.15%29.aspx)|
|[Next](http://msdn.microsoft.com/library/498db129-c3bd-2f9c-5897-fcfda6ce5d14%28Office.15%29.aspx)|
|[NextField](http://msdn.microsoft.com/library/40007462-3bb5-59a7-89cb-27d654795e76%28Office.15%29.aspx)|
|[NextRevision](http://msdn.microsoft.com/library/990e3c20-9991-b2cb-aa3b-e64ae8936b34%28Office.15%29.aspx)|
|[NextSubdocument](http://msdn.microsoft.com/library/e8527994-23f4-c9a1-d96c-c2018e07efad%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/f09e3a0f-2c24-6bcb-0a97-eb33318fe6f4%28Office.15%29.aspx)|
|[PasteAndFormat](http://msdn.microsoft.com/library/7ed87209-b786-280e-f3f0-dd81eda6f82d%28Office.15%29.aspx)|
|[PasteAppendTable](http://msdn.microsoft.com/library/60e12397-563f-f8bc-160f-f24a12794d01%28Office.15%29.aspx)|
|[PasteAsNestedTable](http://msdn.microsoft.com/library/42a2f604-694e-6b39-23d2-d8c453618222%28Office.15%29.aspx)|
|[PasteExcelTable](http://msdn.microsoft.com/library/bfa7f82c-e5c0-3919-f492-a71c9eabb919%28Office.15%29.aspx)|
|[PasteFormat](http://msdn.microsoft.com/library/5c8a69fa-4d07-619c-950a-5ff11fa99003%28Office.15%29.aspx)|
|[PasteSpecial](http://msdn.microsoft.com/library/186ddf42-f8ab-e334-ccfe-245b2cc82224%28Office.15%29.aspx)|
|[Previous](http://msdn.microsoft.com/library/85679323-fe2c-f37a-5373-2c9e6d8494eb%28Office.15%29.aspx)|
|[PreviousField](http://msdn.microsoft.com/library/9361a318-9ee2-fd72-9d52-106abfd8d44e%28Office.15%29.aspx)|
|[PreviousRevision](http://msdn.microsoft.com/library/e516037f-047d-5cd2-19b4-3b7870a14b5a%28Office.15%29.aspx)|
|[PreviousSubdocument](http://msdn.microsoft.com/library/135932fa-c165-56d1-97c7-f04fd7108ab2%28Office.15%29.aspx)|
|[ReadingModeGrowFont](http://msdn.microsoft.com/library/5a23b50e-073f-1cbd-e1df-6ee846cb1ecf%28Office.15%29.aspx)|
|[ReadingModeShrinkFont](http://msdn.microsoft.com/library/58472c33-7f8e-dc3b-04d8-7b50ca911ed4%28Office.15%29.aspx)|
|[RtlPara](http://msdn.microsoft.com/library/b417897d-de70-6c3a-12cd-8786e12bdb43%28Office.15%29.aspx)|
|[RtlRun](http://msdn.microsoft.com/library/759a16cd-24d7-7c0a-6315-47d395560c73%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/129ca04c-05f0-90b5-c2fa-789038c34b2f%28Office.15%29.aspx)|
|[SelectCell](http://msdn.microsoft.com/library/49df8e0c-795d-5d5b-79e4-56e0bd64c222%28Office.15%29.aspx)|
|[SelectColumn](http://msdn.microsoft.com/library/a8e742df-0a8e-739d-e71a-da2536b6abec%28Office.15%29.aspx)|
|[SelectCurrentAlignment](http://msdn.microsoft.com/library/89d76005-c75a-7548-c1da-da292183d5ab%28Office.15%29.aspx)|
|[SelectCurrentColor](http://msdn.microsoft.com/library/f7d23b80-7e1a-40a5-b292-820c3db500a6%28Office.15%29.aspx)|
|[SelectCurrentFont](http://msdn.microsoft.com/library/66539ab3-280f-40a5-1fc0-1507b66d50fd%28Office.15%29.aspx)|
|[SelectCurrentIndent](http://msdn.microsoft.com/library/3a71080e-935c-fc3c-40b9-e82acf9d28cc%28Office.15%29.aspx)|
|[SelectCurrentSpacing](http://msdn.microsoft.com/library/1a49caa6-d261-e9d7-9d64-c564c30a7e29%28Office.15%29.aspx)|
|[SelectCurrentTabs](http://msdn.microsoft.com/library/38b0ba64-eedc-9ef5-5622-5499b50bbd3e%28Office.15%29.aspx)|
|[SelectRow](http://msdn.microsoft.com/library/0d821d49-2829-2469-4742-0355440e4775%28Office.15%29.aspx)|
|[SetRange](http://msdn.microsoft.com/library/232a681e-4205-05ae-f442-9dc1a2df96f1%28Office.15%29.aspx)|
|[Shrink](http://msdn.microsoft.com/library/ed364c95-3b9d-44dc-b120-db23aedfeaed%28Office.15%29.aspx)|
|[ShrinkDiscontiguousSelection](http://msdn.microsoft.com/library/ce703cb4-8a20-b59d-ccf7-c0c93327a9ca%28Office.15%29.aspx)|
|[Sort](http://msdn.microsoft.com/library/3f29f6bf-a6b4-1638-b078-f61a4f36c17e%28Office.15%29.aspx)|
|[SortAscending](http://msdn.microsoft.com/library/8092fdac-b89c-9a6e-1151-9611f69d0dc4%28Office.15%29.aspx)|
|[SortByHeadings](http://msdn.microsoft.com/library/fc38c337-b658-7b8d-2191-2ee98a93b48e%28Office.15%29.aspx)|
|[SortDescending](http://msdn.microsoft.com/library/057461e9-d8f8-7d9b-c985-c634dd8bef6a%28Office.15%29.aspx)|
|[SplitTable](http://msdn.microsoft.com/library/5d68a031-1927-ae5c-de11-963bca9c1d2c%28Office.15%29.aspx)|
|[StartOf](http://msdn.microsoft.com/library/570df152-3579-d7a6-f555-86c9da229e1b%28Office.15%29.aspx)|
|[ToggleCharacterCode](http://msdn.microsoft.com/library/e59774bc-cdd5-577b-8175-f988a18c0538%28Office.15%29.aspx)|
|[TypeBackspace](http://msdn.microsoft.com/library/479f2e0e-06d6-cd62-dc3e-09a5fafafbfa%28Office.15%29.aspx)|
|[TypeParagraph](http://msdn.microsoft.com/library/e866733b-4800-8e2c-7026-4e9603ccf585%28Office.15%29.aspx)|
|[TypeText](http://msdn.microsoft.com/library/fb8e58cc-0c49-0efa-d60a-8be6c3d4435c%28Office.15%29.aspx)|
|[WholeStory](http://msdn.microsoft.com/library/ecd50a78-ecbd-75a9-2565-31d7e6ac449a%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Active](http://msdn.microsoft.com/library/a279837e-8ae7-24ec-71f0-de82c5a33ad8%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/20834ace-7239-718f-cb5c-c017576216c5%28Office.15%29.aspx)|
|[BookmarkID](http://msdn.microsoft.com/library/f48d317c-b5ed-ff0e-4a22-13b68aa10be1%28Office.15%29.aspx)|
|[Bookmarks](http://msdn.microsoft.com/library/32e25786-512a-5bee-4ba6-42c801b49176%28Office.15%29.aspx)|
|[Borders](http://msdn.microsoft.com/library/2e70c7be-c7dc-db59-0a99-a11770ffc220%28Office.15%29.aspx)|
|[Cells](http://msdn.microsoft.com/library/4b808b86-42ba-ccb4-b19a-87b134df3b79%28Office.15%29.aspx)|
|[Characters](http://msdn.microsoft.com/library/605c0fc5-f5b9-6782-9fdd-54589040d243%28Office.15%29.aspx)|
|[ChildShapeRange](http://msdn.microsoft.com/library/1b7c1010-19e1-e849-0040-70e231aac133%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/444726a7-0bbe-8d66-b3ca-113af074e673%28Office.15%29.aspx)|
|[ColumnSelectMode](http://msdn.microsoft.com/library/de146d32-63aa-3a17-6eeb-32cccf3f8bfd%28Office.15%29.aspx)|
|[Comments](http://msdn.microsoft.com/library/8f6fda0e-7070-eb42-3e1b-3a2a0654b330%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/612e00fa-49cd-5f9e-7a10-9504d4e7b408%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/03b4bfd7-8d4a-f069-0c28-41be2ead8614%28Office.15%29.aspx)|
|[Editors](http://msdn.microsoft.com/library/ba750568-88c9-9ed8-61ee-45f89dfa4dea%28Office.15%29.aspx)|
|[End](http://msdn.microsoft.com/library/99e3bd79-a8f1-8057-1ac2-b9e76eab99ff%28Office.15%29.aspx)|
|[EndnoteOptions](http://msdn.microsoft.com/library/23b7263c-7322-3221-6436-ee0c614fa577%28Office.15%29.aspx)|
|[Endnotes](http://msdn.microsoft.com/library/fea9ea39-4091-cccd-9025-36be2e4b95ff%28Office.15%29.aspx)|
|[EnhMetaFileBits](http://msdn.microsoft.com/library/ecc28cc8-6c0f-3207-f52c-4a7b77c23445%28Office.15%29.aspx)|
|[ExtendMode](http://msdn.microsoft.com/library/7b12cf8b-9be1-6ebc-de96-e7734eaad3b6%28Office.15%29.aspx)|
|[Fields](http://msdn.microsoft.com/library/15060502-c0cf-1c94-93ba-0db0bb045c66%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/66004412-4da2-586d-887c-6f9867e06ea6%28Office.15%29.aspx)|
|[FitTextWidth](http://msdn.microsoft.com/library/7f7409b4-c533-9c21-2663-e4016416efb7%28Office.15%29.aspx)|
|[Flags](http://msdn.microsoft.com/library/bca92e77-077c-57d0-3012-8c064e93f112%28Office.15%29.aspx)|
|[Font](http://msdn.microsoft.com/library/c2a24190-62fa-09c4-7c47-90a7ecf20d97%28Office.15%29.aspx)|
|[FootnoteOptions](http://msdn.microsoft.com/library/064bb3c1-cbaa-9d8f-5b97-a4337b0cfeae%28Office.15%29.aspx)|
|[Footnotes](http://msdn.microsoft.com/library/61829c93-46e9-c1c5-1424-fb34a812a76d%28Office.15%29.aspx)|
|[FormattedText](http://msdn.microsoft.com/library/b16da3f9-1aa6-e722-0a9c-8a4c30922450%28Office.15%29.aspx)|
|[FormFields](http://msdn.microsoft.com/library/d6d5259b-9971-929f-16f7-ca2b2d585c77%28Office.15%29.aspx)|
|[Frames](http://msdn.microsoft.com/library/cc589559-858a-2ebb-00dd-64f97966859f%28Office.15%29.aspx)|
|[HasChildShapeRange](http://msdn.microsoft.com/library/1917754f-6080-8303-533e-b62607b87d41%28Office.15%29.aspx)|
|[HeaderFooter](http://msdn.microsoft.com/library/b2eeeb83-49bf-236e-e795-6231ff20e368%28Office.15%29.aspx)|
|[HTMLDivisions](http://msdn.microsoft.com/library/707a44e8-80a4-bd78-f1d6-cda05910bb23%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/c90c3779-cbb9-4174-3002-850750b4bb41%28Office.15%29.aspx)|
|[Information](http://msdn.microsoft.com/library/73028751-6339-47e6-9629-9584cc4c59ec%28Office.15%29.aspx)|
|[InlineShapes](http://msdn.microsoft.com/library/2fbbf39c-b70e-e332-2547-089166e718ca%28Office.15%29.aspx)|
|[IPAtEndOfLine](http://msdn.microsoft.com/library/8db37c0f-6c68-7ccd-0c34-9a40b62b9e9d%28Office.15%29.aspx)|
|[IsEndOfRowMark](http://msdn.microsoft.com/library/0729a8f2-628c-902b-fca1-488742234873%28Office.15%29.aspx)|
|[LanguageDetected](http://msdn.microsoft.com/library/289e6a01-1945-a17f-f6a0-e472cfa263eb%28Office.15%29.aspx)|
|[LanguageID](http://msdn.microsoft.com/library/d92be532-99db-8b46-3e64-8a3fca65004e%28Office.15%29.aspx)|
|[LanguageIDFarEast](http://msdn.microsoft.com/library/59f5b72f-3ba5-cff8-8465-6759d2194d26%28Office.15%29.aspx)|
|[LanguageIDOther](http://msdn.microsoft.com/library/197474ff-8d79-b48f-e1bf-ac2f164e70e3%28Office.15%29.aspx)|
|[NoProofing](http://msdn.microsoft.com/library/5feca11c-5afa-80aa-b854-bab86b49a749%28Office.15%29.aspx)|
|[OMaths](http://msdn.microsoft.com/library/ca63d636-1f78-e075-087b-2d8d55254406%28Office.15%29.aspx)|
|[Orientation](http://msdn.microsoft.com/library/830f85de-51bc-1cb9-6312-7e2b2751430d%28Office.15%29.aspx)|
|[PageSetup](http://msdn.microsoft.com/library/65e8953b-0b52-b31f-1c81-e607a37ba916%28Office.15%29.aspx)|
|[ParagraphFormat](http://msdn.microsoft.com/library/3a3a3b4e-396f-fbe5-dc30-649ef7a9a8f9%28Office.15%29.aspx)|
|[Paragraphs](http://msdn.microsoft.com/library/f237788a-01e4-62ce-d698-3af619c90272%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/dc2ad5ae-a9ae-4f79-2900-895782cd78f6%28Office.15%29.aspx)|
|[PreviousBookmarkID](http://msdn.microsoft.com/library/33d7490d-1b48-81a1-a7d5-9154c1d92230%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/10161d3b-0fa9-209e-a120-be690746ccfc%28Office.15%29.aspx)|
|[Rows](http://msdn.microsoft.com/library/800edca7-fc0f-ed73-ae3a-400eadcccf8b%28Office.15%29.aspx)|
|[Sections](http://msdn.microsoft.com/library/ac8c26f3-6870-cd25-6f10-21efd16d47d9%28Office.15%29.aspx)|
|[Sentences](http://msdn.microsoft.com/library/38d0e311-5033-bada-005b-3be642a618c1%28Office.15%29.aspx)|
|[Shading](http://msdn.microsoft.com/library/1e259969-7a0a-aaf3-af6c-81e0b37b6f79%28Office.15%29.aspx)|
|[ShapeRange](http://msdn.microsoft.com/library/b327da9a-8858-1ec1-8a0d-cb36bd44fede%28Office.15%29.aspx)|
|[Start](http://msdn.microsoft.com/library/e1928372-2473-e377-4ba1-894b104fcf43%28Office.15%29.aspx)|
|[StartIsActive](http://msdn.microsoft.com/library/734e5368-dd6e-d84a-b445-30540948ac7a%28Office.15%29.aspx)|
|[StoryLength](http://msdn.microsoft.com/library/adc9f016-5e8f-d9ef-bd5c-9f322a6c0e58%28Office.15%29.aspx)|
|[StoryType](http://msdn.microsoft.com/library/17709b74-ea6b-9d58-885d-01e6b2ddac55%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/d9295c79-97bd-3866-8321-45b708154716%28Office.15%29.aspx)|
|[Tables](http://msdn.microsoft.com/library/1639cfda-d347-0227-6a4c-8f269c81230f%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/2acf885b-8d4a-7ebc-79aa-902921bc33bb%28Office.15%29.aspx)|
|[TopLevelTables](http://msdn.microsoft.com/library/7ab1b2a3-85a8-8892-53b9-dc85ff747078%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/75af6b1a-c9d3-e3ad-52a8-41d91c79b007%28Office.15%29.aspx)|
|[WordOpenXML](http://msdn.microsoft.com/library/750cf7f5-5f72-cae3-e026-2205df2f32c2%28Office.15%29.aspx)|
|[Words](http://msdn.microsoft.com/library/bbbc7c5f-ce5a-2608-ba0c-e9769bff287a%28Office.15%29.aspx)|
|[XML](http://msdn.microsoft.com/library/d7a810ea-10c0-5ac6-b8dd-55a934e5df42%28Office.15%29.aspx)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

