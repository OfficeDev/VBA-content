---
title: Range Object (Word)
keywords: vbawd10.chm2398
f1_keywords:
- vbawd10.chm2398
ms.prod: word
api_name:
- Word.Range
ms.assetid: 15a7a1c4-5f3f-5b6e-60e9-29688de3f274
ms.date: 06/08/2017
---


# Range Object (Word)

Represents a contiguous area in a document. Each  **Range** object is defined by a starting and ending character position.


## Remarks

Similar to the way bookmarks are used in a document,  **Range** objects are used in Visual Basic procedures to identify specific portions of a document. However, unlike a bookmark, a **Range** object only exists while the procedure that defined it is running. **Range** objects are independent of the selection. That is, you can define and manipulate a range without changing the selection. You can also define multiple ranges in a document, while there can be only one selection per pane.

Use the  **Range** method to return a **Range** object defined by the given starting and ending character positions. The following example returns a **Range** object that refers to the first 10 characters in the active document.




```
Set myRange = ActiveDocument.Range(Start:=0, End:=10)
```

Use the  **Range** property to return a **Range** object defined by the beginning and end of another object. The **Range** property applies to many objects (for example, **Paragraph**, **Bookmark**, and **Cell** ). The following example returns a **Range** object that refers to the first paragraph in the active document.




```
Set aRange = ActiveDocument.Paragraphs(1).Range
```

The following example returns a  **Range** object that refers to the second through fourth paragraphs in the active document




```
Set aRange = ActiveDocument.Range( _ 
 Start:=ActiveDocument.Paragraphs(2).Range.Start, _ 
 End:=ActiveDocument.Paragraphs(4).Range.End)
```

For more information about working with  **Range** objects, see [Working with Range Objects](http://msdn.microsoft.com/library/9e240aa7-8608-9d70-aee3-2e202687459e%28Office.15%29.aspx).


## Methods



|**Name**|
|:-----|
|[AutoFormat](http://msdn.microsoft.com/library/09d53c59-bcd7-6d7a-cc48-9b50017a5912%28Office.15%29.aspx)|
|[Calculate](http://msdn.microsoft.com/library/756d6143-bf92-7669-f686-be23246c3a29%28Office.15%29.aspx)|
|[CheckGrammar](http://msdn.microsoft.com/library/3ae0e80f-0165-be96-af12-b231d1f3a1b4%28Office.15%29.aspx)|
|[CheckSpelling](http://msdn.microsoft.com/library/41873962-8cac-84a4-4e01-712985513cd4%28Office.15%29.aspx)|
|[CheckSynonyms](http://msdn.microsoft.com/library/e28026bf-aa5e-8cf4-e765-7350afd57741%28Office.15%29.aspx)|
|[Collapse](http://msdn.microsoft.com/library/fa5cae70-f047-e300-52f7-bd75d9c613da%28Office.15%29.aspx)|
|[ComputeStatistics](http://msdn.microsoft.com/library/5fbeeffd-f592-3078-cd5b-1e2a90ee5092%28Office.15%29.aspx)|
|[ConvertHangulAndHanja](http://msdn.microsoft.com/library/2b640faf-da3c-a3b6-976b-d7dca3cb710f%28Office.15%29.aspx)|
|[ConvertToTable](http://msdn.microsoft.com/library/a7d005ec-774e-151c-ff38-64df3ea36646%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/c13c5310-cad2-c520-7304-507b81112551%28Office.15%29.aspx)|
|[CopyAsPicture](http://msdn.microsoft.com/library/b104bb78-9e76-37c7-2102-f71a3d8ddabb%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/51d28896-7552-d90c-5280-e8c8f0203f64%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/066b6dda-db9c-43aa-b65c-556b06b5b445%28Office.15%29.aspx)|
|[DetectLanguage](http://msdn.microsoft.com/library/4b4149fa-011a-2489-8779-66d75897174f%28Office.15%29.aspx)|
|[EndOf](http://msdn.microsoft.com/library/b9bda3b3-fee5-6359-f0ab-10fbe6b635b1%28Office.15%29.aspx)|
|[Expand](http://msdn.microsoft.com/library/cf4a5705-ebda-fedb-4929-3e115d42a432%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/d1cf9c7d-f2f3-1962-eccf-262568a56ad9%28Office.15%29.aspx)|
|[ExportFragment](http://msdn.microsoft.com/library/85c72276-9118-4156-22f9-84d00e7746da%28Office.15%29.aspx)|
|[GetSpellingSuggestions](http://msdn.microsoft.com/library/5ab65e3e-65d8-4e49-2874-609b1974888e%28Office.15%29.aspx)|
|[GoTo](http://msdn.microsoft.com/library/9e7cdfcc-756c-4bc8-902e-12479388ea03%28Office.15%29.aspx)|
|[GoToEditableRange](http://msdn.microsoft.com/library/4901bcef-56a7-c00e-409e-da0d442344c6%28Office.15%29.aspx)|
|[GoToNext](http://msdn.microsoft.com/library/011de2d6-c0fc-608f-8d7e-faac5947978d%28Office.15%29.aspx)|
|[GoToPrevious](http://msdn.microsoft.com/library/b1a6d089-c36a-1e10-fd8e-090d5b736a88%28Office.15%29.aspx)|
|[ImportFragment](http://msdn.microsoft.com/library/d9feca50-6370-c1c2-00c0-e64ff7a5adb9%28Office.15%29.aspx)|
|[InRange](http://msdn.microsoft.com/library/8d6b2093-7720-b100-6e9e-6be761cabaf5%28Office.15%29.aspx)|
|[InsertAfter](http://msdn.microsoft.com/library/25b2c0be-e9c7-1e42-09ea-308bbdcde7c6%28Office.15%29.aspx)|
|[InsertAlignmentTab](http://msdn.microsoft.com/library/1ca21f95-ca53-e911-c789-b0203d7bf0c7%28Office.15%29.aspx)|
|[InsertAutoText](http://msdn.microsoft.com/library/d87ae18c-e527-bcf4-4939-5512a6fdaaf5%28Office.15%29.aspx)|
|[InsertBefore](http://msdn.microsoft.com/library/ac77dcf7-ffcd-b109-8e17-ea6db169e85a%28Office.15%29.aspx)|
|[InsertBreak](http://msdn.microsoft.com/library/9c565036-e060-f26e-2e12-9c340331233e%28Office.15%29.aspx)|
|[InsertCaption](http://msdn.microsoft.com/library/fee41e81-1a78-2886-9693-dcf90da7c1bc%28Office.15%29.aspx)|
|[InsertCrossReference](http://msdn.microsoft.com/library/5899db5b-254c-17ac-4c4b-943a5a5b44cb%28Office.15%29.aspx)|
|[InsertDatabase](http://msdn.microsoft.com/library/c8bcddda-0943-9619-e5ee-9ef0956b0f43%28Office.15%29.aspx)|
|[InsertDateTime](http://msdn.microsoft.com/library/2203a0bb-6c90-ee55-6bdc-73f6761e4603%28Office.15%29.aspx)|
|[InsertFile](http://msdn.microsoft.com/library/9f35bacd-1cf3-42a4-c8ab-8c1cf183d2ab%28Office.15%29.aspx)|
|[InsertParagraph](http://msdn.microsoft.com/library/5686967c-38c3-6664-70ee-53937fbd920e%28Office.15%29.aspx)|
|[InsertParagraphAfter](http://msdn.microsoft.com/library/87c0a373-e066-5e53-7b50-e059a1a81b7b%28Office.15%29.aspx)|
|[InsertParagraphBefore](http://msdn.microsoft.com/library/78d62099-fa2c-911d-690b-93a9ee4f58eb%28Office.15%29.aspx)|
|[InsertSymbol](http://msdn.microsoft.com/library/2fa843fa-4966-a4e6-1411-028b14029bdf%28Office.15%29.aspx)|
|[InsertXML](http://msdn.microsoft.com/library/daee0fee-01cb-5ad7-f61d-ea6ebec1d04a%28Office.15%29.aspx)|
|[InStory](http://msdn.microsoft.com/library/62452309-4d4a-5207-3e1b-28b109ca1b1e%28Office.15%29.aspx)|
|[IsEqual](http://msdn.microsoft.com/library/cd6269d9-4693-897d-d9b2-69f45c815ba3%28Office.15%29.aspx)|
|[LookupNameProperties](http://msdn.microsoft.com/library/a3a0facf-898a-d8c9-033a-b48416b53266%28Office.15%29.aspx)|
|[ModifyEnclosure](http://msdn.microsoft.com/library/173c5b41-5245-4fc5-b9d9-9fd7cea0aab8%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/40c73c63-12da-4e8c-05c3-121f4df57f3f%28Office.15%29.aspx)|
|[MoveEnd](http://msdn.microsoft.com/library/44aa26e6-7bb1-af51-8d23-244444e0795c%28Office.15%29.aspx)|
|[MoveEndUntil](http://msdn.microsoft.com/library/62ac37a2-1116-73de-dcd8-0ff74ae7803b%28Office.15%29.aspx)|
|[MoveEndWhile](http://msdn.microsoft.com/library/9fab0517-a66a-2ae3-1900-77047b61cafa%28Office.15%29.aspx)|
|[MoveStart](http://msdn.microsoft.com/library/9097c636-594d-8a2e-8209-dc0db850812a%28Office.15%29.aspx)|
|[MoveStartUntil](http://msdn.microsoft.com/library/2506e3ec-593c-27ba-69b0-230351094f64%28Office.15%29.aspx)|
|[MoveStartWhile](http://msdn.microsoft.com/library/d0cff673-9248-88ae-7624-a838ce104e4b%28Office.15%29.aspx)|
|[MoveUntil](http://msdn.microsoft.com/library/f0f44ae5-1d61-9e05-4095-a28091feda6f%28Office.15%29.aspx)|
|[MoveWhile](http://msdn.microsoft.com/library/282464eb-60e6-df03-344f-6e666af8b01f%28Office.15%29.aspx)|
|[Next](http://msdn.microsoft.com/library/8d3a295d-543c-7e17-337d-b4fdfeda96e6%28Office.15%29.aspx)|
|[NextSubdocument](http://msdn.microsoft.com/library/4c048cc7-a2f6-38b1-e675-4d8870947130%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/06621016-de31-c61b-a9d0-6544b2d7e0a4%28Office.15%29.aspx)|
|[PasteAndFormat](http://msdn.microsoft.com/library/39dd8d10-0ab7-10d3-9e48-39a5e342553d%28Office.15%29.aspx)|
|[PasteAppendTable](http://msdn.microsoft.com/library/dc3b9914-b0d6-aa85-a357-a96475680caf%28Office.15%29.aspx)|
|[PasteAsNestedTable](http://msdn.microsoft.com/library/8d7a3fc6-5fc2-9cbc-d551-b4606af54619%28Office.15%29.aspx)|
|[PasteExcelTable](http://msdn.microsoft.com/library/2f682b61-6980-4287-5512-6cef62390b70%28Office.15%29.aspx)|
|[PasteSpecial](http://msdn.microsoft.com/library/76d074ee-f0d8-8bdd-e7c2-d0aa7b5f6702%28Office.15%29.aspx)|
|[PhoneticGuide](http://msdn.microsoft.com/library/f720cf42-4d61-977c-8e09-6346a48afecf%28Office.15%29.aspx)|
|[Previous](http://msdn.microsoft.com/library/ee1135ec-6f88-ec52-c3cc-0fb8183ac4cd%28Office.15%29.aspx)|
|[PreviousSubdocument](http://msdn.microsoft.com/library/542149f4-1a0c-bf1b-1cf6-9e8097af321e%28Office.15%29.aspx)|
|[Relocate](http://msdn.microsoft.com/library/2df77535-627f-d8ba-6ea2-15676b24221c%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/732c2aca-d8b4-3537-984f-d44d4eed870a%28Office.15%29.aspx)|
|[SetListLevel](http://msdn.microsoft.com/library/80cce7e2-49d1-614d-eb61-543d42aa5645%28Office.15%29.aspx)|
|[SetRange](http://msdn.microsoft.com/library/91097079-406c-98f4-d37c-cca8dab7aef0%28Office.15%29.aspx)|
|[Sort](http://msdn.microsoft.com/library/2030f99e-0307-d2b7-9e14-1d0888f3fda6%28Office.15%29.aspx)|
|[SortAscending](http://msdn.microsoft.com/library/2e7cd40d-6ddd-c191-c082-1e5c852e80a7%28Office.15%29.aspx)|
|[SortByHeadings](http://msdn.microsoft.com/library/8fd2b026-4744-7dad-7d68-06768ce4c35c%28Office.15%29.aspx)|
|[SortDescending](http://msdn.microsoft.com/library/018f7566-29cb-ad7f-87ae-55f041ab72a1%28Office.15%29.aspx)|
|[StartOf](http://msdn.microsoft.com/library/4d8d5a97-cb5e-cb27-c6dc-35c96c840ae1%28Office.15%29.aspx)|
|[TCSCConverter](http://msdn.microsoft.com/library/71684cdd-fca8-37b7-04fe-eeeb35dcfe66%28Office.15%29.aspx)|
|[WholeStory](http://msdn.microsoft.com/library/bb55c363-b3c0-e1aa-5e25-74cf2a1954c8%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/5777ab91-1ebf-3549-1b03-e54ab5f57f18%28Office.15%29.aspx)|
|[Bold](http://msdn.microsoft.com/library/04723b36-43bb-4721-90a5-33447a9b742e%28Office.15%29.aspx)|
|[BoldBi](http://msdn.microsoft.com/library/80a4e893-0337-41ef-5a45-506deea43f29%28Office.15%29.aspx)|
|[BookmarkID](http://msdn.microsoft.com/library/11157160-6cd5-38d7-dc92-be14399509f4%28Office.15%29.aspx)|
|[Bookmarks](http://msdn.microsoft.com/library/4a8d133a-fe6f-50ac-4b71-5265a919f5f1%28Office.15%29.aspx)|
|[Borders](http://msdn.microsoft.com/library/a09b85ab-4075-533b-5aa4-8cb7d10e436d%28Office.15%29.aspx)|
|[Case](http://msdn.microsoft.com/library/983f7bd3-10b4-882f-5b4d-01e44127676f%28Office.15%29.aspx)|
|[Cells](http://msdn.microsoft.com/library/aa081698-53d0-2234-5ec3-6e9a4091caef%28Office.15%29.aspx)|
|[Characters](http://msdn.microsoft.com/library/0d6ffe59-14ef-a198-e70f-6ccef0a83398%28Office.15%29.aspx)|
|[CharacterStyle](http://msdn.microsoft.com/library/22b57138-4e16-d144-9246-18b94ce463e7%28Office.15%29.aspx)|
|[CharacterWidth](http://msdn.microsoft.com/library/83eadb2b-5c79-d246-d1f1-fd6a9e1f4bd8%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/667b808a-e885-a7b7-0a68-5b2466ddd869%28Office.15%29.aspx)|
|[CombineCharacters](http://msdn.microsoft.com/library/4852ebb7-b6cc-0bed-d1db-8a2efe14fc17%28Office.15%29.aspx)|
|[Comments](http://msdn.microsoft.com/library/1fe73a8e-7341-e85c-5a72-daadfd3b0b22%28Office.15%29.aspx)|
|[Conflicts](http://msdn.microsoft.com/library/908b36ff-a87a-255c-2b5d-e47dd6489bf7%28Office.15%29.aspx)|
|[ContentControls](http://msdn.microsoft.com/library/e8c715af-067f-871e-7dec-28aa4302d9f9%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/df19ebef-edb2-ac75-e878-bce9e35794b3%28Office.15%29.aspx)|
|[DisableCharacterSpaceGrid](http://msdn.microsoft.com/library/042fcf3e-f163-0da2-9e05-8111b4353ace%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/400f7e50-6b7a-363b-77fc-e6fce215c002%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/bee676c8-cbc5-eaf9-0248-ad6098ce3c7f%28Office.15%29.aspx)|
|[Editors](http://msdn.microsoft.com/library/fe491d3f-e559-aa3d-53ce-bf4aea0de5f8%28Office.15%29.aspx)|
|[EmphasisMark](http://msdn.microsoft.com/library/6f0f7d19-efba-8fee-7e6c-abb1defe8529%28Office.15%29.aspx)|
|[End](http://msdn.microsoft.com/library/fe90f321-c7b5-bea2-fa60-e6b750b33cf7%28Office.15%29.aspx)|
|[EndnoteOptions](http://msdn.microsoft.com/library/48b2cf9e-edba-e6ed-a3b5-d93e26e17fe5%28Office.15%29.aspx)|
|[Endnotes](http://msdn.microsoft.com/library/311f3c49-bfdc-02e3-fbd9-a0f6614612b3%28Office.15%29.aspx)|
|[EnhMetaFileBits](http://msdn.microsoft.com/library/1e43483a-fb1d-5855-ec42-047f9bc9ef44%28Office.15%29.aspx)|
|[Fields](http://msdn.microsoft.com/library/106c1cb4-0836-3ff3-3138-223356a4a42c%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/53c18061-5fb8-d331-33ff-5b81b628d509%28Office.15%29.aspx)|
|[FitTextWidth](http://msdn.microsoft.com/library/6322c657-21db-bc45-e2d6-cb559edfc047%28Office.15%29.aspx)|
|[Font](http://msdn.microsoft.com/library/7582a7ed-0f16-e8f3-73f7-5d7b91193679%28Office.15%29.aspx)|
|[FootnoteOptions](http://msdn.microsoft.com/library/4adc72b6-cf26-8029-8c72-d2eed6583c27%28Office.15%29.aspx)|
|[Footnotes](http://msdn.microsoft.com/library/5c06672f-8de2-29e9-dd68-15408348faa5%28Office.15%29.aspx)|
|[FormattedText](http://msdn.microsoft.com/library/26221da8-e3d7-4da5-f23a-cd678d8ab2f5%28Office.15%29.aspx)|
|[FormFields](http://msdn.microsoft.com/library/9777dc22-1fe5-c442-a4bf-e3dae4549168%28Office.15%29.aspx)|
|[Frames](http://msdn.microsoft.com/library/c30bb71d-3998-42fe-2850-a76c3975418b%28Office.15%29.aspx)|
|[GrammarChecked](http://msdn.microsoft.com/library/f10af296-28f0-dd4b-fdab-70bad8d3e924%28Office.15%29.aspx)|
|[GrammaticalErrors](http://msdn.microsoft.com/library/2535ba4d-1c5c-3dc2-2ddc-14c8a5625f41%28Office.15%29.aspx)|
|[HighlightColorIndex](http://msdn.microsoft.com/library/ff6e0f1a-8b37-1bdd-8da6-ac492d399ad2%28Office.15%29.aspx)|
|[HorizontalInVertical](http://msdn.microsoft.com/library/1d0ec26c-62a1-26ef-1fef-f2ab497244cb%28Office.15%29.aspx)|
|[HTMLDivisions](http://msdn.microsoft.com/library/4712d81f-7028-357b-a7ff-dc4f382cc5e3%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/c8eb84af-b090-82ee-8001-b251c6cc1f24%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/5b2145f3-b21f-5550-e058-9c81ccdaa0e3%28Office.15%29.aspx)|
|[Information](http://msdn.microsoft.com/library/967e9a22-5f98-e4bd-557c-7367cb7c5d2b%28Office.15%29.aspx)|
|[InlineShapes](http://msdn.microsoft.com/library/4c0335ac-95a2-412c-650c-afc323ae58ca%28Office.15%29.aspx)|
|[IsEndOfRowMark](http://msdn.microsoft.com/library/0b1a7638-75ea-fb03-3a52-8bc759794408%28Office.15%29.aspx)|
|[Italic](http://msdn.microsoft.com/library/7d52781a-46f2-7bca-067e-dc41772149fc%28Office.15%29.aspx)|
|[ItalicBi](http://msdn.microsoft.com/library/69f2ace2-0e12-b704-531c-e4d769d738ec%28Office.15%29.aspx)|
|[Kana](http://msdn.microsoft.com/library/ed64b73e-6970-3099-6f75-0beac6bba84e%28Office.15%29.aspx)|
|[LanguageDetected](http://msdn.microsoft.com/library/dfe307e5-ad87-9a6b-ecbe-521c6354b349%28Office.15%29.aspx)|
|[LanguageID](http://msdn.microsoft.com/library/dc163c7b-8a44-4b8a-5674-845984f1b682%28Office.15%29.aspx)|
|[LanguageIDFarEast](http://msdn.microsoft.com/library/324eaba2-2a48-71e3-6a96-9b7a092d0c6d%28Office.15%29.aspx)|
|[LanguageIDOther](http://msdn.microsoft.com/library/00b07195-df7d-a979-2534-370cf6540c79%28Office.15%29.aspx)|
|[ListFormat](http://msdn.microsoft.com/library/509365dc-0b93-96d9-6614-74f2d85bfd45%28Office.15%29.aspx)|
|[ListParagraphs](http://msdn.microsoft.com/library/d581249c-1f63-9043-8d8c-32b0e3bb2a5c%28Office.15%29.aspx)|
|[ListStyle](http://msdn.microsoft.com/library/5bbeaeab-5dfa-6c3e-ba42-fb0af2940674%28Office.15%29.aspx)|
|[Locks](http://msdn.microsoft.com/library/102673f2-8cb0-d235-c158-c65759592d56%28Office.15%29.aspx)|
|[NextStoryRange](http://msdn.microsoft.com/library/392b17ff-335f-9b2b-7641-62ae44d7e919%28Office.15%29.aspx)|
|[NoProofing](http://msdn.microsoft.com/library/0344239d-10bc-0e3e-9601-41c3c3bb6227%28Office.15%29.aspx)|
|[OMaths](http://msdn.microsoft.com/library/8721e30e-b36d-e216-1f52-304d0b8737f7%28Office.15%29.aspx)|
|[Orientation](http://msdn.microsoft.com/library/52fde061-7ae9-61a4-c66d-7ffe691e1f97%28Office.15%29.aspx)|
|[PageSetup](http://msdn.microsoft.com/library/29a1d7cb-42dd-3d3b-1cb6-7905987f962f%28Office.15%29.aspx)|
|[ParagraphFormat](http://msdn.microsoft.com/library/98afe866-4d92-7a1d-f5c6-a0128d247df0%28Office.15%29.aspx)|
|[Paragraphs](http://msdn.microsoft.com/library/b5c9df62-a477-ce1a-4a94-027100527a6f%28Office.15%29.aspx)|
|[ParagraphStyle](http://msdn.microsoft.com/library/55bfbbe2-1e17-e37b-8010-9142fe080e1f%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/8c644100-7aa7-fccd-95c3-2aba0389b47d%28Office.15%29.aspx)|
|[ParentContentControl](http://msdn.microsoft.com/library/e5805628-bfe0-64a6-78c4-f008098450d1%28Office.15%29.aspx)|
|[PreviousBookmarkID](http://msdn.microsoft.com/library/19aab6c4-bc86-3f65-4fbc-206fdf3dbb3a%28Office.15%29.aspx)|
|[ReadabilityStatistics](http://msdn.microsoft.com/library/c0dcf3e8-2c1a-3d23-48e9-4dfcd0d75893%28Office.15%29.aspx)|
|[Revisions](http://msdn.microsoft.com/library/cf71b684-991a-fb6d-09bc-eeecb16edec5%28Office.15%29.aspx)|
|[Rows](http://msdn.microsoft.com/library/fd2c7ecd-07de-c25f-4a51-4a14abad9951%28Office.15%29.aspx)|
|[Scripts](http://msdn.microsoft.com/library/233acf3a-3151-f4f2-e5df-815edeca1dd1%28Office.15%29.aspx)|
|[Sections](http://msdn.microsoft.com/library/98340968-d810-1e9c-0989-c1d03e614c14%28Office.15%29.aspx)|
|[Sentences](http://msdn.microsoft.com/library/fe870f13-d09f-efbf-1d2f-745f2c318c28%28Office.15%29.aspx)|
|[Shading](http://msdn.microsoft.com/library/8e09cd74-a16e-6547-5ada-97322cf32b99%28Office.15%29.aspx)|
|[ShapeRange](http://msdn.microsoft.com/library/b8e6e1f7-d29a-5fb5-8d00-22b3907d6f54%28Office.15%29.aspx)|
|[ShowAll](http://msdn.microsoft.com/library/751077ec-5ea4-c60a-ac92-d8a5a3c13620%28Office.15%29.aspx)|
|[SpellingChecked](http://msdn.microsoft.com/library/5a58fb94-186b-d30c-bef4-d42a295fdeb6%28Office.15%29.aspx)|
|[SpellingErrors](http://msdn.microsoft.com/library/4b35a13d-2a5f-e9cd-0667-58aae00a48f1%28Office.15%29.aspx)|
|[Start](http://msdn.microsoft.com/library/aadedbb7-1ee2-9e5a-296d-0ebe25b6d8f4%28Office.15%29.aspx)|
|[StoryLength](http://msdn.microsoft.com/library/0dd342e2-2a90-bbf9-2989-a2629fcf40a5%28Office.15%29.aspx)|
|[StoryType](http://msdn.microsoft.com/library/bf11ba94-de45-ae76-09fa-9463cd2c4723%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/aeceef42-cbdc-3d55-2f43-0afffd933cc2%28Office.15%29.aspx)|
|[Subdocuments](http://msdn.microsoft.com/library/c06afeb9-7e83-d858-d863-9582962c8254%28Office.15%29.aspx)|
|[SynonymInfo](http://msdn.microsoft.com/library/b63d2a0b-baa1-306d-10ee-72223099a9f2%28Office.15%29.aspx)|
|[Tables](http://msdn.microsoft.com/library/1c6604be-233c-efb2-5d05-63fc5aa78481%28Office.15%29.aspx)|
|[TableStyle](http://msdn.microsoft.com/library/ff392d59-eb86-7ba3-c811-67090fe9889f%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/495fe06e-ba87-0d96-9f6e-3e62fd71d4a5%28Office.15%29.aspx)|
|[TextRetrievalMode](http://msdn.microsoft.com/library/e3992479-ba69-e8d3-17e3-73b533f27d26%28Office.15%29.aspx)|
|[TextVisibleOnScreen](http://msdn.microsoft.com/library/ced8fc7c-61a2-b0dd-20ba-ee6a4281d44d%28Office.15%29.aspx)|
|[TopLevelTables](http://msdn.microsoft.com/library/43cd13b8-f779-69cd-ee60-d4ba734008f0%28Office.15%29.aspx)|
|[TwoLinesInOne](http://msdn.microsoft.com/library/08e91e95-4826-7df9-22a9-3c7b9c25042d%28Office.15%29.aspx)|
|[Underline](http://msdn.microsoft.com/library/8221338d-3da6-b1ae-c424-87f762b61bd7%28Office.15%29.aspx)|
|[Updates](http://msdn.microsoft.com/library/584c9a40-0975-75d9-e3d4-32e857fb62e5%28Office.15%29.aspx)|
|[WordOpenXML](http://msdn.microsoft.com/library/bb4aa9c3-dd69-e27f-9c72-4dc4795fbd26%28Office.15%29.aspx)|
|[Words](http://msdn.microsoft.com/library/ada98916-b87c-7592-ee2d-561ed7067f39%28Office.15%29.aspx)|
|[XML](http://msdn.microsoft.com/library/3752b0bc-0f9c-d5ca-1926-763db9d1b1cc%28Office.15%29.aspx)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

