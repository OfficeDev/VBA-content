---
title: WdFieldType Enumeration (Word)
ms.prod: word
api_name:
- Word.WdFieldType
ms.assetid: 220d280c-0ff4-080c-4273-e5c8c437333f
ms.date: 06/08/2017
---


# WdFieldType Enumeration (Word)

Specifies a Microsoft Word field. Unless otherwise specified, the field types described in this enumeration can be added interactively to a Word document by using the  **Field** dialog box. See the Word Help for more information about specific field codes.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **wdFieldAddin**|81|Add-in field. Not available through the  **Field** dialog box. Used to store data that is hidden from the user interface.|
| **wdFieldAddressBlock**|93|AddressBlock field.|
| **wdFieldAdvance**|84|Advance field.|
| **wdFieldAsk**|38|Ask field.|
| **wdFieldAuthor**|17|Author field.|
| **wdFieldAutoNum**|54|AutoNum field.|
| **wdFieldAutoNumLegal**|53|AutoNumLgl field.|
| **wdFieldAutoNumOutline**|52|AutoNumOut field.|
| **wdFieldAutoText**|79|AutoText field.|
| **wdFieldAutoTextList**|89|AutoTextList field.|
| **wdFieldBarCode**|63|BarCode field.|
| **wdFieldBidiOutline**|92|BidiOutline field. |
| **wdFieldComments**|19|Comments field.|
| **wdFieldCompare**|80|Compare field.|
| **wdFieldCreateDate**|21|CreateDate field.|
| **wdFieldData**|40|Data field.|
| **wdFieldDatabase**|78|Database field.|
| **wdFieldDate**|31|Date field.|
| **wdFieldDDE**|45|DDE field. No longer available through the  **Field** dialog box, but supported for documents created in earlier versions of Word.|
| **wdFieldDDEAuto**|46|DDEAuto field. No longer available through the  **Field** dialog box, but supported for documents created in earlier versions of Word.|
| **wdFieldDisplayBarcode**|99|DisplayBarcode field.|
| **wdFieldDocProperty**|85|DocProperty field.|
| **wdFieldDocVariable**|64|DocVariable field.|
| **wdFieldEditTime**|25|EditTime field.|
| **wdFieldEmbed**|58|Embedded field.|
| **wdFieldEmpty**|-1|Empty field. Acts as a placeholder for field content that has not yet been added. A field added by pressing Ctrl+F9 in the user interface is an Empty field.|
| **wdFieldExpression**|34|= (Formula) field.|
| **wdFieldFileName**|29|FileName field.|
| **wdFieldFileSize**|69|FileSize field.|
| **wdFieldFillIn**|39|Fill-In field.|
| **wdFieldFootnoteRef**|5|FootnoteRef field. Not available through the  **Field** dialog box. Inserted programmatically or interactively.|
| **wdFieldFormCheckBox**|71|FormCheckBox field. |
| **wdFieldFormDropDown**|83|FormDropDown field. |
| **wdFieldFormTextInput**|70|FormText field. |
| **wdFieldFormula**|49|EQ (Equation) field.|
| **wdFieldGlossary**|47|Glossary field. No longer supported in Word.|
| **wdFieldGoToButton**|50|GoToButton field.|
| **wdFieldGreetingLine**|94|GreetingLine field.|
| **wdFieldHTMLActiveX**|91|HTMLActiveX field. Not currently supported.|
| **wdFieldHyperlink**|88|Hyperlink field.|
| **wdFieldIf**|7|If field.|
| **wdFieldImport**|55|Import field. Cannot be added through the  **Field** dialog box, but can be added interactively or through code.|
| **wdFieldInclude**|36|Include field. Cannot be added through the  **Field** dialog box, but can be added interactively or through code.|
| **wdFieldIncludePicture**|67|IncludePicture field.|
| **wdFieldIncludeText**|68|IncludeText field.|
| **wdFieldIndex**|8|Index field.|
| **wdFieldIndexEntry**|4|XE (Index Entry) field.|
| **wdFieldInfo**|14|Info field.|
| **wdFieldKeyWord**|18|Keywords field.|
| **wdFieldLastSavedBy**|20|LastSavedBy field.|
| **wdFieldLink**|56|Link field.|
| **wdFieldListNum**|90|ListNum field.|
| **wdFieldMacroButton**|51|MacroButton field.|
| **wdFieldMergeBarcode**|98|MergeBarcode field.|
| **wdFieldMergeField**|59|MergeField field.|
| **wdFieldMergeRec**|44|MergeRec field.|
| **wdFieldMergeSeq**|75|MergeSeq field.|
| **wdFieldNext**|41|Next field.|
| **wdFieldNextIf**|42|NextIf field.|
| **wdFieldNoteRef**|72|NoteRef field.|
| **wdFieldNumChars**|28|NumChars field.|
| **wdFieldNumPages**|26|NumPages field.|
| **wdFieldNumWords**|27|NumWords field.|
| **wdFieldOCX**|87|OCX field. Cannot be added through the  **Field** dialog box, but can be added through code by using the **AddOLEControl** method of the **[Shapes](shapes-object-word.md)** collection or of the **[InlineShapes](inlineshapes-object-word.md)** collection.|
| **wdFieldPage**|33|Page field.|
| **wdFieldPageRef**|37|PageRef field.|
| **wdFieldPrint**|48|Print field.|
| **wdFieldPrintDate**|23|PrintDate field.|
| **wdFieldPrivate**|77|Private field.|
| **wdFieldQuote**|35|Quote field.|
| **wdFieldRef**|3|Ref field.|
| **wdFieldRefDoc**|11|RD (Reference Document) field.|
| **wdFieldRevisionNum**|24|RevNum field.|
| **wdFieldSaveDate**|22|SaveDate field.|
| **wdFieldSection**|65|Section field.|
| **wdFieldSectionPages**|66|SectionPages field.|
| **wdFieldSequence**|12|Seq (Sequence) field.|
| **wdFieldSet**|6|Set field.|
| **wdFieldShape**|95|Shape field. Automatically created for any drawn picture.|
| **wdFieldSkipIf**|43|SkipIf field.|
| **wdFieldStyleRef**|10|StyleRef field.|
| **wdFieldSubject**|16|Subject field.|
| **wdFieldSubscriber**|82|Macintosh only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.|
| **wdFieldSymbol**|57|Symbol field.|
| **wdFieldTemplate**|30|Template field.|
| **wdFieldTime**|32|Time field.|
| **wdFieldTitle**|15|Title field.|
| **wdFieldTOA**|73|TOA (Table of Authorities) field.|
| **wdFieldTOAEntry**|74|TOA (Table of Authorities Entry) field.|
| **wdFieldTOC**|13|TOC (Table of Contents) field.|
| **wdFieldTOCEntry**|9|TOC (Table of Contents Entry) field.|
| **wdFieldUserAddress**|62|UserAddress field.|
| **wdFieldUserInitials**|61|UserInitials field.|
| **wdFieldUserName**|60|UserName field.|
| **wdFieldBibliography**|97|Bibliography field.|
| **wdFieldCitation**|96|Citation field.|

