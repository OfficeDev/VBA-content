---
title: WdRecoveryType Enumeration (Word)
ms.prod: word
api_name:
- Word.WdRecoveryType
ms.assetid: 031525aa-6df9-2b28-8507-fa3c869beba8
ms.date: 06/08/2017
---


# WdRecoveryType Enumeration (Word)

Specifies the formatting to use when pasting the selected table cells.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **wdChart**|14|Pastes a Microsoft Office Excel chart as an embedded OLE object.|
| **wdChartLinked**|15|Pastes an Excel chart and links it to the original Excel spreadsheet.|
| **wdChartPicture**|13|Pastes an Excel chart as a picture.|
| **wdFormatOriginalFormatting**|16|Preserves original formatting of the pasted material.|
| **wdFormatPlainText**|22|Pastes as plain, unformatted text.|
| **wdFormatSurroundingFormattingWithEmphasis**|20|Matches the formatting of the pasted text to the formatting of surrounding text.|
| **wdListCombineWithExistingList**|24|Merges a pasted list with neighboring lists.|
| **wdListContinueNumbering**|7|Continues numbering of a pasted list from the list in the document.|
| **wdListDontMerge**|25|Not supported.|
| **wdListRestartNumbering**|8|Restarts numbering of a pasted list.|
| **wdPasteDefault**|0|Not supported.|
| **wdSingleCellTable**|6|Pastes a single cell table as a separate table.|
| **wdSingleCellText**|5|Pastes a single cell as text.|
| **wdTableAppendTable**|10|Merges pasted cells into an existing table by inserting the pasted rows between the selected rows.|
| **wdTableInsertAsRows**|11|Inserts a pasted table as rows between two rows in the target table.|
| **wdTableOriginalFormatting**|12|Pastes an appended table without merging table styles.|
| **wdTableOverwriteCells**|23|Pastes table cells and overwrites existing table cells.|
| **wdUseDestinationStylesRecovery**|19|Uses the styles that are in use in the destination document.|

