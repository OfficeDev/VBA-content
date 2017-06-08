---
title: Application.CompareDocuments Method (Word)
keywords: vbawd10.chm158335446
f1_keywords:
- vbawd10.chm158335446
ms.prod: word
api_name:
- Word.Application.CompareDocuments
ms.assetid: 511c811f-3f2b-9b93-f339-32324569a765
ms.date: 06/08/2017
---


# Application.CompareDocuments Method (Word)

Compares two documents and returns a  **Document** object that represents the document that contains the differences between the two documents, marked using tracked changes.


## Syntax

 _expression_ . **CompareDocuments**( **_OriginalDocument_** , **_RevisedDocument_** , **_Destination_** , **_Granularity_** , **_CompareFormatting_** , **_CompareCaseChanges_** , **_CompareWhitespace_** , **_CompareTables_** , **_CompareHeaders_** , **_CompareFootnotes_** , **_CompareTextboxes_** , **_CompareFields_** , **_CompareComments_** , **_RevisedAuthor_** , **_IgnoreAllComparisonWarnings_** )

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _OriginalDocument_|Required| **Document**|Specifies the path and file name of the original document.|
| _RevisedDocument_|Required| **Document**|Specifies the path and file name of the revised document to which to compare the original document.|
| _Destination_|Optional| **[WdCompareDestination](wdcomparedestination-enumeration-word.md)**|Specifies whether to create a new file or whether to mark the differences between the two documents in the original document or in the revised document. Default value is  **wdCompareDestinationNew** .|
| _Granularity_|Optional| **[WdGranularity](wdgranularity-enumeration-word.md)**|Specifies whether changes are tracked by character or by word. Default value is  **wdGranularityWordLevel** .|
| _CompareFormatting_|Optional| **Boolean**|Specifies whether to mark differences in formatting between the two documents. Default value is  **True** .|
| _CompareCaseChanges_|Optional| **Boolean**|Specifies whether to mark differences in case between the two documents. Default value is  **True** .|
| _CompareWhitespace_|Optional| **Boolean**|Specifies whether to mark differences in white space, such as paragraphs or spaces, between the two documents. Default value is  **True** .|
| _CompareTables_|Optional| **Boolean**|Specifies whether to compare the differences in data contained in tables between the two documents. Default value is  **True** .|
| _CompareHeaders_|Optional| **Boolean**|Specifies whether to compare differences in headers and footers between the two documents. Default value is  **True** .|
| _CompareFootnotes_|Optional| **Boolean**|Specifies whether to compare differences in footnotes and endnotes between the two documents. Default value is  **True** .|
| _CompareTextboxes_|Optional| **Boolean**|Specifies whether to compare differences in the data contained within text boxes between the two documents. Default value is  **True** .|
| _CompareFields_|Optional| **Boolean**|Specifies whether to compare differences in fields between the two documents. Default value is  **True** .|
| _CompareComments_|Optional| **Boolean**|Specifies whether to compare differences in comments between the two documents. Default value is  **True** .|
| _RevisedAuthor_|Optional| **String**|Specifies the name of the person to whom to attribute changes when comparing the two documents.|
| _IgnoreAllComparisonWarnings_|Optional| **Boolean**|Specifies whether to ignore warnings when comparing the two documents.|

### Return Value

Document


## See also


#### Concepts


[Application Object](application-object-word.md)

