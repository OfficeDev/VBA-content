---
title: Document.CreateLetterContent Method (Word)
keywords: vbawd10.chm158007556
f1_keywords:
- vbawd10.chm158007556
ms.prod: word
api_name:
- Word.Document.CreateLetterContent
ms.assetid: 33f47344-31d2-4099-45fc-91af2d79dc7c
ms.date: 06/08/2017
---


# Document.CreateLetterContent Method (Word)

Creates and returns a  **LetterContent** object based on the specified letter elements. **LetterContent** object.


## Syntax

 _expression_ . **CreateLetterContent**( **_DateFormat_** , **_IncludeHeaderFooter_** , **_PageDesign_** , **_LetterStyle_** , **_Letterhead_** , **_LetterheadLocation_** , **_LetterheadSize_** , **_RecipientName_** , **_RecipientAddress_** , **_Salutation_** , **_SalutationType_** , **_RecipientReference_** , **_MailingInstructions_** , **_AttentionLine_** , **_Subject_** , **_CCList_** , **_ReturnAddress_** , **_SenderName_** , **_Closing_** , **_SenderCompany_** , **_SenderJobTitle_** , **_SenderInitials_** , **_EnclosureNumber_** , **_InfoBlock_** , **_RecipientCode_** , **_RecipientGender_** , **_ReturnAddressShortForm_** , **_SenderCity_** , **_SenderCode_** , **_SenderGender_** , **_SenderReference_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DateFormat_|Required| **String**|The date for the letter.|
| _IncludeHeaderFooter_|Required| **Boolean**| **True** to include the header and footer from the page design template.|
| _PageDesign_|Required| **String**|The name of the template attached to the document.|
| _LetterStyle_|Required| **WdLetterStyle**|The document layout.|
| _Letterhead_|Required| **Boolean**| **True** to reserve space for a preprinted letterhead.|
| _LetterheadLocation_|Required| **WdLetterheadLocation**|The location of the preprinted letterhead.|
| _LetterheadSize_|Required| **Single**|The amount of space (in points) to be reserved for a preprinted letterhead.|
| _RecipientName_|Required| **String**|The name of the person who'll be receiving the letter.|
| _RecipientAddress_|Required| **String**|The mailing address of the person who'll be receiving the letter.|
| _Salutation_|Required| **String**|The salutation text for the letter.|
| _SalutationType_|Required| **WdSalutationType**|The salutation type for the letter.|
| _RecipientReference_|Required| **String**|The reference line text for the letter (for example, "In reply to:").|
| _MailingInstructions_|Required| **String**|The mailing instruction text for the letter (for example, "Certified Mail").|
| _AttentionLine_|Required| **String**|The attention line text for the letter (for example, "Attention:").|
| _Subject_|Required| **String**|The subject text for the specified letter.|
| _CCList_|Required| **String**|The names of the carbon copy (CC) recipients for the letter.|
| _ReturnAddress_|Required| **String**|The text of the return mailing address for the letter.|
| _SenderName_|Required| **String**|The name of the person sending the letter.|
| _Closing_|Required| **String**|The closing text for the letter.|
| _SenderCompany_|Required| **String**|The company name of the person creating the letter.|
| _SenderJobTitle_|Required| **String**|The job title of the person creating the letter.|
| _SenderInitials_|Required| **String**|The initials of the person creating the letter.|
| _EnclosureNumber_|Required| **Long**|The number of enclosures for the letter.|
| _InfoBlock_|Optional| **Variant**|This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _RecipientCode_|Optional| **Variant**|This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _RecipientGender_|Optional| **Variant**|This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _ReturnAddressShortForm_|Optional| **Variant**|This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _SenderCity_|Optional| **Variant**|This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _SenderCode_|Optional| **Variant**|This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _SenderGender_|Optional| **Variant**|This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _SenderReference_|Optional| **Variant**|This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|

### Return Value

LetterContent


## Example

The following example uses the  **CreateLetterContent** method to create a new **LetterContent** object in the active document and then uses this object with the **RunLetterWizard** method.


```vb
Set myLetter = ActiveDocument _ 
 .CreateLetterContent(DateFormat:="July 31, 1996", _ 
 IncludeHeaderFooter:=False, PageDesign:="", _ 
 LetterStyle:=wdFullBlock, Letterhead:=True, _ 
 LetterheadLocation:=wdLetterTop, _ 
 LetterheadSize:=InchesToPoints(1.5), _ 
 RecipientName:="Dave Edson", _ 
 RecipientAddress:="436 SE Main St." &; vbCr _ 
 &; "Bellevue, WA 98004", _ 
 Salutation:="Dear Dave,", _ 
 SalutationType:=wdSalutationInformal, _ 
 RecipientReference:="", MailingInstructions:="", _ 
 AttentionLine:="", Subject:="End of year report", _ 
 CCList:="", ReturnAddress:="", _ 
 SenderName:="", Closing:="Sincerely yours,", _ 
 SenderCompany:="", SenderJobTitle:="", _ 
 SenderInitials:="", EnclosureNumber:=0) 
ActiveDocument.RunLetterWizard LetterContent:=myLetter
```


## See also


#### Concepts


[Document Object](document-object-word.md)

