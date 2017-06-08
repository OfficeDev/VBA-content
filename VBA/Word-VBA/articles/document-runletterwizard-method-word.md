---
title: Document.RunLetterWizard Method (Word)
keywords: vbawd10.chm158007419
f1_keywords:
- vbawd10.chm158007419
ms.prod: word
api_name:
- Word.Document.RunLetterWizard
ms.assetid: 7da6e2b9-607a-0d3e-7d0d-762a8900a486
ms.date: 06/08/2017
---


# Document.RunLetterWizard Method (Word)

Runs the Letter Wizard on the specified document.


## Syntax

 _expression_ . **RunLetterWizard**( **_LetterContent_** , **_WizardMode_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _LetterContent_|Optional| **Variant**|A  **[LetterContent](lettercontent-object-word.md)** object. Any filled properties in the **LetterContent** object show up as prefilled elements in the Letter Wizard dialog boxes. If this argument is omitted, the **GetLetterContent** method is automatically used to get a **LetterContent** object from the specified document.|
| _WizardMode_|Optional| **Variant**| **True** to display the **Letter Wizard** dialog box as a series of steps with a **Next**,  **Back**, and  **Finish** button. **False** to display the **Letter Wizard** dialog box as if it were opened from the **Tools** menu (a properties dialog box with an **OK** button and a **Cancel** button). The default value is **True** .|

## Remarks

Use the  **CreateLetterContent** method to return a **LetterContent** object, given various letter element properties. Use the **GetLetterContent** method to return a **LetterContent** object based on the contents of the specified document. You can use the resulting **LetterContent** object with the **RunLetterWizard** method to preset elements in the **Letter Wizard** dialog box.


## Example

This example creates a new  **LetterContent** object, sets several properties for it, and then runs the Letter Wizard by using the **RunLetterWizard** method.


```vb
Set myContent = New LetterContent 
With myContent 
 .Salutation ="Hello" 
 .SalutationType = wdSalutationOther 
 .SenderName = Application.UserName 
 .SenderInitials =Application.UserInitials 
End With 
Documents.Add.RunLetterWizard _ 
 LetterContent:=myContent, WizardMode:=True
```

The following example uses the  **CreateLetterContent** method to create a new **LetterContent** object in the active document, and then it uses this object with the **RunLetterWizard** method.




```vb
Set myLetter = ActiveDocument _ 
 .CreateLetterContent(DateFormat:="July 31, 1999", _ 
 IncludeHeaderFooter:=False, _ 
 PageDesign:="C:\MSOffice\Templates" _ 
 &; "\Letters &; Faxes\Contemporary Letter.dot", _ 
 LetterStyle:=wdFullBlock, Letterhead:=True, _ 
 LetterheadLocation:=wdLetterTop, _ 
 LetterheadSize:=InchesToPoints(1.5), _ 
 RecipientName:="Dave Edson", _ 
 RecipientAddress:="436 SE Main St." _ 
 &; vbCr &; "Bellevue, WA 98004", _ 
 Salutation:="Dear Dave,", _ 
 SalutationType:=wdSalutationInformal, _ 
 RecipientReference:="", MailingInstructions:="", _ 
 AttentionLine:="", Subject:="End of year report", _ 
 CCList:="", ReturnAddress:="", SenderName:="", _ 
 Closing:="Sincerely yours,", SenderCompany:="", _ 
 SenderJobTitle:="", SenderInitials:="", _ 
 EnclosureNumber:=0) 
ActiveDocument.RunLetterWizard LetterContent:=myLetter
```


## See also


#### Concepts


[Document Object](document-object-word.md)

