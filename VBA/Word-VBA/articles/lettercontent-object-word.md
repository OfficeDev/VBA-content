---
title: LetterContent Object (Word)
keywords: vbawd10.chm2465
f1_keywords:
- vbawd10.chm2465
ms.prod: word
api_name:
- Word.LetterContent
ms.assetid: 62a4e17a-6598-c904-f27d-817c19c04981
ms.date: 06/08/2017
---


# LetterContent Object (Word)

Represents the elements of a letter created by the Letter Wizard.


## Remarks

Use the  **GetLetterContent** or **CreateLetterContent** method to return a **LetterContent** object. The following example retrieves and displays the letter recipient's name from the active document.


```vb
Set myLetterContent = ActiveDocument.GetLetterContent 
MsgBox myLetterContent.RecipientName
```

The following example uses the  **CreateLetterContent** method to create a new **LetterContent** object, which is then used with the **RunLetterWizard** method.




```vb
Set myLetter = ActiveDocument _ 
 .CreateLetterContent(DateFormat:="July 11, 1996", _ 
 IncludeHeaderFooter:=False, _ 
 PageDesign:="C:\MSOffice\Templates\Letters &; " _ 
 &; "Faxes\Contemporary Letter.dot", _ 
 LetterStyle:=wdFullBlock, Letterhead:=True, _ 
 LetterheadLocation:=wdLetterTop, _ 
 LetterheadSize:=InchesToPoints(1.5), _ 
 RecipientName:="Dave Edson", _ 
 RecipientAddress:="100 Main St." &; vbCr _ 
 &; "Bellevue, WA 98004", _ 
 Salutation:="Dear Dave,", _ 
 SalutationType:=wdSalutationInformal, _ 
 RecipientReference:="", MailingInstructions:="", _ 
 AttentionLine:="", _ 
 Subject:="End of year report", CCList:="", ReturnAddress:="", _ 
 SenderName:="", Closing:="Sincerely yours,", _ 
 SenderCompany:="", _ 
 SenderJobTitle:="", SenderInitials:="", EnclosureNumber:=0) 
ActiveDocument.RunLetterWizard _ 
 LetterContent:=myLetter, WizardMode:=True
```

The  **CreateLetterContent** method creates a **LetterContent** object; however, there are numerous required arguments. If you want to set only a few properties, use the **New** keyword to create a new, stand-alone **LetterContent** object. The following example creates a **LetterContent** object, sets some of its properties, and then uses the **LetterContent** object with the **RunLetterWizard** method to run the Letter Wizard, using the preset values as the default settings.




```vb
Set myLetter = New LetterContent 
With myLetter 
 .AttentionLine = "Read this" 
 .EnclosureNumber = 1 
 .Letterhead = True 
 .LetterheadLocation = wdLetterTop 
 .LetterheadSize = InchesToPoints(2) 
End With 
Documents.Add.RunLetterWizard LetterContent:=myLetter, _ 
 WizardMode:=True
```

You can duplicate a  **LetterContent** object by using the **Duplicate** property. The following example retrieves the letter elements in the active document and makes a duplicate copy. The example assigns the duplicate copy to _aLetter_ and resets the recipient's name and address to empty strings. The **RunLetterWizard** method is used to run the Letter Wizard, using the values in the revised **LetterContent** object ( _aLetter_ ) as the default settings.




```vb
Set aLetter = ActiveDocument.GetLetterContent.Duplicate 
With aLetter 
 .RecipientName = "" 
 .RecipientAddress = "" 
End With 
Documents.Add.RunLetterWizard LetterContent:=aLetter, _ 
 WizardMode:=True
```

The  **SetLetterContent** method inserts the contents of the specified **LetterContent** object in a document. The following example retrieves the letter elements from the active document, changes the attention line, and then uses the **SetLetterContent** method to update the active document to reflect the change.




```vb
Set myLetterContent = ActiveDocument.GetLetterContent 
myLetterContent.AttentionLine = "Greetings" 
ActiveDocument.SetLetterContent LetterContent:=myLetterContent
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


