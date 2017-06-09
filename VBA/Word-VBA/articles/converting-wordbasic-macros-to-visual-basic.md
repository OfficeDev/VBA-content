---
title: Converting WordBasic Macros to Visual Basic
keywords: vbawd10.chm5210330
f1_keywords:
- vbawd10.chm5210330
ms.prod: word
ms.assetid: 44a08969-f0e9-291e-7663-b7cc2e3660db
ms.date: 06/08/2017
---


# Converting WordBasic Macros to Visual Basic

Word 2003 and Word 2007 automatically convert the macros in a Word 6.x or Word 95 template the first time you do any of the following:


- Open the template
    
- Create a document based on the template
    
- Manually attach the template to a document
    

A message is displayed on the status bar while the macros are being converted. After the conversion is complete, you must save the template to save the converted macros. If you don't save the template, Word converts the macros again the next time you use the template.


 **Note**  Word cannot convert Word 2.x macros directly. Instead, you need to open and save your Word 2.x templates in Word 6.x or Word 95 and then open them in Word.

The conversion process converts each macro to a Visual Basic module. To see the converted macros, press Alt-F8. The macro names in the  **Macros** dialog box appear as _macroname_.Main, where Main refers to the main subroutine in the converted macro (the subroutine that began with Sub MAIN in earlier versions of Word). To edit the converted macro, select a macro name and click  **Edit** to display the Visual Basic module in the Visual Basic Editor.
Each WordBasic statement is modified to work with Visual Basic for Applications. The converted WordBasic macros are functionally equivalent to new Visual Basic for Applications macros you might write or record, but they are not identical. The following example is a WordBasic macro in a Word 95 template.



```vb
Sub MAIN 
FormatFont .Name = "Arial", .Points = 10 
Insert "Hello World" 
End Sub
```

When the template is opened in Word, the macro is converted to the following code.



```vb
Public Sub Main() 
WordBasic.FormatFont Font:="Arial", Points:=10 
WordBasic.Insert "Hello World" 
End Sub
```

Each statement in the converted macro begins with the  **[WordBasic](application-wordbasic-property-word.md)** property.  **WordBasic** is a property in the Word object model that returns an object with all the WordBasic statements and functions; this object makes it possible to run WordBasic macros in Word.

 **Note**  If you save the template over the original template, the WordBasic macros will be permanently lost and previous versions of Word will not be able to use the converted macros.

The following Visual Basic macro is functionally the same as the preceding WordBasic macro, but does not use the  **WordBasic** property.



```vb
Public Sub Main() 
 With Selection.Font 
 .Name = "Arial" 
 .Size = 10 
 End With 
 Selection.TypeText Text:="Hello World" 
End Sub
```


