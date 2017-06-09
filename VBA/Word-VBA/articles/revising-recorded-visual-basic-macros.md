---
title: Revising Recorded Visual Basic Macros
ms.prod: word
ms.assetid: e17875d2-f11a-825c-1f92-a0ba6cb3309f
ms.date: 06/08/2017
---


# Revising Recorded Visual Basic Macros

The macro recorder is a great tool for discovering the Visual Basic methods and properties that you want to use. If you do not know what properties or methods to use, turn on the macro recorder and manually perform the action. The macro recorder translates your actions into Visual Basic code. There are, however, some limitations to recording macros. You cannot record the following:


- Conditional branches
    
- Variable assignments
    
- Looping structures
    
- Custom user forms
    
- Error handling
    
- Text selections made with the mouse (you must use keyboard combinations)
    

To enhance your macros, you may want to revise the code recorded into your module.


## Removing the Selection property

Macros created using the macro recorder depend on the selection. At the beginning of most recorded macro instructions, you see  `Selection`. Recorded macros use the  **[Selection](global-selection-property-word.md)** property to return the **[Selection](selection-object-word.md)** object. For example, the following example moves the selection to the Temp bookmark and inserts text after the bookmark.


```vb
Sub Macro1() 
    Selection.Goto What:=wdGotoBookmark, Name:="Temp" 
    Selection.MoveRight Unit:=wdCharacter, Count:=1 
    Selection.TypeText Text:="New text" 
End Sub
```

This macro accomplishes the task, but there are a couple of drawbacks. First, if the document does not have a bookmark named Temp, the macro posts an error. Second, the macro moves the selection, which may not be appropriate. Both of these issues can be resolved by revising the macro so that it does not use the  **Selection** object. This is the revised macro.




```vb
Sub MyMacro() 
    If ActiveDocument.Bookmarks.Exists("Temp") = True Then 
        endloc = ActiveDocument.Bookmarks("Temp").End 
        ActiveDocument.Range(Start:=endloc, _ 
        End:=endloc).InsertAfter "New text" 
    End If 
End Sub
```

The  **[Exists](bookmarks-exists-method-word.md)** method is used to check for the existence of the bookmark named Temp. If the bookmark is found, the bookmark's ending character position is returned by using the **[End](bookmark-end-property-word.md)** property. Finally, the **[Range](document-range-method-word.md)** method of the **Document** object is used to return a **[Range](range-object-word.md)** object that refers to the bookmark's ending position, so that text can be inserted using the **[InsertAfter](range-insertafter-method-word.md)** method of the **Range** object. For more information about defining **Range** objects, see [Working with Range objects](working-with-range-objects.md).


## Using With…End With

Macro instructions that refer to the same object can be simplified using a  **With…End With** structure. For example, the following macro was recorded when a title was added at the top of a document.


```vb
Sub Macro1() 
    Selection.HomeKey Unit:=wdStory 
    Selection.TypeText Text:="Title" 
    Selection.ParagraphAlignment.Alignment = wdAlignParagraphCenter 
End Sub
```

The  **Selection** property is used with each instruction to return a **Selection** object. The macro can be simplified so that the **Selection** property is used only once.




```vb
Sub MyMacro() 
    With Selection 
        .HomeKey Unit:=wdStory 
        .TypeText Text:="Title" 
        .ParagraphAlignment.Alignment = wdAlignParagraphCenter 
    End With 
End Sub
```

The same task can also be performed without using the  **Selection** object. The following macro uses a **Range** object at the beginning of the active document to accomplish the same task.




```vb
Sub MyMacro() 
    With ActiveDocument.Range(Start:=0, End:=0) 
        .InsertAfter "Title" 
        .ParagraphFormat.Alignment = wdAlignParagraphCenter 
    End With 
End Sub
```


## Removing unnecessary properties

If you record a macro that involves selecting an option in a dialog box, the macro recorder records the settings of all the options in the dialog box, even if you only change one or two options. If you do not need to change all of the options, you can remove the unnecessary properties from the recorded macro. The following recorded macro includes a number of options from the  **Paragraph** dialog box ( **Format** menu).


```vb
Sub Macro1() 
    With Selection.ParagraphFormat 
        .LeftIndent = InchesToPoints(0) 
        .RightIndent = InchesToPoints(0) 
        .SpaceBefore = 6 
        .SpaceAfter = 6 
        .LineSpacingRule = 0 
        .Alignment = wdAlignParagraphLeft 
        .WidowControl = True 
        .KeepWithNext = False 
        .KeepTogether = False 
        .PageBreakBefore = False 
        .NoLineNumber = False 
        .Hyphenation = True 
        .FirstLineIndent = InchesToPoints(0) 
        .OutlineLevel = 10 
    End With 
End Sub
```

However, if you only want to change the spacing before and after the paragraph, you can change the macro to the following.




```vb
Sub MyMacro() 
    With Selection.ParagraphFormat 
        .SpaceBefore = 6 
        .SpaceAfter = 6 
    End With 
End Sub
```

The simplified macro executes faster because it sets fewer properties. Only the spacing before and after are changed; all of the other settings for the selected paragraphs are unchanged.


## Removing unnecessary arguments

When the macro recorder records a method, the values of all of the arguments are included. The following macro was recorded when the document named Test.doc was opened. The resulting macro includes all of the arguments for the  **[Open](documents-open-method-word.md)** method.


```vb
Sub Macro1() 
    Documents.Open FileName:="C:\My Documents\Test.doc", _ 
        ConfirmConversions:= False, ReadOnly:=False, _ 
        AddToRecentFiles:=False, PasswordDocument:="", _ 
        PasswordTemplate:="", Revert:=False, _ 
        WritePasswordDocument:="", _ 
        WritePasswordTemplate:="", Format:=wdOpenFormatAuto 
End Sub
```

The arguments that are not needed can be removed from the recorded macro. For example, you could remove all of arguments set to an empty string (for example,  `WritePasswordDocument:=""`), as shown.




```vb
Sub MyMacro() 
    Documents.Open FileName:="C:\My Documents\Test.doc", _ 
        ConfirmConversions:= False, _ 
        ReadOnly:=False, AddToRecentFiles:=False, _ 
        Revert:=False, Format:=wdOpenFormatAuto 
End Sub
```


