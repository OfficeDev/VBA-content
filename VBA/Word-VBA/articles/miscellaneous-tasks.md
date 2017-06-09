---
title: Miscellaneous Tasks
ms.prod: word
ms.assetid: 5e690651-b220-88d4-f9a1-a7901cb14ec1
ms.date: 06/08/2017
---


# Miscellaneous Tasks

This topic includes Visual Basic examples for the following tasks:


-  [Changing the view](#Changingview)
    
-  [Setting text in a header or footer](#Settingtext)
    
-  [Setting options](#Settingoptions)
    
-  [Changing the document layout](#Changinglayout)
    
-  [Looping through paragraphs in a document](#looping)
    
-  [Customizing menus and toolbars](#Customizing)
    

## Changing the view

The  **[View](view-object-word.md)** object includes properties and methods related to view attributes (such as show all, field shading, and table gridlines) for a window or pane. The following example changes the view to print view.


```vb
Sub ChangeView() 
    ActiveDocument.ActiveWindow.View.Type = wdPrintView 
End Sub
```


## Setting text in a header or footer

The  **[HeaderFooter](headerfooter-object-word.md)** object is returned by the **Headers**,  **Footers**, and  **HeaderFooter** properties. The following example changes the text of the current page header.


```vb
Sub AddHeaderText() 
    With ActiveDocument.ActiveWindow.View 
        .SeekView = wdSeekCurrentPageHeader 
        Selection.HeaderFooter.Range.Text = "Header text" 
        .SeekView = wdSeekMainDocument 
    End With 
End Sub
```

This example creates a  **Range** object, `rngFooter`, that references the primary footer for the first section in the active document. After the  **Range** object is set, the existing footer text is deleted. The FILENAME field is added to the footer along with two tabs and the AUTHOR field.




```vb
Sub AddFooterText() 
    Dim rngFooter As Range 
    Set rngFooter = ActiveDocument.Sections(1) _ 
        .Footers(wdHeaderFooterPrimary).Range 
    With rngFooter 
        .Delete 
        .Fields.Add Range:=rngFooter, Type:=wdFieldFileName, Text:="\p" 
        .InsertAfter Text:=vbTab &; vbTab 
        .Collapse Direction:=wdCollapseStart 
        .Fields.Add Range:=rngFooter, Type:=wdFieldAuthor 
    End With 
End Sub
```


## Setting options

The  **[Options](options-object-word.md)** object includes properties that correspond to optional settings that are available in various menus and dialogs throughout Word. The following example sets three application settings for Word.


```vb
Sub SetOptions() 
    With Options 
        .AllowDragAndDrop = True 
        .ConfirmConversions = False 
        .MeasurementUnit = wdPoints 
    End With 
End Sub
```


## Changing the document layout

The  **[PageSetup](pagesetup-object-word.md)** contains all the page setup attributes of a document (such as left margin, bottom margin, and paper size) as properties. The following example sets the margin values for the active document.


```vb
Sub ChangeDocumentLayout() 
    With ActiveDocument.PageSetup 
        .LeftMargin = InchesToPoints(0.75) 
        .RightMargin = InchesToPoints(0.75) 
        .TopMargin = InchesToPoints(1.5) 
        .BottomMargin = InchesToPoints(1) 
    End With 
End Sub
```


## Looping through paragraphs in a document

This example loops through all of the paragraphs in the active document. If the space-before setting for a paragraph is 6 points, this example changes the spacing to 12 points.


```vb
Sub LoopParagraphs() 
    Dim parCount As Paragraph 
    For Each parCount In ActiveDocument.Paragraphs 
        If parCount.SpaceBefore = 12 Then parCount.SpaceBefore = 6 
    Next parCount 
End Sub
```

For more information, see  [Looping through a collection](looping-through-a-collection.md).


## Customizing menus and toolbars

The  **CommandBar** object represents both menus and toolbars (in versions of Word that do not use the ribbon). Use the **[CommandBars](application-commandbars-property-word.md)** property with a menu or toolbar name to return a single **CommandBar** object. The **Controls** property returns a **CommandBarControls** object that refers to the items on the specified command bar. The following example adds the **Word Count** command to the **Standard** menu.


```vb
Sub AddToolbarItem() 
    Dim btnNew As CommandBarButton 
    CustomizationContext = NormalTemplate 
    Set btnNew = CommandBars("Standard").Controls.Add _ 
        (Type:=msoControlButton, ID:=792, Before:=6) 
    With btnNew 
        .BeginGroup = True 
        .FaceId = 700 
        .TooltipText = "Word Count" 
    End With 
End Sub
```

The following example adds the  **Double Underline** command to the **Formatting** toolbar.




```vb
Sub AddDoubleUnderlineButton() 
    CustomizationContext = NormalTemplate 
    CommandBars("Formatting").Controls.Add _ 
        Type:=msoControlButton, ID:=60, Before:=7 
End Sub
```

Turn on the macro recorder and customize a menu or toolbar to determine the  **ID** value for a particular command (for example, ID 60 is the **Double Underline** command).


