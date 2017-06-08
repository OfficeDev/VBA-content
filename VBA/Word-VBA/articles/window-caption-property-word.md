---
title: Window.Caption Property (Word)
keywords: vbawd10.chm157417472
f1_keywords:
- vbawd10.chm157417472
ms.prod: word
api_name:
- Word.Window.Caption
ms.assetid: 8d8df29a-7d32-65c8-a714-a356d06b0969
ms.date: 06/08/2017
---


# Window.Caption Property (Word)

Returns or sets the caption text for the window that is displayed in the title bar of the document or application window. Read/write  **String** .


## Syntax

 _expression_ . **Caption**

 _expression_ A variable that represents a **[Window](window-object-word.md)** object.


## Remarks

To change the caption of the application window to the default text, set this property to an empty string ("").


## Example

This example displays the caption of each window in the  **Windows** collection.


```vb
Count = 1 
For Each win In Windows 
 MsgBox Prompt:=win.Caption, Title:="Window" &; Str(Count) &; _ 
 " Caption" 
 Count = Count + 1 
Next win
```

This example resets the caption of the application window.




```vb
Application.Caption = ""
```

This example sets the caption of the active window to the active document name.




```vb
ActiveDocument.ActiveWindow.Caption = ActiveDocument.FullName
```

This example changes the caption of the Word application window to include the user name.




```vb
Application.Caption = UserName &; "'s copy of Word"
```

This example inserts a Table caption and then changes the caption of the first table of figures to "Table."




```vb
Selection.Collapse Direction:=wdCollapseStart 
Selection.Range.InsertCaption "Table" 
If ActiveDocument.TablesOfFigures.Count >= 1 Then 
 ActiveDocument.TablesOfFigures(1).Caption = "Table" 
End If
```


## See also


#### Concepts


[Window Object](window-object-word.md)

