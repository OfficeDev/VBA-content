---
title: Windows Object (Word)
keywords: vbawd10.chm2401
f1_keywords:
- vbawd10.chm2401
ms.prod: word
ms.assetid: 377b493b-e73c-0132-869c-3876c3beaef7
ms.date: 06/08/2017
---


# Windows Object (Word)

A collection of  **[Window](window-object-word.md)** objects that represent all the available windows. The **Windows** collection for the **Application** object contains all the windows in the application, whereas the **Windows** collection for the **Document** object contains only the windows that display the specified document.


## Remarks

Use the  **Windows** property to return the **Windows** collection. The following example tiles all the windows so that they don't overlap one another.


```
Windows.Arrange ArrangeStyle:=wdTiled
```

Use the  **Add** method or the **NewWindow** method to add a new window to the **Windows** collection. Each of the following statements creates a new window for the document in the active window.




```vb
ActiveDocument.ActiveWindow.NewWindow 
NewWindow 
Windows.Add
```

Use  **Windows** (Index), where Index is the window name or the index number, to return a single **Window** object. The following example maximizes the Document1 window.




```
Windows("Document1").WindowState = wdWindowStateMaximize
```

The index number is the number to the left of the window name on the  **Window** menu. The following example displays the caption of the first window in the **Windows** collection.




```vb
MsgBox Windows(1).Caption
```

A colon (:) and a number appear in the window caption when more than one window is open for a document.

When you switch the view to print preview, a new window is created. This window is removed from the  **Windows** collection when you close print preview.


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

