---
title: Tasks Object (Word)
keywords: vbawd10.chm2435
f1_keywords:
- vbawd10.chm2435
ms.prod: word
ms.assetid: ff521e20-8a25-f9f6-dccf-effea9debeb7
ms.date: 06/08/2017
---


# Tasks Object (Word)

A collection of  **[Task](task-object-word.md)** objects that represents all the tasks currently running on the system.


## Remarks

Use the  **Tasks** property to return the **Tasks** collection. The following example determines whether Microsoft Excel is running. If it is, this example switches to it and maximizes it; otherwise, the example starts it.


```vb
If Tasks.Exists("Microsoft Excel") = True Then 
 Tasks("Microsoft Excel").Activate 
 Tasks("Microsoft Excel").WindowState = wdWindowStateMaximize 
Else 
 Shell "C:\Program Files\" &; _ 
 "Microsoft Office\Office10\Excel.exe" 
End If
```

Use Visual Basic's  **Shell** function to run an executable program and add the program to the **Tasks** collection.

Use  **Tasks** (Index), where Index is the application name or the index number, to return a single **Task** object. The following example opens and resizes the application window for the first visible task in the **Tasks** collection.




```vb
With Tasks(1) 
 If .Visible = True Then 
 .Activate 
 .Width = 400 
 .Height = 200 
 End If 
End With
```

The following example restores the Calculator application window if the application is in the  **Tasks** collection.




```vb
If Tasks.Exists("Calculator") = True Then 
 Tasks("Calculator").WindowState = wdWindowStateNormal 
End If
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

