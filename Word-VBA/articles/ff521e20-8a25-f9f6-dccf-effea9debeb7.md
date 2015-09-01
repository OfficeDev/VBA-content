
# Tasks Object (Word)

 **Last modified:** July 28, 2015

A collection of  ** [Task](8802fcd5-0947-2ea0-308a-376077633e34.md)**objects that represents all the tasks currently running on the system.

## Remarks

Use the  **Tasks** property to return the **Tasks** collection. The following example determines whether Microsoft Excel is running. If it is, this example switches to it and maximizes it; otherwise, the example starts it.


```
If Tasks.Exists("Microsoft Excel") = True Then 
 Tasks("Microsoft Excel").Activate 
 Tasks("Microsoft Excel").WindowState = wdWindowStateMaximize 
Else 
 Shell "C:\Program Files\" &amp; _ 
 "Microsoft Office\Office10\Excel.exe" 
End If
```

Use Visual Basic's  **Shell** function to run an executable program and add the program to the **Tasks** collection.

Use  **Tasks**(Index), where Index is the application name or the index number, to return a single  **Task** object. The following example opens and resizes the application window for the first visible task in the **Tasks** collection.




```
With Tasks(1) 
 If .Visible = True Then 
 .Activate 
 .Width = 400 
 .Height = 200 
 End If 
End With
```

The following example restores the Calculator application window if the application is in the  **Tasks** collection.




```
If Tasks.Exists("Calculator") = True Then 
 Tasks("Calculator").WindowState = wdWindowStateNormal 
End If
```


## See also


#### Concepts


 [Word Object Model Reference](be452561-b436-bb9b-6f94-3faa9a74a6fd.md)
#### Other resources


 [Tasks Object Members](e6ca78c6-132d-6e7b-9f83-ea044a395040.md)
