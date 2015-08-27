
# Options.AllowBackgroundSave Property (Publisher)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


 **True** (default) for Microsoft Publisher to save publications in the background, allowing users to perform other actions at the same time. Read/write **Boolean**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **AllowBackgroundSave**

 _expression_A variable that represents an  **Options** object.


### Return Value

Boolean


## Remarks
<a name="sectionSection1"> </a>

This setting is saved for each individual user and persists from one session to another.


## Example
<a name="sectionSection2"> </a>

This example turns off background save, so publications do not save in the background.


```
Sub DoNotSaveInBackground() 
 Options.AllowBackgroundSave = False 
End Sub
```

