
# Options.AllowBackgroundSave Property (Publisher)

 **Last modified:** July 28, 2015

 _**Applies to:** Publisher 2013 | VBA_

 **True** (default) for Microsoft Publisher to save publications in the background, allowing users to perform other actions at the same time. Read/write **Boolean**.


## Syntax

 _expression_. **AllowBackgroundSave**

 _expression_A variable that represents an  **Options** object.


### Return Value

Boolean


## Remarks

This setting is saved for each individual user and persists from one session to another.


## Example

This example turns off background save, so publications do not save in the background.


```
Sub DoNotSaveInBackground() 
 Options.AllowBackgroundSave = False 
End Sub
```

