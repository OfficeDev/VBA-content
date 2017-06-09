---
title: Dialog.Show Method (Word)
keywords: vbawd10.chm163053904
f1_keywords:
- vbawd10.chm163053904
ms.prod: word
api_name:
- Word.Dialog.Show
ms.assetid: 6b236493-342d-934b-f360-00b7846789e8
ms.date: 06/08/2017
---


# Dialog.Show Method (Word)

Displays and carries out actions initiated in the specified built-in Word dialog box. Returns a  **Long** that indicates which button was clicked to close the dialog box.


## Syntax

 _expression_ . **Show**( **_TimeOut_** )

 _expression_ Required. A variable that represents a **[Dialog](dialog-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TimeOut_|Optional| **Variant**|The amount of time that Word will wait before closing the dialog box automatically. One unit is approximately 0.001 second. Concurrent system activity may increase the effective time value. If this argument is omitted, the dialog box is closed when the user dismisses it.|

### Return Value

Long


## Remarks

The following table shows the meaning of the values that the  **Show** method returns.



|**Return value**|**Description**|
|:-----|:-----|
|-2|The  **Close** button.|
|-1|The  **OK** button.|
|0 (zero)|The  **Cancel** button.|
|> 0 (zero)|A command button: 1 is the first button, 2 is the second button, and so on.|

## Example

This example displays the  **Find and Replace** dialog box with the word "Blue" preset in the **Find what** text box.


```vb
With Dialogs(wdDialogEditFind) 
 .Find = "Blue" 
 .Show 
End With
```

This example displays and carries out any action initiated in the  **Open** dialog box. The file name is set to *.* so that all file names are displayed.




```vb
With Dialogs(wdDialogFileOpen) 
 .Name = "*.*" 
 .Show 
End With
```

This example displays and carries out any action initiated in the  **Zoom** dialog box. If there are no actions initiated for approximately 9 seconds, the dialog box is closed.




```
Dialogs(wdDialogViewZoom).Show TimeOut:=9000
```


## See also


#### Concepts


[Dialog Object](dialog-object-word.md)

