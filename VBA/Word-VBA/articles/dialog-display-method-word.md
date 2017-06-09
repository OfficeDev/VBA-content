---
title: Dialog.Display Method (Word)
keywords: vbawd10.chm163053906
f1_keywords:
- vbawd10.chm163053906
ms.prod: word
api_name:
- Word.Dialog.Display
ms.assetid: a9aaa413-ed2f-6fcd-c03e-d76f97783f9a
ms.date: 06/08/2017
---


# Dialog.Display Method (Word)

Displays the specified built-in Word dialog box until either the user closes it or the specified amount of time has passed. Returns a  **Long** that indicates which button was clicked to close the dialog box.


## Syntax

 _expression_ . **Display**( **_TimeOut_** )

 _expression_ Required. A variable that represents a **[Dialog](dialog-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TimeOut_|Optional| **Variant**|The amount of time that Word will wait before closing the dialog box automatically. One unit is approximately 0.001 second. Concurrent system activity may increase the effective time value. If this argument is omitted, the dialog box is closed when the user closes it.|

### Return Value

Long


## Remarks

The  **Display** method returns the following possible values.



|**Return value**|**Description**|
|:-----|:-----|
|-2|The  **Close** button.|
|-1|The  **OK** button.|
|0 (zero)|The  **Cancel** button.|
|> 0 (zero)|A command button: 1 is the first button, 2 is the second button, and so on.|

## Example

This example displays the  **About** dialog box for approximately ten seconds.


```vb
Dim dlgAbout As Dialog 
 
Set dlgAbout = Dialogs(wdDialogHelpAbout) 
dlgAbout.Display TimeOut:=10000
```

This example displays the  **Customize** dialog box.




```
Dialogs(wdDialogToolsCustomize).Display
```


## See also


#### Concepts


[Dialog Object](dialog-object-word.md)

