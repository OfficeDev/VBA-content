---
title: Dialog.Show Method (Excel)
keywords: vbaxl10.chm256073
f1_keywords:
- vbaxl10.chm256073
ms.prod: excel
api_name:
- Excel.Dialog.Show
ms.assetid: 7c69ecc2-fdd5-c91b-1c66-e3099bd69cb7
ms.date: 06/08/2017
---


# Dialog.Show Method (Excel)

Displays the built-in dialog box, waits for the user to input data and returns a  **Boolean** value that represents the user's response.


## Syntax

 _expression_ . **Show**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **Dialog** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1-Arg30_|Optional| **Variant**|For built-in dialog boxes only, the initial arguments for the command.|

### Return Value

A Boolean value that, for built-in dialog boxes, returns  **True** if the user clicks OK, or it returns **False** if the user clicks Cancel.


## Remarks

For built in dialog boxes, this method returns  **True** if the user clicks **OK**, or it returns  **False** if the user clicks **Cancel**.

You can use a single dialog box to change many properties at the same time. For example, you can use the Format Cells dialog box to change all the properties of the  **[Font](font-object-excel.md)** object.

For some built-in dialog boxes (the  **Open** dialog box, for example), you can set initial values using _arg1_,  _arg2_, ...,  _arg30_. To find the arguments to set, locate the corresponding dialog box constant in  **Built-In Dialog Box Argument Lists** . For example, search for the **xlDialogOpen** constant to find the arguments for the **Open** dialog box. For more information about built-in dialog boxes, see the **Dialogs** collection.


## Example

This example displays the  **Open** dialog box.


```vb
Application.Dialogs(xlDialogOpen).Show
```


## See also


#### Concepts


[Dialog Object](dialog-object-excel.md)

