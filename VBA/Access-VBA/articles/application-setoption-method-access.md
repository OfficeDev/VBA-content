---
title: Application.SetOption Method (Access)
keywords: vbaac10.chm12504
f1_keywords:
- vbaac10.chm12504
ms.prod: access
api_name:
- Access.Application.SetOption
ms.assetid: 6cb1f036-01c2-16bf-f62a-e5235dfb3c65
ms.date: 06/08/2017
---


# Application.SetOption Method (Access)

The  **SetOption** method sets the current value of an option in the **Access Options** dialog box.


## Syntax

 _expression_. **SetOption**( ** _OptionName_**, ** _Setting_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _OptionName_|Required|**String**|The name of the option. For a list of optionname argument strings, see [Set Options from Visual Basic](set-options-from-visual-basic.md).|
| _Setting_|Required|**Variant**|A value corresponding to the option setting. The value of the setting argument depends on the possible settings for a particular option.|

## Remarks

The available option settings depend on the type of option being set. There are three general types of options:


- Yes/No options that can be set by selecting or clearing a check box.
    
- Options that can be set by entering a string or numeric value.
    
- Predefined options that can be chosen from a list box, combo box, or option group.
    
For options that the user sets by selecting or clearing a check box, using the  **SetOption** method, specify **True** or **False** for the setting argument, as in the following example:




```vb
Application.SetOption "Show Status Bar", True
```

To set a type of option using the  **SetOption** method, specify the string or numeric value that would be typed in the dialog box. The following example sets the default form template to OrderTemplate:




```vb
Application.SetOption "Form Template", "OrderTemplate"
```

For options with settings that are choices in list boxes or combo boxes, specify the option's numeric position within the list as the setting argument for the  **SetOption** method. The following example sets the **Default Field Type** option to AutoNumber:




```vb
Application.SetOption "Default Field Type", 5
```

To set an option that's a member of an option group, specify the index number of the option within the group. The following example sets  **Selection Behavior** to Fully Enclosed:




```vb
Application.SetOption "Selection Behavior", 1
```

|**Note**|
|:-----|
|When you quit Microsoft Access, you can reset all options to their original settings by using the  **SetOption** method on all changed options. You may want to create public variables to store the values of the original settings. You might include code to reset options in the Close event procedure for a form, or in a custom exit procedure that the user must run to quit the application.|
  
## See also


#### Concepts


[Application Object](application-object-access.md)

