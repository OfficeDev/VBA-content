---
title: Match Entries in a List to Entries Typed by Users
ms.prod: outlook
ms.assetid: 629c3c16-e132-b062-c733-7ecb4a856694
ms.date: 06/08/2017
---


# Match Entries in a List to Entries Typed by Users

1. In the Form Designer, drag the  [ListBox](listbox-object-outlook-forms-script.md) or [ComboBox](combobox-object-outlook-forms-script.md) control from the [Control Toolbox](show-or-hide-the-control-toolbox.md) to the form.
    
2. Right-click the list box or combo box, and then click  **Advanced Properties**. 
    
3. To set the  **MatchEntry** property, click the property, specify a value in the **Properties** box, and then click **Apply**.
    
    For more information about the property for the specific control, see the following:
    
      -  [MatchEntry](listbox-matchentry-property-outlook-forms-script.md) property for the **ListBox** control.
    
  -  [MatchEntry](combobox-matchentry-property-outlook-forms-script.md) property for the **ComboBox** control.
    


|**Set this MatchEntry value **|**To**|
|:-----|:-----|
| **No matching**|Provide no matching.|
| **First letter**|Compare the most recently typed letter to the first letter of each entry in the list (the first match in the list is selected).|
| **Complete**|Compare the user's entry and an exact match in an entry from the list.|

 **Note**  


