---
title: CommandBarComboBox Members (Office)
ms.prod: office
ms.assetid: 223c51c0-4564-d14a-a8bf-d315a6a50b32
ms.date: 06/08/2017
---


# CommandBarComboBox Members (Office)
Represents a combo box control on a command bar.

Represents a combo box control on a command bar.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[Change](commandbarcombobox-change-event-office.md)|Occurs when the end user changes the selection in a  **CommandBar** combo box.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddItem](commandbarcombobox-additem-method-office.md)|Adds a list item to the specified command bar combo box control. The combo box control must be a custom control and must be a drop-down list box or a combo box.|
|[Clear](commandbarcombobox-clear-method-office.md)|Removes all list items from a command bar combo box control (a drop-down list box or a combo box).|
|[Copy](commandbarcombobox-copy-method-office.md)|Copies a command bar combo box control to an existing command bar.|
|[Delete](commandbarcombobox-delete-method-office.md)|Deletes  **CommandBarCombo** control object from its collection.|
|[Execute](commandbarcombobox-execute-method-office.md)|Runs the procedure or built-in command assigned to the specified  **CommandBarComboBox** control.|
|[Move](commandbarcombobox-move-method-office.md)|Moves the specified control to an existing command bar.|
|[RemoveItem](commandbarcombobox-removeitem-method-office.md)|Removes an item from a  **CommandBarComboBox** control.|
|[Reset](commandbarcombobox-reset-method-office.md)|Resets a built-in command bar to its default configuration, or resets a built-in  **CommandBarComboBox** control to its original function and face.|
|[SetFocus](commandbarcombobox-setfocus-method-office.md)|Moves the keyboard focus to the specified  **CommandBarComboBox** control. If the control is disabled or isn't visible, this method will fail.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](commandbarcombobox-application-property-office.md)|Gets an  **Application** object that represents the container application for the **CommandBarComboBox** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[BeginGroup](commandbarcombobox-begingroup-property-office.md)|Gets True if the specified command bar control appears at the beginning of a group of controls on the command bar. Read/write.|
|[BuiltIn](commandbarcombobox-builtin-property-office.md)|Gets  **True** if the specified command bar control is a built-in control of the container application. Returns **False** if it's a custom control, or if it's a built-in control whose **OnAction** property has been set. Read-only.|
|[Caption](commandbarcombobox-caption-property-office.md)|Gets or sets the caption text for a command bar control. Read/write.|
|[Creator](commandbarcombobox-creator-property-office.md)|Gets a 32-bit integer that indicates the application in which the  **CommandBarComboBox** object was created. Read-only.|
|[DescriptionText](commandbarcombobox-descriptiontext-property-office.md)|Gets or sets the description for a command bar combo box control. Read/write.|
|[DropDownLines](commandbarcombobox-dropdownlines-property-office.md)|Gets or sets the number of lines in a command bar combo box control. The combo box control must be a custom control and it must be a drop-down list box or a combo box. Read/write.|
|[DropDownWidth](commandbarcombobox-dropdownwidth-property-office.md)|Gets or sets the width (in pixels) of the list for the specified command bar combo box control. Read/write.|
|[Enabled](commandbarcombobox-enabled-property-office.md)|Gets or sets a  **Boolean** value that specifies whether the **CommandBarComboBox** is enabled. Read/write.|
|[Height](commandbarcombobox-height-property-office.md)|Gets or sets the height of a  **CommandBarComboBox** control. Read/write.|
|[HelpContextId](commandbarcombobox-helpcontextid-property-office.md)|Gets or sets the Help context Id number for the Help topic attached to the  **CommandBarComboBox** control. Read/write.|
|[HelpFile](commandbarcombobox-helpfile-property-office.md)|Gets or sets the file name for the Help topic attached to the  **CommandBarComboBox** control. Read/write.|
|[Id](commandbarcombobox-id-property-office.md)|Gets the ID for a built-in  **CommandBarComboBox** control. Read-only.|
|[Index](commandbarcombobox-index-property-office.md)|Gets a  **Long** representing the index number for a **CommandBarComboBox** object in the collection. Read-only.|
|[IsPriorityDropped](commandbarcombobox-isprioritydropped-property-office.md)|Gets  **True** if the control is currently dropped from the menu or toolbar based on usage statistics and layout space. (Note that this is not the same as the control's visibility, as set by the **Visible** property). Read-only.|
|[Left](commandbarcombobox-left-property-office.md)|Gets the horizontal position of the  **CommandBarComboBox** control (in pixels) relative to the left edge of the screen. Returns the distance from the left side of the docking area. Read-only.|
|[List](commandbarcombobox-list-property-office.md)|Gets or sets an item in the  **CommandBarComboBox** control. Read/write.|
|[ListCount](commandbarcombobox-listcount-property-office.md)|Gets the number of list items in a  **CommandBarComboBox** control. Read-only.|
|[ListHeaderCount](commandbarcombobox-listheadercount-property-office.md)|Gets or sets the number of list items in a  **CommandBarComboBox** control that appears above the separator line. Read/write.|
|[ListIndex](commandbarcombobox-listindex-property-office.md)|Gets or sets the index number of the selected item in the list portion of the  **CommandBarComboBox** control. If nothing is selected in the list, this property returns zero. Read/write.|
|[OLEUsage](commandbarcombobox-oleusage-property-office.md)|Gets or sets the OLE client and OLE server roles in which a  **CommandBarComboBox** control will be used when two Microsoft Office applications are merged. Read/write.|
|[OnAction](commandbarcombobox-onaction-property-office.md)|Gets or sets the name of a Visual Basic procedure that will run when the user clicks or changes the value of a  **CommandBarComboBox** control. Read/write.|
|[Parameter](commandbarcombobox-parameter-property-office.md)|Gets or sets a string that an application can use to execute a command from a  **CommandBarComboBox** control. Read/write.|
|[Parent](commandbarcombobox-parent-property-office.md)|Gets the  **Parent** object for the **CommandBarComboBox** object. Read-only.|
|[Priority](commandbarcombobox-priority-property-office.md)|Gets or sets the priority of a  **CommandBarComboBox** control. A control's priority determines whether the control can be dropped from a docked command bar if the command bar controls can't fit in a single row. Read/write.|
|[Style](commandbarcombobox-style-property-office.md)|Gets or sets the way a  **CommandBarComboBox** control is displayed. Can be either of the following **MsoComboStyle** constants: **msoComboLabel** or **msoComboNormal**. Read/write.|
|[Tag](commandbarcombobox-tag-property-office.md)|Gets or sets information about the  **CommandBarComboBox** control, such as data that can be used as an argument in procedures, or information that identifies the control. Read/write.|
|[Text](commandbarcombobox-text-property-office.md)|Gets or sets the text in the display or edit portion of the  **CommandBarComboBox** control. Read/write.|
|[TooltipText](commandbarcombobox-tooltiptext-property-office.md)|Gets or sets the text displayed in a  **CommandBarComboBox's** **ScreenTip**. Read/write.|
|[Top](commandbarcombobox-top-property-office.md)|Gets the distance (in pixels) from the top edge of the specified  **CommandBarComboBox** control to the top edge of the screen. Read-only.|
|[Type](commandbarcombobox-type-property-office.md)|Gets the type of  **CommandBarComboBox** control. Read-only.|
|[Visible](commandbarcombobox-visible-property-office.md)|Gets or sets the  **Visible** property for the **CommandBarComboBox** control. **True** if the **CommandBarControl** is visible. Read/write.|
|[Width](commandbarcombobox-width-property-office.md)|Gets or sets the width (in pixels) of the specified  **CommandBarComboBox** control. Read/write.|

