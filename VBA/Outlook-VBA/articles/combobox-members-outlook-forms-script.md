---
title: ComboBox Members (Outlook Forms Script)
ms.prod: outlook
ms.assetid: a1b81d23-dc10-46cb-b6b3-29fc8968d4ad
ms.date: 06/08/2017
---


# ComboBox Members (Outlook Forms Script)

Combines the features of a  [ListBox](listbox-object-outlook-forms-script.md) and a [TextBox](textbox-object-outlook-forms-script.md).


## Methods





|**Name**|**Description**|
|:-----|:-----|
| [AddItem](combobox-additem-method-outlook-forms-script.md)|For a single-column  [ComboBox](combobox-object-outlook-forms-script.md), the  **AddItem** method adds an item to the list. For a multicolumn **ComboBox**, this method adds a row to the list.|
| [Clear](combobox-clear-method-outlook-forms-script.md)|Removes all entries in the list in a  **ComboBox**.|
| [Copy](combobox-copy-method-outlook-forms-script.md)|Copies the contents of an object to the Clipboard.|
| [Cut](combobox-cut-method-outlook-forms-script.md)|Removes selected information from an object and transfers it to the Clipboard.|
| [DropDown](combobox-dropdown-method-outlook-forms-script.md)|Displays the list portion of a  **ComboBox**.|
| [Paste](combobox-paste-method-outlook-forms-script.md)|Transfers the contents of the Clipboard to an object.|
| [RemoveItem](combobox-removeitem-method-outlook-forms-script.md)|Removes a row from the list in a  **ComboBox**.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [AutoSize](combobox-autosize-property-outlook-forms-script.md)|Returns or sets a  **Boolean** that specifies whether an object automatically resizes to display its entire contents. Read/write.|
| [AutoTab](combobox-autotab-property-outlook-forms-script.md)|Returns or sets a  **Boolean** that specifies whether an automatic tab occurs when a user enters the maximum allowable number of characters into the text box portion of a **ComboBox**. Read/write.|
| [AutoWordSelect](combobox-autowordselect-property-outlook-forms-script.md)|Returns or sets a  **Boolean** that specifies whether the basic unit used to extend a selection is a word or a single character. Read/write.|
| [BackColor](combobox-backcolor-property-outlook-forms-script.md)|Returns or sets a  **Long** that specifies the background color of the object. Read/write.|
| [BackStyle](combobox-backstyle-property-outlook-forms-script.md)|Returns or sets an  **Integer** that specifies the background style for an object. Read/write.|
| [BorderColor](combobox-bordercolor-property-outlook-forms-script.md)|Returns or sets a  **Long** that specifies the border color of an object. Read/write.|
| [BorderStyle](combobox-borderstyle-property-outlook-forms-script.md)|Returns or sets an  **Integer** that specifies the type of border of the control. Read/write.|
| [BoundColumn](combobox-boundcolumn-property-outlook-forms-script.md)|Returns or sets a  **Variant** that identifies the source of data in a multicolumn [ComboBox](combobox-object-outlook-forms-script.md). Read/write.|
| [CanPaste](combobox-canpaste-property-outlook-forms-script.md)|Returns a  **Boolean** that specifies whether the Clipboard contains data that the object supports. Read-only.|
| [Column](combobox-column-property-outlook-forms-script.md)|Returns or sets a  **Variant** that represents a single value, a column of values, or a two-dimensional array to load into a **ComboBox**. Read/write.|
| [ColumnCount](combobox-columncount-property-outlook-forms-script.md)|Returns or sets a  **Long** that represents the number of columns to display in a combo box. Read/write.|
| [ColumnHeads](combobox-columnheads-property-outlook-forms-script.md)|Returns or sets a  **Boolean** that specifies whether a single row of column headings are displayed. Read/write.|
| [ColumnWidths](combobox-columnwidths-property-outlook-forms-script.md)|Returns or sets a  **String** that specifies the width of each column in a multicolumn **ComboBox**. Read/write.|
| [CurTargetX](combobox-curtargetx-property-outlook-forms-script.md)|Returns a  **Long** that represents the preferred horizontal position of the insertion point in a multiline **ComboBox**. Read-only.|
| [CurX](combobox-curx-property-outlook-forms-script.md)|Returns or sets a  **Long** that represents the current horizontal position of the insertion point in a multiline **ComboBox**. Read/write.|
| [DragBehavior](combobox-dragbehavior-property-outlook-forms-script.md)|Returns or sets an  **Integer** that specifies whether the system enables the drag-and-drop feature for the control. Read/write.|
| [DropButtonStyle](combobox-dropbuttonstyle-property-outlook-forms-script.md)|Returns or sets a  **fmDropButtonStyle** value that represents the symbol displayed on the drop button in a **ComboBox**. Read/write.|
| [Enabled](combobox-enabled-property-outlook-forms-script.md)|Returns or sets a  **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.|
| [EnterFieldBehavior](combobox-enterfieldbehavior-property-outlook-forms-script.md)|Returns or sets an  **Integer** that specifies the selection behavior when entering a **ComboBox**. Read/write.|
| [ForeColor](combobox-forecolor-property-outlook-forms-script.md)|Returns or sets a  **Long** that specifies the foreground color of an object. Read/write.|
| [HideSelection](combobox-hideselection-property-outlook-forms-script.md)|Returns or sets a  **Boolean** that specifies whether selected text remains highlighted when a control does not have the focus. Read/write.|
| [IMEMode](combobox-imemode-property-outlook-forms-script.md)|Returns or sets an  **Integer** that specifies the default run-time mode of the Input Method Editor (IME) for a control. Read/write.|
| [LineCount](combobox-linecount-property-outlook-forms-script.md)|Returns a  **Long** that specifies the number of text lines in a **ComboBox**. Read-only.|
| [List](combobox-list-property-outlook-forms-script.md)|Returns or sets a  **Variant** that represents the specified entry in a **ComboBox**. Read/write.|
| [ListCount](combobox-listcount-property-outlook-forms-script.md)|Returns a  **Long** that represents the number of list entries in a control. Read-only.|
| [ListIndex](combobox-listindex-property-outlook-forms-script.md)|Returns or sets a  **Variant** that represents the currently selected item in a **ComboBox**. Read/write.|
| [ListRows](combobox-listrows-property-outlook-forms-script.md)|Returns or sets a  **Long** that specifies the maximum number of rows to display in the list. Read/write.|
| [ListStyle](combobox-liststyle-property-outlook-forms-script.md)|Returns or sets an  **Integer** that specifies the visual appearance of the list in a **ComboBox**. Read/write.|
| [ListWidth](combobox-listwidth-property-outlook-forms-script.md)|Returns or sets a  **Variant** that specifies the width of the list in a **ComboBox**. Read/write.|
| [Locked](combobox-locked-property-outlook-forms-script.md)|Returns or sets a  **Boolean** that specifies whether a control can be edited. Read/write.|
| [MatchEntry](combobox-matchentry-property-outlook-forms-script.md)|Returns or sets an  **Integer** that indicates how a **ComboBox** searches its list as the user types. Read/write.|
| [MatchFound](combobox-matchfound-property-outlook-forms-script.md)|Returns a  **Boolean** value that indicates whether the text that a user has typed into a **ComboBox** matches any of the entries in the list. Read-only.|
| [MatchRequired](combobox-matchrequired-property-outlook-forms-script.md)|Returns or sets a  **Boolean** that specifies whether a value entered in the text portion of a **ComboBox** must match an entry in the existing list portion of the control. Read/write.|
| [MaxLength](combobox-maxlength-property-outlook-forms-script.md)|Returns or sets a  **Long** that specifies the maximum number of characters a user can enter in a **ComboBox**. Read/write.|
| [MouseIcon](combobox-mouseicon-property-outlook-forms-script.md)|Returns a  **String** that represents the full path name of a custom icon that is to be assigned to the control. Read-only.|
| [MousePointer](combobox-mousepointer-property-outlook-forms-script.md)|Returns or sets an  **Integer** that specifies the type of pointer displayed when the user positions the mouse over a particular object. Read/write.|
| [SelectionMargin](combobox-selectionmargin-property-outlook-forms-script.md)|Returns or sets a  **Boolean** that specifies whether the user can select a line of text by clicking in the region to the left of the text. Read/write.|
| [SelLength](combobox-sellength-property-outlook-forms-script.md)|Returns or sets a  **Long** that represents the number of characters selected in the text portion of a **ComboBox**. Read/write.|
| [SelStart](combobox-selstart-property-outlook-forms-script.md)|Returns or sets a  **Long** that represents the starting point of selected text, or the insertion point if no text is selected. Read/write.|
| [SelText](combobox-seltext-property-outlook-forms-script.md)|Returns or sets a  **String** that represents the selected text of a control. Read/write.|
| [ShowDropButtonWhen](combobox-showdropbuttonwhen-property-outlook-forms-script.md)|Returns or sets a  **fmShowDropButtonWhen** value that specifies when to show the drop-down button for a **ComboBox**. Read/write.|
| [SpecialEffect](combobox-specialeffect-property-outlook-forms-script.md)|Returns or sets an  **Integer** that specifies the visual appearance of an object. Read/write.|
| [Style](combobox-style-property-outlook-forms-script.md)|Returns or sets an  **Integer** that specifies how the user can choose or set the control's value. Read/write.|
| [Text](combobox-text-property-outlook-forms-script.md)|Returns or sets a  **String** that specifies text in a **ComboBox**, changing the selected row in the control. Read/write.|
| [TextAlign](combobox-textalign-property-outlook-forms-script.md)|Returns or sets an  **Integer** that specifies how text is aligned in a control. Read/write.|
| [TextColumn](combobox-textcolumn-property-outlook-forms-script.md)|Returns or sets a  **Variant** that identifies the column in a **ComboBox** to display to the user. Read/write.|
| [TextLength](combobox-textlength-property-outlook-forms-script.md)|Returns a  **Long** that represents the length, in number of characters, of text in the edit region of a **ComboBox**. Read-only.|
| [TopIndex](combobox-topindex-property-outlook-forms-script.md)|Returns or sets a  **Long** that represents the index of the item displayed in the topmost position in the list portion of the **ComboBox**. Read/write.|
| [Value](combobox-value-property-outlook-forms-script.md)|Returns or sets a  **Variant** that specifies the value in the [BoundColumn](combobox-boundcolumn-property-outlook-forms-script.md) of the currently selected rows. Read/write.|



## Events



|**Name**|**Description**|
|:-----|:-----|
| [Click](combobox-click-event-outlook-forms-script.md)|Occurs when the user definitively selects a value for the control that has more than one possible value.|



