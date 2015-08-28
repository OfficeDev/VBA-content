
# TextBox Members (Access)
This object corresponds to a text box. Text boxes on a form or report display data from a record source.

 **Last modified:** July 28, 2015

 **In this article**
 [Events](#sectionSection0)
 [Methods](#sectionSection1)
 [Properties](#sectionSection2)


## Events
<a name="sectionSection0"> </a>



|**Name**|**Description**|
|:-----|:-----|
| [AfterUpdate](609ef5f3-3894-85eb-4879-5db3fc7ff188.md)|The  **AfterUpdate** event occurs after changed data in a control or record is updated.|
| [BeforeUpdate](0d57cbce-bdbf-e19e-7f6a-11a00cb6c5f4.md)|The  **BeforeUpdate** event occurs before changed data in a control or record is updated.|
| [Change](adde0a6d-d37a-a457-0dea-f2358adbb665.md)|The  **Change** event occurs when the contents of the specified control changes.|
| [Click](d102a526-2051-3a36-0f7a-fc234f126c47.md)|The  **Click** event occurs when the user presses and then releases a mouse button over an object.|
| [DblClick](ae8787e1-3425-bfbf-acf4-bbb97d42d2da.md)|The  **DblClick** event occurs when the user presses and releases the left mouse button twice over an object within the double-click time limit of the system.|
| [Dirty](d6073892-7618-8e23-1fb1-795d3c76c2b6.md)|The Dirty event occurs when the contents of the specified control changes.|
| [Enter](970dc73b-8b8e-5811-bd4b-c23a96306bd2.md)|The  **Enter** event occurs before a control actually receives the focus from a control on the same form or report.|
| [Exit](05b5afca-4cb9-f12b-e05b-8702e35380d0.md)|The  **Exit** event occurs just before a control loses the focus to another control on the same form or report.|
| [GotFocus](bc5d12a2-476b-a91d-2ad4-cdd6f46dd44c.md)|The  **GotFocus** event occurs when the specified object receives the focus.|
| [KeyDown](00324700-f101-48a0-242f-bdabf4f2d70d.md)|The  **KeyDown** event occurs when the user presses a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.|
| [KeyPress](87db62a8-30f6-03d8-63ae-f1a1a50caea3.md)|The  **KeyPress** event occurs when the user presses and releases a key or key combination that corresponds to an ANSI code while a form or control has the focus. This event also occurs if you send an ANSI keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.|
| [KeyUp](2219075d-92e5-a472-c16a-8a99dfd991c2.md)|The  **KeyUp** event occurs when the user releases a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.|
| [LostFocus](4c3a2696-5a78-5be9-7af7-205e7eb84dcd.md)|The  **LostFocus** event occurs when the specified object loses the focus.|
| [MouseDown](ae184752-4c7f-3d79-5b3a-08407225f9d9.md)|The  **MouseDown** event occurs when the user presses a mouse button.|
| [MouseMove](90d5d17b-8802-ec93-11ad-6be846bb1efe.md)|The  **MouseMove** event occurs when the user moves the mouse.|
| [MouseUp](0dfdc0b3-4a31-fd96-481c-d13db8197edd.md)|The  **MouseUp** event occurs when the user releases a mouse button.|
| [Undo](ee009e53-41be-0c9a-a92d-15572f6213b6.md)|Occurs when the user undoes a change.|

## Methods
<a name="sectionSection1"> </a>



|**Name**|**Description**|
|:-----|:-----|
| [Move](50b25305-0b91-378d-514f-d35b8d7aed6e.md)|Moves the specified object to the coordinates specified by the argument values.|
| [Requery](b1f8991e-7ccc-4f0b-c50f-1d51a0abda7e.md)|The  **Requery** method updates the data underlying a specified control that's on the active form by requerying the source of data for the control.|
| [SetFocus](dc5edcd0-09af-2fdb-0b94-49af0bfa705b.md)|The  **SetFocus** method moves the focus to the specified form, the specified control on the active form, or the specified field on the active datasheet.|
| [SizeToFit](17289703-1943-2499-48c5-f34f200fd304.md)|You can use the  **SizeToFit** method to size a control so it fits the text or image that it contains.|
| [Undo](b019355a-7b78-4f03-878f-d2830c20117d.md)|You can use the  **Undo** method to reset a control or form when its value has been changed.|

## Properties
<a name="sectionSection2"> </a>



|**Name**|**Description**|
|:-----|:-----|
| [AddColon](0a908d65-921b-7722-b564-cfe7a7fa8aed.md)|Specifies whether a colon follows the text in labels for new controls. Read/write  **Boolean**.|
| [AfterUpdate](690bc0cd-9717-7712-c022-75ba457ca0e3.md)|Returns or sets which macro, event procedure, or user-defined function runs when the  **AfterUpdate**event occurs. Read/write  **String**.|
| [AllowAutoCorrect](9cafa161-c073-855f-edee-c7c9cb32be99.md)|You can use the  **AllowAutoCorrect** property to specify whetherthe specified control will automatically correct entries made by the user. Read/write **Boolean**.|
| [Application](84a7ea86-f31c-775d-2383-5ac8751dd0f1.md)|You can use the  **Application** property to access the active Microsoft Access ** [Application](aefb0713-97e6-e2c7-e530-8fd2e1316a55.md)**object and its related properties. Read-only  **Application** object.|
| [AsianLineBreak](2ee42bb4-e6ae-c6b4-ef6a-71de5d35edad.md)|Returns or sets a  **Boolean** indicating whether line breaks in text boxes follow rules governing East Asian languages. **True** to control line breaks based on East Asian language rules. Read/write.|
| [AutoLabel](a5e6e68c-eadc-a242-ef83-8b388f6ca41f.md)|Specifies whether labels are automatically created and attached to new controls. Read/write  **Boolean**.|
| [AutoTab](27b17921-cd58-e243-e091-2686c64a7c02.md)|You can use the  **AutoTab** property to specify whether an automatic tab occurs when the last character permitted by a text box control's input mask is entered. An automatic tab moves the focus to the next control in the form's tab order. Read/write **Boolean**.|
| [BackColor](7880c596-7a47-39b6-74ad-8036355a8e0f.md)|Gets or sets the interior color of the specified object. Read/write  **Long**.|
| [BackShade](36db2540-6d5b-ed43-a303-70b6282398cf.md)|Gets or sets the shade applied to the theme color in the  **BackColor** property of the specified object. Read/write **Single**.|
| [BackStyle](95a277c8-df48-79a5-c232-2cfe32eae8f2.md)|You can use the  **BackStyle** property to specify whether a control will be transparent. Read/write **Byte**.|
| [BackThemeColorIndex](a66a4839-3ab9-4867-b725-e613527bc94b.md)|Gets or sets a value that represents a color in the applied color theme associated with the  **BackColor** property of the specified object. Read/write **Long**.|
| [BackTint](3740b360-334c-db71-9fb6-1f7aab304811.md)|Gets or sets the tint that is applied to the theme color in the  **BackColor** property of the specified object. Read/write **Single**.|
| [BeforeUpdate](de841054-a98a-7108-0d7d-020175edb1ce.md)|Returns or sets which macro, event procedure, or user-defined function runs when the  **BeforeUpdate**event occurs. Read/write  **String**.|
| [BorderColor](7522b663-4ce6-34a6-51db-7de503e01f04.md)|You can use the  **BorderColor** property to specify the color of a control's border. Read/write **Long**.|
| [BorderShade](554920e1-e5ae-1c48-f5d5-ab964070bec0.md)|Gets or sets the shade that is applied to the theme color in the  **BorderColor** property of the specified object. Read/write **Single**.|
| [BorderStyle](783c9424-669f-fcc7-b23d-6f5de03bad79.md)|Specifies how a control's border appears.Read/write  **Byte**.|
| [BorderThemeColorIndex](44f012fa-9021-0910-85c0-48a3b6c82141.md)|Gets or sets a value that represents a color in the applied color theme associated with the  **BorderColor** property of the specified object. Read/write **Long**.|
| [BorderTint](3e48aa7c-ed95-aa27-f092-70d5fb2f9fb1.md)|Gets or sets the tint that is applied to the theme color in the  **BorderColor** property of the specified object. Read/write **Single**.|
| [BorderWidth](e842887f-9ec1-4405-0558-6b3b3d3d221c.md)|You can use the  **BorderWidth** property to specify the width of a control's border. Read/write **Byte**.|
| [BottomMargin](a6ef1155-24c8-1254-614b-c912fda8dae2.md)|Along with the  **LeftMargin**,  **RightMargin**, and  **TopMargin** properties, specifies the location of information displayed within a text box control. Read/write **Integer**.|
| [BottomPadding](75d2b8bb-c5c5-1d00-b175-8db80a7525c5.md)|Gets or sets the amount of space (in inches) between the text box and its bottom gridline. Read/write  **Integer**.|
| [CanGrow](5e96e693-9e1a-1f1f-5d5d-672e6232c330.md)|Gets or sets whether the specified control automatically adjusts vertically to print or preview all the data the control contains. Read/write  **Boolean**.|
| [CanShrink](d4ac842c-18ea-a3be-a90a-5dd9d10d7b8f.md)|Gets or sets whether the specified control automatically adjusts vertically to print or preview all the data the section or control contains. Read/write  **Boolean**.|
| [ColumnHidden](4014ea78-92f8-f1a8-6d73-ae7b2c5088cb.md)|You can use the  **ColumnHidden** property to show or hide a specified column in Datasheet view. Read/write **Boolean**.|
| [ColumnOrder](b5b271bc-5b3c-9b2c-ec87-524be29597d0.md)|You can use the  **ColumnOrder** property to specify the order of the columns in Datasheet view. Read/write **Integer**.|
| [ColumnWidth](19060aac-ccb0-3998-39c7-42f1454c339e.md)|You can use the  **ColumnWidth** property to specify the width of a column in Datasheet view. Read/write **Integer**.|
| [Controls](00d5dede-0583-9f0e-191a-28f91a0327b3.md)|Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.|
| [ControlSource](be912167-402a-1bc4-6feb-c3551eb058a8.md)|You can use the  **ControlSource** property to specify what data appears in a control. You can display and edit data bound to a field in a table, query, or SQL statement. You can also display the result of an expression. Read/write **String**.|
| [ControlTipText](a63f3624-8f31-97f6-c2cb-8c34c82c825b.md)|You can use the  **ControlTipText** property to specify the text that appears in a ScreenTip when you hold the mouse pointer over a control. Read/write **String**.|
| [ControlType](4cc842d9-2985-b65e-e259-697cedaa56fc.md)|You can use the  **ControlType** property in Visual Basic to determine the type of a control on a form or report. Read/write **Byte**.|
| [DecimalPlaces](cd032c51-34d1-18d3-c378-7473938ec1d7.md)|You can use the  **DecimalPlaces** property to specify the number of decimal places Microsoft Access uses to display numbers. Read/write **Byte**.|
| [DefaultValue](fab86da0-e865-478c-80c6-7681c5733059.md)|Specifies a value that is automatically entered in a field when a new record is created. For example, in an Addresses table you can set the default value for the City field to New York. When users add a record to the table, they can either accept this value or enter the name of a different city. Read/write  **String**.|
| [DisplayAsHyperlink](4741039e-9985-ac0a-9b74-309fcac860bf.md)|Gets or sets an  ** [AcDisplayAsHyperlink](fb9d9af3-9aff-3031-2f94-6715211d6ee4.md)** constant that specifies whether to display the contents of the specified text box as a hyperlink. Read/write.|
| [DisplayWhen](6e5fa1c0-a264-cbc1-6fdf-9aef6c7f6bab.md)|You can use the  **DisplayWhen** property to specify which of a form's controls you want displayed on screen and in print. Read/write **Byte**.|
| [Enabled](a13297e5-091c-7e83-78cd-fa67f5b81153.md)|You can use the  **Enabled** property to set or return the status of the conditional format in the ** [FormatCondition](a31deaae-b32d-c45b-b3b2-113a9e62cc7a.md)**object. Read/write  **Boolean**.|
| [EnterKeyBehavior](b7830316-a1aa-ddc1-094f-5976c5298bc1.md)|You can use the  **EnterKeyBehavior** property to specify what happens when you press ENTER in a text box control in Form view or Datasheet view. Read/write **Boolean**.|
| [EventProcPrefix](a8cd7cdc-605b-473c-95b1-9d1736e0ec96.md)|Gets or sets the prefix portion of an event procedure name. Read/write  **String**.|
| [FilterLookup](5c568366-94a5-8d7a-1fb4-80b4b3ab6c7f.md)|You can use the  **FilterLookup** property to specify whether values appear in a bound text box control when using the Filter By Form or Server Filter By Form window. Read/write **Byte**.|
| [FontBold](147d151a-b51c-5be2-56ef-8a94c212cb0b.md)|You can use the  **FontBold** property to specify whether a font appears in a bold style in the following situations:|
| [FontItalic](f982c1ce-ad47-a05e-6b12-1eb51dbc0eb7.md)|You can use the  **FontItalic** property to specify whether text is italic in the following situations:|
| [FontName](4eb7cbbe-1e7d-9e29-cbff-867a83605741.md)|You can use the  **FontName** property to specify the font for text in the following situations:|
| [FontSize](73bf8d74-c616-8824-c2e0-8eed072df582.md)|You can use the  **FontSize** property to specify the point size for text in the following situations:|
| [FontUnderline](67bf0551-21c0-73cd-9418-dc7b3582f53c.md)|You can use the  **FontUnderline** property to specify whether text is underlined in the following situations:|
| [FontWeight](4dbf8092-c09c-c6ec-9476-20af2e9cf051.md)|You can use the  **DatasheetFontWeight** property to specify the line width of the font used to display and print characters for field names and data in Datasheet view. Read/write **Integer**.|
| [ForeColor](125bc04a-b747-6397-33ff-31de47004633.md)|You can use the  **ForeColor** property to specify the color for text in a control. Read/write **Long**.|
| [ForeShade](b8437ede-edd1-7d86-1c2f-78d4ed1c3d0e.md)|Gets or sets the shade that is applied to the theme color in the  **ForeColor** property of the specified object. Read/write **Single**.|
| [ForeThemeColorIndex](9b49e363-fe5b-0536-c3ed-b4836acb383b.md)|Gets or sets a value that represents a color in the applied color theme associated with the  **ForeColor** property of the specified object. Read/write **Long**.|
| [ForeTint](8229f864-5ed3-309e-ba29-6a45bf9d59a8.md)|Gets or sets the tint that is applied to the theme color in the  **ForeColor** property of the specified object. Read/write **Single**.|
| [Format](c89491e2-09f8-d928-1aed-9d839545a694.md)|You can use the  **Format** property to customize the way numbers, dates, times, and text are displayed and printed. Read/write **String**.|
| [FormatConditions](6c643d8b-9b90-2b50-2ba0-c46bb821d38d.md)|You can use the  **FormatConditions** property to return a read-only reference to the ** [FormatConditions](0a1cd89b-6690-8272-ebd9-d841e9fb1d4c.md)**collection and its related properties.|
| [FuriganaControl](7d70cffa-06bb-fa9d-686a-0031558aa5a3.md)||
| [GridlineColor](849e0843-ab35-90d6-02a6-44faa316c83f.md)|Gets or sets the color of the gridline for the specified text box. Read/write  **Long**.|
| [GridlineShade](33daf4ec-1587-63c8-4b23-2abdf5087bbe.md)|Gets or sets the shade applied to the theme color in the  **GridlineColor** property of the specified object. Read/write **Single**.|
| [GridlineStyleBottom](c58d8030-fc96-a53b-4cb4-5bb21237e20e.md)|Gets or sets the bottom gridline style of the specified text box. Read/write  **Byte**.|
| [GridlineStyleLeft](f1c71748-a37c-d0d0-5d8e-9899cf1efba5.md)|Gets or sets the width of the bottom gridline for the specified text box. Read/write  **Byte**.|
| [GridlineStyleRight](c841157d-6e8d-8cd4-e23a-77d00d0af8e6.md)|Gets or sets the right gridline style of the specified text box. Read/write  **Byte**.|
| [GridlineStyleTop](57a47306-5b85-06e0-e59f-f86e617d9c75.md)|Gets or sets the top gridline style of the specified text box. Read/write  **Byte**.|
| [GridlineThemeColorIndex](2c67d4b5-47d6-5430-cac0-bc05c3151305.md)|Gets or sets the theme color index that represents a color in the applied color theme associated with the  **GridlineColor** property of the specified object. Read/write **Long**.|
| [GridlineTint](5dbbd8a7-0942-c39d-b702-a3c0e569e3c1.md)|Gets or sets the tint applied to the theme color in the  **GridlineColor** property of the specified object. Read/write **Single**. |
| [GridlineWidthBottom](4569d053-008b-a4ce-374f-6078f5ea9bee.md)|Gets or sets the width of the bottom gridline for the specified text box. Read/write  **Byte**.|
| [GridlineWidthLeft](0794df4f-88e2-5c75-13ba-88bbb8d7eb40.md)|Gets or sets the width of the left gridline for the specified text box. Read/write  **Byte**.|
| [GridlineWidthRight](6abe0945-a6b9-72b2-e63c-1109fc7455a8.md)|Gets or sets the width of the right gridline for the specified text box. Read/write  **Byte**.|
| [GridlineWidthTop](bb49f001-83a9-f1b8-c095-33b8b3f820b3.md)|Gets or sets the width of the top gridline for the specified text box. Read/write  **Byte**.|
| [Height](3f3d88d9-3561-a25b-5f22-a21b9cad6673.md)|Gets or sets the height of the specified object in twips. Read/write  **Integer**.|
| [HelpContextId](6829c95e-d7fc-c3c6-a8ab-0051c8e9af24.md)|The  **HelpContextID** property specifies the context ID of a topic in the custom Help file specified by the **HelpFile** property setting. Read/write **Long**.|
| [HideDuplicates](a625d232-07d8-23d9-a28a-d01c70aa3a95.md)|You can use the  **HideDuplicates** property to hide a control on a report when its value is the same as in the preceding record. Read/write **Boolean**.|
| [HorizontalAnchor](85dc54b2-7a20-4667-ade9-47202f77d524.md)|Gets or sets an  ** [AcHorizontalAnchor](2b9f0574-252d-7957-d25d-cb382d2cee73.md)** constant that indicates how the text box is anchored horizontally within its layout. Read/write.|
| [Hyperlink](a5d80cd4-d03d-41ea-9394-214537dd6c8c.md)|You can use the  **Hyperlink** property to return a reference to a **Hyperlink**object. You can use the  **Hyperlink** property to access the properties and methods of a control's hyperlink. Read-only.|
| [IMEHold](0cb93c85-07ff-a10f-5cd0-dc4045ce1079.md)| [Language-specific information](47c3b4cf-01ef-0b87-5cf1-50967397893f.md)You can use the  **IMEHold/Hold KanjiConversionMode** property to show whether the Kanji Conversion Mode is maintained when the control loses the focus. Read/write **Boolean**.|
| [IMEMode](fa4adf03-7c20-eade-4a28-e3c3ac64ebc3.md)||
| [IMESentenceMode](399a28d4-83a9-33d2-5f00-4f388efe048b.md)||
| [InputMask](a705c2a4-ff2f-74d1-4a7c-1eade3b00ae8.md)|You can use the  **InputMask** property to make data entry easier and to control the values users can enter in a text boxcontrol. Read/write **String**.|
| [InSelection](6ebb497c-00d0-a854-be22-6b034deae98a.md)|You can use the  **InSelection** property to determine or specify whether a control on a form in Design view is selected. Read/write **Boolean**.|
| [IsHyperlink](68d2ca6a-7ea2-a44d-db32-1fa040475267.md)|You can use the  **IsHyperlink** property to specify or determine if the data contained in a text box is a hyperlink. Read/write **Boolean**.|
| [IsVisible](34487db4-6377-04f2-6848-a27dc5f4bab6.md)|You can use the  **IsVisible** property in to determine whether a control on a report is visible. Read/write **Boolean**.|
| [KeyboardLanguage](a3b55e3e-16a9-87c7-6c03-bc8392e72c17.md)||
| [LabelAlign](4714927a-9ce9-b1b0-dbeb-302aaa48a6c8.md)|The property specifies the text alignment within attached labels on new controls. Read/write  **Byte**.|
| [LabelX](4d3ce486-bd79-cd6d-5886-a0a64cc7abb5.md)|The  **LabelX** property (along with the **LabelY** property) specifies the placement of the label for a new control. Read/write **Integer**.|
| [LabelY](398b268c-271d-566a-36a6-1d703bdb0345.md)|The  **LabelY** property (along with the **LabelX** property) specifies the placement of the label for a new control. Read/write **Integer**.|
| [Layout](a1c841e6-221b-3ba6-4212-d76066afda48.md)|Returns the type of layout for the specified text box. Read-only  ** [AcLayoutType](ee963ed0-9293-8ad8-5694-4b93a5e4d89a.md)**.|
| [LayoutID](b77ccc32-fbaf-e574-b0ae-293d6f999879.md)|Returns the unique identifier for the layout that contains the specified text box. Read-only  **Long**.|
| [Left](a184b336-215d-ffe0-d7ce-92f1fdc3b656.md)|You can use the  **Left** property to specify an object's location on a form or report. Read/write **Integer**.|
| [LeftMargin](9c5b798b-4afe-85be-aa06-eeff98888850.md)|Along with the  **TopMargin**,  **RightMargin**, and  **BottomMargin** properties. specifies the location of information displayed within a text box control. Read/write **Integer**. .|
| [LeftPadding](0ceae1bc-f075-2e5f-48bf-7f749bae0630.md)|Gets or sets the amount of space (in inches) between the text box and its left gridline. Read/write  **Integer**.|
| [LineSpacing](3ac1c335-4b26-1a14-e4dc-bd5d56f44a2b.md)|You can use the  **LineSpacing** property to specify or determine the location of information displayed within a label or text box control. Read/write **Integer**.|
| [Locked](025b88db-7409-4cb6-bcc0-c72a6a3850d3.md)|The  **Locked** property specifies whether you can edit data in a control in Form view. Read/write **Boolean**.|
| [Name](e97043b5-216f-2c5c-a531-45b29477cb77.md)|You can use the  **Name** property to specify or determine the string expression that identifies the name of an object. Read/write **String**.|
| [NumeralShapes](f0fda4bb-2522-622c-24ab-d3324a4b8dca.md)||
| [OldBorderStyle](6064f8b9-31ec-da00-0346-cd259b917daa.md)|You can use this property to set or returns the unedited value of the  **BorderStyle** property for a form or control. This property is useful if you need to revert to an unedited or preferred border style. Read/write **Byte**.|
| [OldValue](d62150d2-6dc6-85c0-0452-e9e5fee199b4.md)|You can use the  **OldValue** property to determine the unedited value of a bound control. Read-only **Variant**.|
| [OnChange](ea25341f-fd30-62b1-476d-29febf4db4b4.md)|Sets or returns the value of the  **On Change** box in the **Properties** window of one of the objects in the Applies To list. Read/write **String**.|
| [OnClick](54f32b3d-16df-376d-b5c0-9bbb2ff0931a.md)|Sets or returns the value of the  **On Click** box in the **Properties** window. Read/write **String**.|
| [OnDblClick](571a01ff-b92b-bb9b-1363-43086ef71c02.md)|Sets or returns the value of the  **On Dbl Click** box in the **Properties** window. Read/write **String**.|
| [OnDirty](312418b3-29cf-0d53-d92f-aaca6ec192b3.md)|Sets or returns the value of the  **On Dirty** box in the **Properties** window of a form or report. Read/write **String**.|
| [OnEnter](d8f7aa7f-5222-ec0e-7be9-4022f5e697af.md)|Sets or returns the value of the  **On Enter** box in the **Properties** window of specified object. Read/write **String**. .|
| [OnExit](2489acdf-4cf5-8b49-e9fe-fc78c07a87f3.md)|Sets or returns the value of the  **On Exit** box in the **Properties** window of specified object. Read/write **String**. .|
| [OnGotFocus](3a180b9a-d415-b124-f884-9ce64dba8358.md)|Sets or returns the value of the  **On Got Focus** box in the **Properties** window of the specified object. Read/write **String**.|
| [OnKeyDown](472e4b96-a6b1-6473-ed56-64af3522281f.md)|Sets or returns the value of the  **On Key Down** box in the **Properties** window. Read/write **String**.|
| [OnKeyPress](458d2e2d-3003-79e4-a911-058928c25cef.md)|Sets or returns the value of the  **On Key Press** box in the **Properties** window. Read/write **String**.|
| [OnKeyUp](77ebdf97-ae3f-73f4-d670-3c99d1f4f87d.md)|Sets or returns the value of the  **On Key Up** box in the **Properties** window. Read/write **String**.|
| [OnLostFocus](1606cb80-bf56-3766-d939-b545c2738e17.md)|Sets or returns the value of the  **On Lost Focus** box in the **Properties** window of the specified object. Read/write **String**.|
| [OnMouseDown](2392c2eb-6c3b-047b-1e4e-2772989ba1f2.md)|Sets or returns the value of the  **On Mouse Down** box in the **Properties** window. Read/write **String**.|
| [OnMouseMove](7201a61b-5b69-c13f-63bf-a2a5f329ecc5.md)|Sets or returns the value of the  **On Mouse Move** box in the **Properties** window. Read/write **String**.|
| [OnMouseUp](acd5de89-de56-e7c4-1a5d-cc560c5cffb6.md)|Sets or returns the value of the  **On Mouse Up** box in the **Properties** window. Read/write **String**.|
| [OnUndo](fa62ba10-c8e8-f4d4-5d48-ab73c074f2ef.md)|Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **Undo**event occurs. Read/write..|
| [Parent](e07da876-e24c-0828-e986-d13a0cb1f78e.md)|Returns the parent object for the specified object. Read-only.|
| [PostalAddress](04fb29c5-909c-a0b8-a4aa-7701abc07037.md)|You can use the  **PostalAddress Property** property to specify or determine the postal code and the Customer Barcode data corresponding to the address information displayed in a specified field/textbox. The PostalAddress Property wizard enables the setting of these properties. Read/write **String**.|
| [Properties](54a6372b-77db-5557-7af1-0c608f6d46a6.md)|Returns a reference to a control's ** [Properties](7e888aad-e783-dfc5-46df-9d92c89cfc35.md)**collection object. Read-only.|
| [ReadingOrder](1b53bb00-9252-ca99-c3b7-3a97d06552c4.md)|You can use the  **ReadingOrder** property to specify or determine the reading order of words in text. Read/write **Byte**.|
| [RightMargin](13f3fe1f-d5c3-33ac-9b9b-897df8ff5ba9.md)|Along with the  **TopMargin**,  **Left Margin**, and  **BottomMargin** properties, specifies the location of information displayed within text box control. Read/write **Integer**.|
| [RightPadding](7f9e2e21-1e36-01c1-f4e7-b3373644f9e5.md)|Gets or sets the amount of space (in inches) between the text box and its right gridline. Read/write  **Integer**.|
| [RunningSum](8918a58c-8c07-84dc-f43c-2486d54cd677.md)|You can use the  **RunningSum** property to calculate record-by-record or group-by-group totals in a report. Read/write **Byte**.|
| [ScrollBarAlign](5a8a77df-571a-7294-8be8-0ff2c4546131.md)|You can use the  **ScrollBarAlign** to specify or determine the alignment of a vertical scroll bar. Read/write **Byte**.|
| [ScrollBars](de3adbf1-4398-8782-0998-d392ab860669.md)|You can use the  **ScrollBars** property to specify whether scroll bars appear on a text box control. Read/write **Byte**.|
| [Section](76a43ccb-a199-b640-623c-d008b7d48e1c.md)|You can identify these controls by the section of a form or report where the control appears. Read/write  **Integer**.|
| [SelLength](0fb2371d-0f60-b0c7-5c4b-7a0689867b21.md)|The  **SelLength** property specifies or determines the number of characters selected in a text box. Read/write **Integer**.|
| [SelStart](51c773bb-2b70-b812-6b6a-9e062e493ebb.md)|The  **SelStart** property specifies or determines the starting point of the selected text or the position of the insertion point if no text is selected. Read/write **Integer**.|
| [SelText](1625b16f-8c2d-a563-6f66-a6714f5419ec.md)|The  **SelText** property returns a string containing the selected text. Read/write **String**.|
| [ShortcutMenuBar](620de877-2164-6426-90b8-c72a6db637fd.md)|You can use the  **ShortcutMenuBar** property to specify the shortcut menu that will appear when you right-click on the specified object. Read/write **String**.|
| [ShowDatePicker](5d65938b-ac7b-abbd-2e50-41f41c0b1558.md)|Gets or sets whether the date picker control is displayed for the specified text box. Read/write  **Integer**.|
| [SmartTags](200175d1-78a2-3036-72ba-4a85dfc21864.md)|Returns a  ** [SmartTags](79c0e84e-e0a1-35b8-b826-9d2cde3bd485.md)** collection that represents the collection of smart tags that have been added to a control. .|
| [SpecialEffect](9d34e61b-9ba9-02e0-4bd8-30da0a043a89.md)|You can use the  **SpecialEffect** property to specify whether special formatting will apply to the specified object. Read/write **Byte**.|
| [StatusBarText](18ae7a69-2e63-7896-1bff-da3f45b62c63.md)|You can use the  **StatusBarText** property to specify the text that is displayed in the status bar when a control is selected. Read/write **String**.|
| [TabIndex](d52e0839-e0aa-1b67-b075-115ad7b2f774.md)|You can use the  **TabIndex** property to specify a control's place in the tab order on a form or report. Read/write **Integer**.|
| [TabStop](ecb9ede6-e7a8-ca91-9ca3-3fad9de2ef90.md)|You can use the  **TabStop** property to specify whether you can use the TAB key to move the focus to a control. Read/write **Boolean**.|
| [Tag](9df21640-6bea-60a9-f9d0-dac90a60af1c.md)|Stores extra information about a form, report, section, or control needed by a Microsoft Access application. Read/write  **String**.|
| [Text](bb510c65-6d0d-468a-c5be-f325d86c2c7f.md)|You can use the  **Text** property to set or return the text contained in a text box. Read/write **String**.|
| [TextAlign](2b6e5ad7-02f5-4e33-47a4-87882a3113b2.md)|The  **TextAlign** property specifies the text alignment in new controls. Read/write **Byte**.|
| [TextFormat](3d164782-9d9c-5462-ac40-51772d475407.md)|Gets or sets whether rich text is displayed in the specified text box. Read/write  ** [AcTextFormat](cce0f7f5-ec7d-b80b-71a4-95052b6b7af1.md)**.|
| [ThemeFontIndex](2abe2063-4658-e441-7a7d-c4d226063172.md)|Gets or sets the font index that represents a font in the applied theme associated with the  **FontName** property of the specified object. Read/write **Long**.|
| [Top](6a220cec-d42c-05e3-c8c0-078687813a8d.md)|You can use the  **Top** property to specify an object's location on a form or report. Read/write **Integer**. .|
| [TopMargin](cd56b2b2-8bb5-b3cf-bacf-13d311e5479b.md)|Along with the  **LeftMargin**,  **RightMargin**, and  **BottomMargin** properties, specifies the location of information displayed within a text box control. Read/write **Integer**.|
| [TopPadding](fd6420f1-c3d9-2374-7b3c-e1fa5dd8199a.md)|Gets or sets the amount of space (in inches) between the text box and its top gridline. Read/write  **Integer**.|
| [ValidationRule](e481fba1-7e08-f8da-b644-5e38c2bf445e.md)|You can use the  **ValidationRule** property to specify requirements for data entered into a record, field, or control. When data is entered that violates the **ValidationRule** setting, you can use the **ValidationText** property to specify the message to be displayed to the user. Read/write **String**.|
| [ValidationText](5d3ab2a3-9166-714f-a0c2-d56d42b19ebc.md)|Use the  **ValidationText** property to specify a message to be displayed to the user when data is entered that violates a **ValidationRule** setting for a record, field, or control. Read/write **String**.|
| [Value](4cb4c33f-dd96-0309-f30b-8e445d123756.md)|Determines or specifies the text in the text box. Read/write  **Variant**.|
| [Vertical](40b9f9c0-daab-5562-395e-3e785d316d91.md)|You can use the  **Vertical** property to set a form control for vertical display and editing or set a report control for vertical display and printing. Read/write **Boolean**.|
| [VerticalAnchor](b515b37f-0566-0483-d387-8bc02c7be980.md)|Gets or sets an  ** [AcVerticalAnchor](08f16c8b-1566-cfad-795a-cb65a91c4e52.md)** constant that indicates how the specified text box is anchored vertically within its layout. Read/write.|
| [Visible](af1b9264-53f9-bf4c-2f05-049288a1d3d5.md)|Returns or sets whether the object is visible. Read/write  **Boolean**.|
| [Width](0bb72524-6682-f783-e9f9-4fd34a757a40.md)|Gets or sets the width of the specified object in twips. Read/write  **Integer**.|
