---
title: ComboBox Object (Access)
keywords: vbaac10.chm11545
f1_keywords:
- vbaac10.chm11545
ms.prod: access
api_name:
- Access.ComboBox
ms.assetid: 1cf508d5-023e-eb38-3991-71e82b2a4e7e
ms.date: 06/08/2017
---


# ComboBox Object (Access)

This object corresponds to a combo box control. The combo box control combines the features of a text box and a list box. Use a combo box when you want the option of either typing a value or selecting a value from a predefined list.


## Remarks


|||
|:-----|:-----|
|**Control**:|**Tool**:|
|![Combo box control](images/t-combox_ZA06053980.gif)|![Combo box tool](images/a_combobox_ZA06047114.gif)|

In Form view, Microsoft Access doesn't display the list until you click the combo box's arrow.

If you have Control Wizards on before you select the combo box tool, you can create a combo box with a wizard. To turn Control Wizards on or off, click the  **Control Wizards** tool in the toolbox.

The setting of the  **LimitToList** property determines whether you can enter values that aren't in the list.

The list can be single- or multiple-column, and the columns can appear with or without headings.

 **Link provided by:** Luke Chung, [FMS, Inc.](http://www.fmsinc.com/)


- [Tips and Techniques for Using and Validating Combo Boxes](http://www.fmsinc.com/free/NewTips/Access/ComboBox/AccessComboBox.asp)
    
 **Links provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The [UtterAccess](http://www.utteraccess.com) community


- [Combo Box](http://www.utteraccess.com/wiki/index.php/Combo_Box)
    
- [Cascading Combo Boxes](http://www.utteraccess.com/wiki/index.php/Cascading_Combo_Boxes)
    
- [Cascading Combo Boxes: Demo](http://www.utteraccess.com/wiki/index.php/Cascading_Combo_Boxes:_Demo)
    
- [Cascading Combo Boxes - Leaving Null Values](http://www.utteraccess.com/wiki/index.php/Cascade_Combo_Leaving_Null_Values)
    
- [Forms: Populate Controls/Text Boxes Based on Combobox Selection](http://www.utteraccess.com/wiki/index.php/Forms:_Populate_Controls/Text_Boxes_Based_on_Combobox_Selection)
    

## Example

The following example shows how to use multiple  **ComboBox** controls to supply criteria for a query.

 **Sample code provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The [UtterAccess](http://www.utteraccess.com) community

UtterAccess members can download a database that contains this example from [here](http://www.utteraccess.com/forum/Dynamic-Query-Examples-t1405533.html).




```
Private Sub cmdSearch_Click()
    Dim db As Database
    Dim qd As QueryDef
    Dim vWhere As Variant
    
    Set db = CurrentDb()
    
    On Error Resume Next
    db.QueryDefs.Delete "Query1"
    On Error GoTo 0
    
    vWhere = Null
    vWhere = vWhere &amp; " AND [PymtTypeID]=" + Me.cboPaymentTypes
    vWhere = vWhere &amp; " AND [RefundTypeID]=" + Me.cboRefundType
    vWhere = vWhere &amp; " AND [RefundCDMID]=" + Me.cboRefundCDM
    vWhere = vWhere &amp; " AND [RefundOptionID]=" + Me.cboRefundOption
    vWhere = vWhere &amp; " AND [RefundCodeID]=" + Me.cboRefundCode
    
    If Nz(vWhere, "") = "" Then
        MsgBox "There are no search criteria selected." &amp; vbCrLf &amp; vbCrLf &amp; _
        "Search Cancelled.", vbInformation, "Search Canceled."
        
    Else
        Set qd = db.CreateQueryDef("Query1", "SELECT * FROM tblRefundData WHERE " &amp; _
        Mid(vWhere, 6))
        
        db.Close
        Set db = Nothing
        
        DoCmd.OpenQuery "Query1", acViewNormal, acReadOnly
    End If
End Sub
```



The following example shows how to set the  **RowSource** property of a combo box when a form is loaded. When the form is displayed, the items stored in the **Departments** field of the **tblDepartment** combo box are displayed in the **cboDept** combo box.

 **Sample code provided by:**
![MVP Contributor](images/odc_OfficeTA_33px_MVPContrib.jpg) Bill Jelen,[MrExcel.com](http://www.mrexcel.com/)




```
Private Sub Form_Load()
    Me.Caption = "Today is " &amp; Format$(Date, "dddd mmm-d-yyyy")
    Me.RecordSource = "tblDepartments"
    DoCmd.Maximize  
    txtDept.ControlSource = "Department"
    cmdClose.Caption = "&amp;Close"
    cboDept.RowSourceType = "Table/Query"
    cboDept.RowSource = "SELECT Department FROM tblDepartments"
End Sub
```



The following example show how to create a combo box that is bound to one column while displaying another. Setting the  **ColumnCount** property to 2 specifies that the **cboDept** combo box will display the first two columns of the data source specified by the **RowSource** property. Setting the **BoundColumn** property to 1 specifies that the value stored in the first column will be returned when you inspect the value of the combo box.

The  **ColumnWidths** property specifies the width of the two columns. By setting the width of the first column to **0in.**, the first column is not displayed in the combo box.

 **Sample code provided by:**
![MVP Contributor](images/odc_OfficeTA_33px_MVPContrib.jpg) Bill Jelen,[MrExcel.com](http://www.mrexcel.com/)




```
Private Sub cboDept_Enter()
    With cboDept
        .RowSource = "SELECT * FROM tblDepartments ORDER BY Department"
        .ColumnCount = 2
        .BoundColumn = 1
        .ColumnWidths = "0in.;1in."
    End With
End Sub
```

The following example shows how to add an item to a bound combo box.

 **Sample code provided by:** The [Microsoft Access 2010 Programmer's Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.html)


```
Private Sub cboMainCategory_NotInList(NewData As String, Response As Integer)

    On Error GoTo Error_Handler
    Dim intAnswer As Integer
    intAnswer = MsgBox("""" &amp; NewData &amp; """ is not an approved category. " &amp; vbcrlf _
        &amp; "Do you want to add it now?" _ vbYesNo + vbQuestion, "Invalid Category")

    Select Case intAnswer
        Case vbYes
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO tlkpCategoryNotInList (Category) "
                &amp; _ "Select """ &amp; NewData &amp; """;"
            DoCmd.SetWarnings True
            Response = acDataErrAdded
        Case vbNo
            MsgBox "Please select an item from the list.", _
                vbExclamation + vbOKOnly, "Invalid Entry"
            Response = acDataErrContinue

    End Select

    Exit_Procedure:
        DoCmd.SetWarnings True
        Exit Sub

    Error_Handler:
        MsgBox Err.Number &amp; ", " &amp; Error Description
        Resume Exit_Procedure
        Resume

End Sub
```

## Events

|**Name**|
|:-----|
|[AfterUpdate](http://msdn.microsoft.com/library/89b45f0c-5ab1-889e-bd26-a34281b49b9e%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/4c4513e2-8596-fc44-a333-ae6ea9dce937%28Office.15%29.aspx)|
|[Change](http://msdn.microsoft.com/library/ed16e578-85f8-12ae-2adc-03df45dadc47%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/7d5d4a8f-a447-8d55-1517-8ffa71f0a123%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/76f71a30-6e66-1677-4d09-24c2a420d404%28Office.15%29.aspx)|
|[Dirty](http://msdn.microsoft.com/library/15273cae-5466-0e5c-1783-796458ceb34d%28Office.15%29.aspx)|
|[Enter](http://msdn.microsoft.com/library/b41de5d4-7037-c020-9f6d-8aeba7984dbe%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/47f37eb3-c0c1-457f-31ec-3b33b02ba986%28Office.15%29.aspx)|
|[GotFocus](http://msdn.microsoft.com/library/7ba8de56-6306-d1b3-288f-687c0f6f6566%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/948985ea-6a7b-ec42-1f09-1ac900962136%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/8417f6e9-7727-c619-0ceb-e68dadd08e3f%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/ab8e8950-7ed3-7c8d-340d-fd9110a103d1%28Office.15%29.aspx)|
|[LostFocus](http://msdn.microsoft.com/library/e25f07da-2399-0258-b3be-bf1fd6a1e171%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/3c780064-35e6-362c-4624-3c326f57080c%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/73c929d1-bd21-3f79-4291-b5d04357ad9f%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/400e2f82-9177-d084-680e-32673164e457%28Office.15%29.aspx)|
|[NotInList](http://msdn.microsoft.com/library/1c8a73e1-ca69-ae31-c86a-c1dc6cb3e860%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/d1064051-bbf9-ce00-c43e-19775879185c%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddItem](http://msdn.microsoft.com/library/dd247136-f29b-b5e2-1e09-c5a808da803f%28Office.15%29.aspx)|
|[Dropdown](http://msdn.microsoft.com/library/f6a4bb90-be0a-930f-56e7-bc6833af73c3%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/74581db7-0039-e59b-4371-9457c198e39d%28Office.15%29.aspx)|
|[RemoveItem](http://msdn.microsoft.com/library/9e70c221-e2fd-d006-1460-2b1902b0b0ea%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/25203ee2-5e4b-4f23-a596-ff3a7ddb0014%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/d17e91d3-5478-942e-41b9-7404e5dfac50%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/720b5380-d673-4cc0-9425-fc6ae5ae7fb5%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/f5d21af8-0e6d-1517-baf8-020bde595b76%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddColon](http://msdn.microsoft.com/library/b4d2c2b4-f638-0327-bbe3-da0f7fb1502c%28Office.15%29.aspx)|
|[AfterUpdate](http://msdn.microsoft.com/library/13261b5e-6c52-f666-14ff-06c20d23c504%28Office.15%29.aspx)|
|[AllowAutoCorrect](http://msdn.microsoft.com/library/ebf48367-20fb-14be-7082-a2d9de923c51%28Office.15%29.aspx)|
|[AllowValueListEdits](http://msdn.microsoft.com/library/558ba7aa-b3b2-4fe8-7338-8e9fbef19159%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/21c195f2-7a1f-a945-504e-6c1a7fa7f01f%28Office.15%29.aspx)|
|[AutoExpand](http://msdn.microsoft.com/library/0b3fabf8-4004-0868-3ddc-aef297514324%28Office.15%29.aspx)|
|[AutoLabel](http://msdn.microsoft.com/library/97f04ad4-fac6-bebe-3eab-720a7e9cd999%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/63e7e016-f06f-4426-748a-b01d5550f727%28Office.15%29.aspx)|
|[BackShade](http://msdn.microsoft.com/library/d1846516-4f38-67bb-3e8c-41bd79ac7a30%28Office.15%29.aspx)|
|[BackStyle](http://msdn.microsoft.com/library/1def822f-6b4a-8384-9d81-72b30e680908%28Office.15%29.aspx)|
|[BackThemeColorIndex](http://msdn.microsoft.com/library/fd8dc917-9cb7-94ca-5bcf-0d8e1f741fbb%28Office.15%29.aspx)|
|[BackTint](http://msdn.microsoft.com/library/37f48215-abce-1628-7efc-ace0d4761873%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/ce748fb1-4f8d-9e96-f77c-5dfc54dfee48%28Office.15%29.aspx)|
|[BorderColor](http://msdn.microsoft.com/library/1863d1e2-b865-5de5-471e-0d9124f34354%28Office.15%29.aspx)|
|[BorderShade](http://msdn.microsoft.com/library/2cb4cc56-c40f-59ce-a989-e792cad915ba%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/ad259ff2-a4b5-14af-3478-3dfc638acab5%28Office.15%29.aspx)|
|[BorderThemeColorIndex](http://msdn.microsoft.com/library/d17a61fb-5e27-5fcf-37ca-ef896b62fe98%28Office.15%29.aspx)|
|[BorderTint](http://msdn.microsoft.com/library/83bb493b-c15e-dcdf-7118-4bdb12f5e264%28Office.15%29.aspx)|
|[BorderWidth](http://msdn.microsoft.com/library/f8eddb71-d8ad-cca1-10ed-e6d3fb10e41a%28Office.15%29.aspx)|
|[BottomMargin](http://msdn.microsoft.com/library/aaabb534-e328-fb1d-92bc-4cbab0e0469f%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/1e65383f-3928-7cbe-e4e3-e244d877043b%28Office.15%29.aspx)|
|[BoundColumn](http://msdn.microsoft.com/library/ba2b5807-5f5a-52bb-d5d3-db7525bccba4%28Office.15%29.aspx)|
|[CanGrow](http://msdn.microsoft.com/library/0abc0d9c-35dc-ea5f-dcb1-dbfe37b7a143%28Office.15%29.aspx)|
|[CanShrink](http://msdn.microsoft.com/library/6f74e442-0b65-1d15-b247-6e12b9a08f1e%28Office.15%29.aspx)|
|[Column](http://msdn.microsoft.com/library/3b410a44-9055-e2c7-b921-4b364f68041b%28Office.15%29.aspx)|
|[ColumnCount](http://msdn.microsoft.com/library/76db2415-ee22-89c6-6753-f20d636d41f8%28Office.15%29.aspx)|
|[ColumnHeads](http://msdn.microsoft.com/library/b2066599-043f-bcad-5f7e-31f66cb33810%28Office.15%29.aspx)|
|[ColumnHidden](http://msdn.microsoft.com/library/631ab036-cdbe-c471-a2bb-10172032bfcf%28Office.15%29.aspx)|
|[ColumnOrder](http://msdn.microsoft.com/library/a32260cc-33b0-0811-1102-63843d5d2a21%28Office.15%29.aspx)|
|[ColumnWidth](http://msdn.microsoft.com/library/938c3d16-5c71-1c36-097f-61782b8ed358%28Office.15%29.aspx)|
|[ColumnWidths](http://msdn.microsoft.com/library/cd7894fd-e989-4f17-d779-073c8ef6c664%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/8f936303-1d90-d1cd-320f-de175df686cf%28Office.15%29.aspx)|
|[ControlSource](http://msdn.microsoft.com/library/d67db09e-d8c5-4605-2789-c75ac652ee0b%28Office.15%29.aspx)|
|[ControlTipText](http://msdn.microsoft.com/library/8562dde5-4bc7-92fb-347b-dd45e0eb413a%28Office.15%29.aspx)|
|[ControlType](http://msdn.microsoft.com/library/2826c41c-ef98-f474-10d2-3b181daf041d%28Office.15%29.aspx)|
|[DecimalPlaces](http://msdn.microsoft.com/library/5d57d9b7-12bd-2555-242e-204fd8dd48be%28Office.15%29.aspx)|
|[DefaultValue](http://msdn.microsoft.com/library/9c8a001f-ba06-f5c4-654d-7f37cabec14e%28Office.15%29.aspx)|
|[DisplayAsHyperlink](http://msdn.microsoft.com/library/7abd6406-9276-e2d2-de15-9450deb94973%28Office.15%29.aspx)|
|[DisplayWhen](http://msdn.microsoft.com/library/cfe800d7-290d-3f5c-fb48-cbc0628cefcd%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/69952de0-af27-32fe-0567-6558e85f53c5%28Office.15%29.aspx)|
|[EventProcPrefix](http://msdn.microsoft.com/library/79af6ac6-8876-ff72-16a8-5ec81ab6a0f8%28Office.15%29.aspx)|
|[FontBold](http://msdn.microsoft.com/library/c125e323-8e4a-4814-3dd6-cc5bef6ebf96%28Office.15%29.aspx)|
|[FontItalic](http://msdn.microsoft.com/library/57a1a671-1001-e614-ff10-8b5e7a16ca43%28Office.15%29.aspx)|
|[FontName](http://msdn.microsoft.com/library/0869818d-225e-c46b-39f3-5d500374361b%28Office.15%29.aspx)|
|[FontSize](http://msdn.microsoft.com/library/6dcd4b7e-01ec-a44d-4ceb-eecaa02ed1d7%28Office.15%29.aspx)|
|[FontUnderline](http://msdn.microsoft.com/library/54ee770c-4e75-fbc7-0453-99fc2c2456c1%28Office.15%29.aspx)|
|[FontWeight](http://msdn.microsoft.com/library/4e1cf348-4114-788d-34a6-c0b17152ee4b%28Office.15%29.aspx)|
|[ForeColor](http://msdn.microsoft.com/library/32327754-0132-0e04-ef61-f94fa6b095f3%28Office.15%29.aspx)|
|[ForeShade](http://msdn.microsoft.com/library/7bf41b29-6f65-d82d-bea7-1f988381c946%28Office.15%29.aspx)|
|[ForeThemeColorIndex](http://msdn.microsoft.com/library/89138cf8-23f1-e795-1d6c-951299c3d90e%28Office.15%29.aspx)|
|[ForeTint](http://msdn.microsoft.com/library/d855214b-df01-7158-75ea-1fc974c9b60b%28Office.15%29.aspx)|
|[Format](http://msdn.microsoft.com/library/9bb18f6a-0a25-9bbf-88ba-adf603c11826%28Office.15%29.aspx)|
|[FormatConditions](http://msdn.microsoft.com/library/0eeb11b4-453b-4a00-0a1f-92e3108ab2b9%28Office.15%29.aspx)|
|[GridlineColor](http://msdn.microsoft.com/library/b0085ca5-3d6c-35c2-fe19-ce7a7776d216%28Office.15%29.aspx)|
|[GridlineShade](http://msdn.microsoft.com/library/286746a1-0098-8991-0074-fe6fa0ceff0a%28Office.15%29.aspx)|
|[GridlineStyleBottom](http://msdn.microsoft.com/library/36ccbfbb-60e4-8d2e-6f15-4b1d22a732bf%28Office.15%29.aspx)|
|[GridlineStyleLeft](http://msdn.microsoft.com/library/00bb1dfe-ce6f-2bcc-75c6-bd1088d1c656%28Office.15%29.aspx)|
|[GridlineStyleRight](http://msdn.microsoft.com/library/d9b7b183-4fc8-26d2-112a-af65fb0bad8d%28Office.15%29.aspx)|
|[GridlineStyleTop](http://msdn.microsoft.com/library/a68481b9-2e6f-fb25-c87f-4e94416aa1dd%28Office.15%29.aspx)|
|[GridlineThemeColorIndex](http://msdn.microsoft.com/library/5ff8140e-4c6a-b719-3fe5-a9a64bb04771%28Office.15%29.aspx)|
|[GridlineTint](http://msdn.microsoft.com/library/b0f0aad3-7355-d594-8874-ec7229c1dff1%28Office.15%29.aspx)|
|[GridlineWidthBottom](http://msdn.microsoft.com/library/861d6f35-9c39-fdad-26c9-bf5c60499fbf%28Office.15%29.aspx)|
|[GridlineWidthLeft](http://msdn.microsoft.com/library/cbcc62ab-90f1-64ed-161f-fba7b465d148%28Office.15%29.aspx)|
|[GridlineWidthRight](http://msdn.microsoft.com/library/4a3292f9-1371-38ef-eca6-616d623a34b8%28Office.15%29.aspx)|
|[GridlineWidthTop](http://msdn.microsoft.com/library/0da57905-51fd-f9fe-374d-2289ad38ff9c%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/9cd1dd69-e7b2-800e-301c-742dc4804d28%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/87ac2391-64d5-5257-d7e8-2ce45b37eeb7%28Office.15%29.aspx)|
|[HideDuplicates](http://msdn.microsoft.com/library/79b64a87-d98e-76a1-e3c7-57796cb1c173%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/9dfa13fe-d062-d7a8-87a3-2e6e37fce5e9%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/239bfe2c-549f-5148-15bb-9e99348cb7ec%28Office.15%29.aspx)|
|[IMEHold](http://msdn.microsoft.com/library/ab128652-1de6-e4a2-4bc5-99936b3fee7f%28Office.15%29.aspx)|
|[IMEMode](http://msdn.microsoft.com/library/117b9f33-004e-40f9-7ec9-bb397fda33c0%28Office.15%29.aspx)|
|[IMESentenceMode](http://msdn.microsoft.com/library/f56b97cb-73c9-f5ff-a467-6e7dcd64e613%28Office.15%29.aspx)|
|[InheritValueList](http://msdn.microsoft.com/library/9189cd24-c4f2-c9a4-289f-0515d4b7fd45%28Office.15%29.aspx)|
|[InputMask](http://msdn.microsoft.com/library/da40a7cb-d962-dcb7-e536-c90c2753aaed%28Office.15%29.aspx)|
|[InSelection](http://msdn.microsoft.com/library/0e0bf471-8d24-52a8-c14c-3e4441a2fc8d%28Office.15%29.aspx)|
|[IsHyperlink](http://msdn.microsoft.com/library/005d21a1-c44c-c0a6-f625-2b3f8f4f8f91%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/8dddc896-d0fa-084d-948d-670e7957e737%28Office.15%29.aspx)|
|[ItemData](http://msdn.microsoft.com/library/9e9a6aab-472a-5715-f7f4-5957b1dcf717%28Office.15%29.aspx)|
|[ItemsSelected](http://msdn.microsoft.com/library/7e4f6f12-3d97-b36a-1211-8c95b43642e6%28Office.15%29.aspx)|
|[KeyboardLanguage](http://msdn.microsoft.com/library/5eb0e03c-c931-45b5-7801-d790c4678768%28Office.15%29.aspx)|
|[LabelAlign](http://msdn.microsoft.com/library/a64df157-b9d6-a426-169e-a0878598b9d9%28Office.15%29.aspx)|
|[LabelX](http://msdn.microsoft.com/library/3878f4b3-6f0d-a857-1988-9fae59c0302b%28Office.15%29.aspx)|
|[LabelY](http://msdn.microsoft.com/library/8577bdf8-b941-688a-fae3-a74aba173996%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/b4a2f19f-de56-b82d-4dab-3c22bc41cf94%28Office.15%29.aspx)|
|[LayoutID](http://msdn.microsoft.com/library/eba00f13-38d0-ee29-5d9e-74dfa21f9443%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/d6eeae85-bc8f-c56e-4014-d1a95e32d18e%28Office.15%29.aspx)|
|[LeftMargin](http://msdn.microsoft.com/library/b478bd94-b36b-b100-f0a0-10040af55b9d%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/4315a484-56ec-efb7-96bb-4eaea1c5c5b3%28Office.15%29.aspx)|
|[LimitToList](http://msdn.microsoft.com/library/885ed814-6e04-b9f1-0acb-3ded28e00f93%28Office.15%29.aspx)|
|[ListCount](http://msdn.microsoft.com/library/5363c040-1845-6e5c-7306-e48f392f0da9%28Office.15%29.aspx)|
|[ListIndex](http://msdn.microsoft.com/library/2165ba25-f129-3378-fb49-ea26ca446e9e%28Office.15%29.aspx)|
|[ListItemsEditForm](http://msdn.microsoft.com/library/5db884d4-4d9f-23b5-9e3a-f6de953a4800%28Office.15%29.aspx)|
|[ListRows](http://msdn.microsoft.com/library/b418e124-71b6-2ffb-101d-b56aadebb1fc%28Office.15%29.aspx)|
|[ListWidth](http://msdn.microsoft.com/library/488a36f0-3ab1-1bb1-ff48-3e5d33a55139%28Office.15%29.aspx)|
|[Locked](http://msdn.microsoft.com/library/6ef9a63c-9b00-126f-f662-0d23d672cfa2%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/d43da3b5-3189-b5be-37e6-6e1fdf99787b%28Office.15%29.aspx)|
|[NumeralShapes](http://msdn.microsoft.com/library/93cb42d2-6274-3af4-0801-87ecf8eb4252%28Office.15%29.aspx)|
|[OldBorderStyle](http://msdn.microsoft.com/library/64e0535d-d64d-1114-e01e-3cb1bcc62b2f%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/ed5ddacb-c447-02b1-3de1-3762a7540bff%28Office.15%29.aspx)|
|[OnChange](http://msdn.microsoft.com/library/e3c26738-a14f-e379-d909-f4919bb37a20%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/9cb266d8-6e7d-80c9-c5e9-1d2406b7d54d%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/737b1fd2-8966-5417-4979-538fa0594ef9%28Office.15%29.aspx)|
|[OnDirty](http://msdn.microsoft.com/library/2ef8c314-65d2-a61d-70e1-c8f8c40d86a8%28Office.15%29.aspx)|
|[OnEnter](http://msdn.microsoft.com/library/be3b353e-7105-010a-0c6a-6c551dcf62d3%28Office.15%29.aspx)|
|[OnExit](http://msdn.microsoft.com/library/0bc23d67-a1c1-8140-1930-2a1d97008fb5%28Office.15%29.aspx)|
|[OnGotFocus](http://msdn.microsoft.com/library/6fb801bd-c2f9-e81d-24b7-0669ece6422d%28Office.15%29.aspx)|
|[OnKeyDown](http://msdn.microsoft.com/library/49921f2f-abab-692f-52ca-bbdf2ce04ae3%28Office.15%29.aspx)|
|[OnKeyPress](http://msdn.microsoft.com/library/ddd9e200-5578-3269-d2c8-5352684e5fab%28Office.15%29.aspx)|
|[OnKeyUp](http://msdn.microsoft.com/library/5ee1855c-c6ba-84a6-4cc8-586ee2b201e0%28Office.15%29.aspx)|
|[OnLostFocus](http://msdn.microsoft.com/library/95356ca4-d76d-9027-7330-b5d36ccf7afc%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/b0078538-a6b2-fcce-56f4-d38260694faa%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/4610d4e9-97a5-2091-095c-f8aa5d8ac427%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/bf6f673c-fd59-d411-9cd2-cf7820bb04b3%28Office.15%29.aspx)|
|[OnNotInList](http://msdn.microsoft.com/library/307e9f0c-6db7-b995-166b-060c697b9f6e%28Office.15%29.aspx)|
|[OnUndo](http://msdn.microsoft.com/library/848f5228-7238-6e56-af49-8334c821ec04%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/31dc2078-38ea-00a4-fcaa-626c4b940fbc%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/a37520ad-f42d-b6b3-4a04-fa528266f2ba%28Office.15%29.aspx)|
|[ReadingOrder](http://msdn.microsoft.com/library/83989cec-fcab-0b83-5b5a-5dedc1a77aea%28Office.15%29.aspx)|
|[Recordset](http://msdn.microsoft.com/library/6985fa39-de4c-3c5b-175b-d156f2730836%28Office.15%29.aspx)|
|[RightMargin](http://msdn.microsoft.com/library/4ee85481-4489-4f81-123b-54062c071b97%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/71089377-d206-24b0-be15-aca3e7f33c2e%28Office.15%29.aspx)|
|[RowSource](http://msdn.microsoft.com/library/1225e566-24e0-244d-09ae-e036c87f3141%28Office.15%29.aspx)|
|[RowSourceType](http://msdn.microsoft.com/library/dd1d6ea8-5479-4bf9-3317-0b95282c7d74%28Office.15%29.aspx)|
|[ScrollBarAlign](http://msdn.microsoft.com/library/ded4533c-2879-d57f-b6ff-cccd20a88090%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/fbf2993a-5360-10dd-1edd-2ab7ac2f567c%28Office.15%29.aspx)|
|[Selected](http://msdn.microsoft.com/library/fc643ebc-084a-c11c-2489-7d1504d5b17b%28Office.15%29.aspx)|
|[SelLength](http://msdn.microsoft.com/library/f465a2a0-2c4c-ac8b-0867-4033ca44e3f4%28Office.15%29.aspx)|
|[SelStart](http://msdn.microsoft.com/library/056196b5-828a-f276-da26-983c8b47cd05%28Office.15%29.aspx)|
|[SelText](http://msdn.microsoft.com/library/dc2b46d7-c688-c9b5-c44f-c490a91589fe%28Office.15%29.aspx)|
|[SeparatorCharacters](http://msdn.microsoft.com/library/7a91ecdf-35e0-d32c-7355-7656d9ed7ad1%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/e010eab3-c24c-b077-b8cd-6fbf708aa3a9%28Office.15%29.aspx)|
|[ShowOnlyRowSourceValues](http://msdn.microsoft.com/library/3400539d-64c2-bd83-6d82-b70bf9ba6654%28Office.15%29.aspx)|
|[SmartTags](http://msdn.microsoft.com/library/b86a8460-48c6-92ad-602b-1d736bb2c38c%28Office.15%29.aspx)|
|[SpecialEffect](http://msdn.microsoft.com/library/d9b82840-8914-7818-990d-9b595da4ba9f%28Office.15%29.aspx)|
|[StatusBarText](http://msdn.microsoft.com/library/430dabc5-ffdb-37fa-473d-359035bac761%28Office.15%29.aspx)|
|[TabIndex](http://msdn.microsoft.com/library/7e04fd77-8f25-eaad-c902-526f69226322%28Office.15%29.aspx)|
|[TabStop](http://msdn.microsoft.com/library/c22f2818-0c7f-522b-b17a-c4e32b26e99a%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/b128fa76-0ab7-48ae-398f-352be0d638ae%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/27f99e99-ce53-f5b9-61ed-1ffc4ba9cc4d%28Office.15%29.aspx)|
|[TextAlign](http://msdn.microsoft.com/library/c5de59ad-f41f-8f19-6056-16ca88a1937d%28Office.15%29.aspx)|
|[ThemeFontIndex](http://msdn.microsoft.com/library/07fec290-0bf3-138f-94cd-55d5979b2aca%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/17e9ca79-0b35-0c50-09f5-bbbc36482081%28Office.15%29.aspx)|
|[TopMargin](http://msdn.microsoft.com/library/fe3a17d8-c345-6dc6-5b26-5fc6f06632ac%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/3e92fdc3-79a3-8ed9-2547-5bb49df29852%28Office.15%29.aspx)|
|[ValidationRule](http://msdn.microsoft.com/library/3ea94f44-46fa-57a7-a9b4-a9e7b58e087b%28Office.15%29.aspx)|
|[ValidationText](http://msdn.microsoft.com/library/479d2067-caae-efcc-92a8-36aa68edb4a4%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/ac29f38d-1b88-0033-709d-6a40e57d188e%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/ac25f518-9954-7422-b0ac-61bb5a8ea758%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/68d69090-cf5f-5d24-de4d-a5304a41bd64%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/e5d7c087-c226-8c85-627f-d63c6b526f20%28Office.15%29.aspx)|

## About the Contributors
<a name="AboutContributors"> </a>

Luke Chung is the founder and president of FMS, Inc., a leading provider of custom database solutions and developer tools. 

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 


## See also
<a name="AboutContributors"> </a>


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)<br/>
[ComboBox Object Members](http://msdn.microsoft.com/library/d0d83ca3-3698-295e-5335-7d0816557d6b%28Office.15%29.aspx)
