---
title: Form Object (Access)
keywords: vbaac10.chm13686
f1_keywords:
- vbaac10.chm13686
ms.prod: access
api_name:
- Access.Form
ms.assetid: 72ef9219-142b-b690-b696-3eba9a5d4522
ms.date: 06/08/2017
---


# Form Object (Access)

A  **Form** object refers to a particular Microsoft Access form.


## Remarks

A  **Form** object is a member of the **Forms** collection, which is a collection of all currently open forms. Within the **Forms** collection, individual forms are indexed beginning with zero. You can refer to an individual **Form** object in the **Forms** collection either by referring to the form by name, or by referring to its index within the collection. If you want to refer to a specific form in the **Forms** collection, it's better to refer to the form by name because a form's collection index may change. If the form name includes a space, the name must be surrounded by brackets ([ ]).



|**Syntax**|**Example**|
|:-----|:-----|
|**AllForms** ! _formname_|AllForms!OrderForm|
|**AllForms** ![ _form name_]|AllForms![Order Form]|
|**AllForms** (" _formname_")|AllForms("OrderForm")|
|**AllForms** ( _formname_)|AllForms(0)|
Each  **Form** object has a **Controls** collection, which contains all controls on the form. You can refer to a control on a form either by implicitly or explicitly referring to the **Controls** collection. Your code will be faster if you refer to the **Controls** collection implicitly. The following examples show two of the ways you might refer to a control named **NewData** on the form called **OrderForm**:




```
 ' Implicit reference. 
Forms!OrderForm!NewData
```




```
' Explicit reference. 
Forms!OrderForm.Controls!NewData
```

The next two examples show how you might refer to a control named  **NewData** on a subform `ctlSubForm` contained in the form called **OrderForm**:




```
Forms!OrderForm.ctlSubForm.Form!Controls.NewData
```




```
Forms!OrderForm.ctlSubForm!NewData
```

 **Links provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) Luke Chung, [FMS, Inc.](http://www.fmsinc.com/)


- [Microsoft Access Form Tips and Avoiding Common Mistakes](http://www.fmsinc.com/tpapers/genaccess/formtips.html)
    
- [Microsoft Office Access 2007 Form Design Tips](http://www.fmsinc.com/tpapers/access/Forms/Access2007FormTips.html)
    
 **Links provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The [UtterAccess](http://www.utteraccess.com) community


- [Display Pictures on a Form](http://www.utteraccess.com/wiki/index.php/Display_Pictures_on_a_Form)
    
- [Display Related Data](http://www.utteraccess.com/wiki/index.php/Display_Related_Data)
    
- [Opening a Detail Form to Related Information](http://www.utteraccess.com/wiki/index.php/Forms:_Open_a_Detail_Form_to_Related_Information)
    
- [Forms: Populate Controls/Text Boxes Based on Combobox Selection](http://www.utteraccess.com/wiki/index.php/Forms:_Populate_Controls/Text_Boxes_Based_on_Combobox_Selection)
    
- [Referring To Properties And Controls On Subforms](http://www.utteraccess.com/wiki/index.php/Referring_To_Properties_And_Controls_On_Subforms)
    

## Example

The following example shows how to use  **TextBox** controls to supply date criteria for a query.

UtterAccess members can download a database that contains this example from [here](http://www.utteraccess.com/forum/Dynamic-Query-Examples-t1405533.html).

 **Sample code provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The [UtterAccess](http://www.utteraccess.com) community




```
Private Sub cmdSearch_Click()

   Dim db As DAO.Database
   Dim qd As QueryDef
   Dim vWhere As Variant

   Set db = CurrentDb()

   On Error Resume Next
   db.QueryDefs.Delete "Query1"
   On Error GoTo 0

   vWhere = Null

   vWhere = vWhere &amp; " AND [PayeeID]=" + Me.cboPayeeID

   If Nz(Me.txtEndDate, "") <> "" And Nz(Me.txtStartDate, "") <> "" Then
      vWhere = vWhere &amp; " AND [RefundProcessed] Between #" &amp; _
      Me.txtStartDate &amp; "# AND #" &amp; Me.txtEndDate &amp; "#"
   Else
      If Nz(Me.txtEndDate, "") = "" And Nz(Me.txtStartDate, "") <> "" Then
         vWhere = vWhere &amp; " AND [RefundProcessed]>=#" _
                  + Me.txtStartDate &amp; "#"
      Else
         If Nz(Me.txtEndDate, "") <> "" And Nz(Me.txtStartDate, "") = "" Then
            vWhere = vWhere &amp; " AND [RefundProcessed] <=#" _
                     + Me.txtEndDate &amp; "#"
      End If
     End If
   End If

   If Nz(vWhere, "") = "" Then
      MsgBox "There are no search criteria selected." &amp; vbCrLf &amp; vbCrLf &amp; _
             "Search Cancelled.", vbInformation, "Search Canceled."
   Else
      Set qd = db.CreateQueryDef("Query1", "SELECT * FROM tblRefundData? &amp; _
               " WHERE " &amp; Mid(vWhere, 6))
      db.Close
      Set db = Nothing

      DoCmd.OpenQuery "Query1", acViewNormal, acReadOnly
   End If
End Sub
```

The following example shows how to use the  **BeforeUpdate** event of a form to require that a value be entered into one control when another control also has data.

 **Sample code provided by:** The [Microsoft Access 2010 Programmer's Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.html)




```
Private Sub Form_BeforeUpdate(Cancel As Integer)
If (IsNull(Me.FieldOne)) Or (Me.FieldOne.Value =  "") Then
    ' No action required
Else
    If (IsNull(Me.FieldTwo)) or (Me.FieldTwo.Value = "") Then
        MsgBox "You must provide data for field 'FieldTwo', " &amp; _
            "if a value is entered in FieldOne", _
            vbOKOnly, "Required Field"
        Me.FieldTwo.SetFocus
        Cancel = True
        Exit Sub
    End If
End If

End Sub
```

The following example shows how to use the  **OpenArgs** property to prevent a form from being opened from the Navigation Pane.




```
Private Sub Form_Open(Cancel As Integer)

If Me.OpenArgs() <> "Valid User" Then
    MsgBox "You are not authorized to use this form!", _
        vbExclamation + vbOKOnly, "Invalid Access"
    Cancel = True
End If
End Sub
```

The following example shows how to use the  _WhereCondition_ argument of the **OpenForm** method to filter the records displayed on a form as it is opened.




```
Private Sub cmdShowOrders_Click()
If Not Me.NewRecord Then
    DoCmd.OpenForm "frmOrder", _
        WhereCondition:="CustomerID=" &amp; Me.txtCustomerID
End If
End Sub
```


## Events



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/1409c52b-8a77-0e0d-1a26-7dc4ce8bb320%28Office.15%29.aspx)|
|[AfterDelConfirm](http://msdn.microsoft.com/library/49f6f575-6f67-08b0-a2aa-913c8182cbe9%28Office.15%29.aspx)|
|[AfterFinalRender](http://msdn.microsoft.com/library/89f9cbb5-f002-4783-dc70-17878763e486%28Office.15%29.aspx)|
|[AfterInsert](http://msdn.microsoft.com/library/07140c13-ce7c-91f2-7451-d7f834653ef2%28Office.15%29.aspx)|
|[AfterLayout](http://msdn.microsoft.com/library/3b500c32-e1aa-ad06-432f-981253767c3d%28Office.15%29.aspx)|
|[AfterRender](http://msdn.microsoft.com/library/3232d72f-4dd4-9797-d9cb-5ac616c68c71%28Office.15%29.aspx)|
|[AfterUpdate](http://msdn.microsoft.com/library/b622d8c9-4802-a915-5cd4-f8a91ba57099%28Office.15%29.aspx)|
|[ApplyFilter](http://msdn.microsoft.com/library/c8aafdbf-1693-21cf-5bdd-1ea6d702aa58%28Office.15%29.aspx)|
|[BeforeDelConfirm](http://msdn.microsoft.com/library/36b9147a-6bfb-d386-117a-b65cc4659da8%28Office.15%29.aspx)|
|[BeforeInsert](http://msdn.microsoft.com/library/de0f6b1a-fc11-4000-2c0c-b0ad9ccfccc2%28Office.15%29.aspx)|
|[BeforeQuery](http://msdn.microsoft.com/library/07d9ba3f-25dc-f448-5c99-8c1e4ca5ab20%28Office.15%29.aspx)|
|[BeforeRender](http://msdn.microsoft.com/library/5661065e-472d-c073-948c-40b19c965848%28Office.15%29.aspx)|
|[BeforeScreenTip](http://msdn.microsoft.com/library/08e67747-9023-e880-c246-1aa9e9c447ed%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/b783fcab-f697-a464-820c-712eac46cb4b%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/43cf0568-c645-60eb-3c46-d9dd0b147d8d%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/e65fe7e0-efc1-dabc-4b2c-787af465ade0%28Office.15%29.aspx)|
|[CommandBeforeExecute](http://msdn.microsoft.com/library/4fb1c072-3781-8a52-bc9a-2e26d2738789%28Office.15%29.aspx)|
|[CommandChecked](http://msdn.microsoft.com/library/ec30f538-bbd2-9935-1ad9-5210f457b15f%28Office.15%29.aspx)|
|[CommandEnabled](http://msdn.microsoft.com/library/4a9ff0dc-5ed2-e841-97d3-a1c4a7ed4d42%28Office.15%29.aspx)|
|[CommandExecute](http://msdn.microsoft.com/library/b4b3bc8e-3e95-5120-ed7e-e17b2f8f23ba%28Office.15%29.aspx)|
|[Current](http://msdn.microsoft.com/library/44961599-2b0a-874e-be64-1e29f47f839f%28Office.15%29.aspx)|
|[DataChange](http://msdn.microsoft.com/library/026fddb4-2a43-095c-9460-98c12378735c%28Office.15%29.aspx)|
|[DataSetChange](http://msdn.microsoft.com/library/b266f48e-ccf9-1be1-edfb-f99892b09c97%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/bac58ee6-3fd8-696e-67d2-ab533760de11%28Office.15%29.aspx)|
|[Deactivate](http://msdn.microsoft.com/library/8702b30b-d38e-fcb6-141e-0ac4e53c63ad%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/89916f81-ec7a-f322-d4e6-a4a42db523cf%28Office.15%29.aspx)|
|[Dirty](http://msdn.microsoft.com/library/e0bcf968-7176-bd24-29c4-d3f014f57adb%28Office.15%29.aspx)|
|[Error](http://msdn.microsoft.com/library/ed8229fb-4169-8be5-dc2e-a543ca3bfff3%28Office.15%29.aspx)|
|[Filter](http://msdn.microsoft.com/library/265f3397-3dc9-21b3-ebac-55fb4e1261c0%28Office.15%29.aspx)|
|[GotFocus](http://msdn.microsoft.com/library/ded3dc26-938e-adb2-8017-e72dd83c9742%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/ceb66db0-695a-e3b1-f0f7-6c9bd9191b2b%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/62ebe014-955a-e47e-6506-f7be9aa44de6%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/dbef316c-6362-f330-0931-e109e1381907%28Office.15%29.aspx)|
|[Load](http://msdn.microsoft.com/library/a7547066-e1eb-6cdc-a170-2ee222081720%28Office.15%29.aspx)|
|[LostFocus](http://msdn.microsoft.com/library/09aa4b18-f4b2-024e-14c0-476faa76f209%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/09b8822b-9e13-3640-5fab-77fd00d8b68f%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/e255bb17-4997-9290-cd13-1a61666017b2%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/b397f122-24ec-18f9-779b-d8431664abc1%28Office.15%29.aspx)|
|[MouseWheel](http://msdn.microsoft.com/library/eec18d43-1cee-463c-37e6-760eccb0b890%28Office.15%29.aspx)|
|[OnConnect](http://msdn.microsoft.com/library/39966052-0e06-bde9-142f-ee74d16a9973%28Office.15%29.aspx)|
|[OnDisconnect](http://msdn.microsoft.com/library/b5b2a18b-d159-c122-c35e-fe749d755f0e%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/8638e6d9-29af-a007-44f5-9bada14adb29%28Office.15%29.aspx)|
|[PivotTableChange](http://msdn.microsoft.com/library/8b4a8c9a-c8a3-648d-968d-edcb7cb94956%28Office.15%29.aspx)|
|[Query](http://msdn.microsoft.com/library/f3070a6f-3064-b496-ff9f-4da165205f90%28Office.15%29.aspx)|
|[Resize](http://msdn.microsoft.com/library/de57e9bf-e4fd-174e-4d56-9ea813ab92ce%28Office.15%29.aspx)|
|[SelectionChange](http://msdn.microsoft.com/library/4c815a6d-4971-6cbd-16ad-905e93ec1b52%28Office.15%29.aspx)|
|[Timer](http://msdn.microsoft.com/library/395c62a1-5731-01b8-a4ea-852bfb30572f%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/fdcf98c1-c560-1c29-586d-6c4eb4a6ccd0%28Office.15%29.aspx)|
|[Unload](http://msdn.microsoft.com/library/13f1f7f4-9d69-128f-7e02-f3d3b99ec0f4%28Office.15%29.aspx)|
|[ViewChange](http://msdn.microsoft.com/library/a3788eca-783f-cb5d-1a7b-1c4a23648629%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[GoToPage](http://msdn.microsoft.com/library/932c15b9-57dd-0cf7-1db2-21364bc214ea%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/21529c39-70be-45ab-fe8a-b54b4f78b4c8%28Office.15%29.aspx)|
|[Recalc](http://msdn.microsoft.com/library/61786e64-dc17-b685-f427-fc7952d0320f%28Office.15%29.aspx)|
|[Refresh](http://msdn.microsoft.com/library/e7a15c34-d3ec-184f-8d03-3e264fcc60d0%28Office.15%29.aspx)|
|[Repaint](http://msdn.microsoft.com/library/ce386055-c4b7-9aa8-7f49-de0010467970%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/26d8d784-9348-6301-9bef-569d15668a0e%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/b0169892-0040-bb61-904f-0ea81eea681a%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/65c71211-8138-40cf-9b59-ceb087d2d7f0%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActiveControl](http://msdn.microsoft.com/library/0bb3cac4-fc88-cdd3-6bc4-1057b02d4eb5%28Office.15%29.aspx)|
|[AfterDelConfirm](http://msdn.microsoft.com/library/fcc1585b-ddb9-7b39-aa21-07de0e50ac00%28Office.15%29.aspx)|
|[AfterFinalRender](http://msdn.microsoft.com/library/c6e294f8-8cd9-1413-eff8-f2b033766326%28Office.15%29.aspx)|
|[AfterInsert](http://msdn.microsoft.com/library/95bc1f0d-a0fa-ffdd-ef5a-e6eb2a854feb%28Office.15%29.aspx)|
|[AfterLayout](http://msdn.microsoft.com/library/8d548e7b-6d68-4631-2c59-f6b8d39cbb12%28Office.15%29.aspx)|
|[AfterRender](http://msdn.microsoft.com/library/868b9a9d-a1e3-d460-fa7c-26cb5791c5ad%28Office.15%29.aspx)|
|[AfterUpdate](http://msdn.microsoft.com/library/5002727c-24bc-4067-0e5e-3c63b8b6427e%28Office.15%29.aspx)|
|[AllowAdditions](http://msdn.microsoft.com/library/8e440a96-7f9e-c009-5055-377c75999267%28Office.15%29.aspx)|
|[AllowDatasheetView](http://msdn.microsoft.com/library/81796b90-94dd-cd27-3613-a2050e2bce21%28Office.15%29.aspx)|
|[AllowDeletions](http://msdn.microsoft.com/library/abcbaa74-9a02-ab9c-613f-0cf6b9ce98b7%28Office.15%29.aspx)|
|[AllowEdits](http://msdn.microsoft.com/library/3f667914-3dcc-7d4e-ca66-4338fc08e63a%28Office.15%29.aspx)|
|[AllowFilters](http://msdn.microsoft.com/library/ca2998b5-d5e0-f1ba-f9da-d89ef24a3701%28Office.15%29.aspx)|
|[AllowFormView](http://msdn.microsoft.com/library/15dc69fc-d4ba-c8e3-d047-71f96c32fe02%28Office.15%29.aspx)|
|[AllowLayoutView](http://msdn.microsoft.com/library/70b273ef-60fa-00b8-b262-3c45e691ed42%28Office.15%29.aspx)|
|[AllowPivotChartView](http://msdn.microsoft.com/library/5585b530-d114-d07e-63cb-8d96dec458e8%28Office.15%29.aspx)|
|[AllowPivotTableView](http://msdn.microsoft.com/library/42bad4b4-7de1-f144-9482-2e114fc5cc4b%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/decaf70d-da61-6d74-9f60-8008a3c2e00f%28Office.15%29.aspx)|
|[AutoCenter](http://msdn.microsoft.com/library/a60f8783-5a25-42b5-da99-c5e2925fd6ea%28Office.15%29.aspx)|
|[AutoResize](http://msdn.microsoft.com/library/5ae98bc8-fa33-7e4b-31c8-ba22aa026a45%28Office.15%29.aspx)|
|[BeforeDelConfirm](http://msdn.microsoft.com/library/8926afb1-5a86-eddd-5b3f-68abe83fb076%28Office.15%29.aspx)|
|[BeforeInsert](http://msdn.microsoft.com/library/634b0480-ddb3-7ef7-b347-57ca9a4eebad%28Office.15%29.aspx)|
|[BeforeQuery](http://msdn.microsoft.com/library/40e763fd-897a-a0b1-72a9-d73ec628e397%28Office.15%29.aspx)|
|[BeforeRender](http://msdn.microsoft.com/library/f80035ac-4ce6-ac8a-203f-c36afab5cd01%28Office.15%29.aspx)|
|[BeforeScreenTip](http://msdn.microsoft.com/library/4829b972-de4e-f8dc-f19c-c6a52c7dd14b%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/b4b39ab8-e37c-8803-b6c3-032707342c92%28Office.15%29.aspx)|
|[Bookmark](http://msdn.microsoft.com/library/e214a924-9110-a3de-9812-b9ec5cbad8ed%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/a6c4d49b-4227-09e9-2999-6f8954bbeb39%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/44dcd309-7a07-c4b3-2d85-d1bc09f98843%28Office.15%29.aspx)|
|[ChartSpace](http://msdn.microsoft.com/library/e05f312f-d02b-bea5-7355-0a427834281c%28Office.15%29.aspx)|
|[CloseButton](http://msdn.microsoft.com/library/c87e3752-0a77-3e5e-9c82-20effaf0af1e%28Office.15%29.aspx)|
|[CommandBeforeExecute](http://msdn.microsoft.com/library/574568fa-e488-6d4d-a42f-07eb7c7f9536%28Office.15%29.aspx)|
|[CommandChecked](http://msdn.microsoft.com/library/4f3bb0fa-6f3f-4836-a0d0-06d480e1d194%28Office.15%29.aspx)|
|[CommandEnabled](http://msdn.microsoft.com/library/07e6989d-9739-e023-32e4-95147eb4bba3%28Office.15%29.aspx)|
|[CommandExecute](http://msdn.microsoft.com/library/b105b107-8123-5cfe-b87d-cb53518e3dba%28Office.15%29.aspx)|
|[ControlBox](http://msdn.microsoft.com/library/c4d9976c-631d-ae99-0c5d-e7008bbdadf9%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/08a31b50-b644-5912-d784-130f58298dd0%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/df11357c-b734-92b0-5793-aa64b4d960ef%28Office.15%29.aspx)|
|[CurrentRecord](http://msdn.microsoft.com/library/a682d187-0b9a-2fc3-3443-f2dcd6df4ca2%28Office.15%29.aspx)|
|[CurrentSectionLeft](http://msdn.microsoft.com/library/5c856f2a-f82c-2b67-6fc6-1773fc5ebe06%28Office.15%29.aspx)|
|[CurrentSectionTop](http://msdn.microsoft.com/library/d6f4f5f6-641f-3092-7d99-195c77722718%28Office.15%29.aspx)|
|[CurrentView](http://msdn.microsoft.com/library/d173222e-99d1-704e-ee3c-246263725706%28Office.15%29.aspx)|
|[Cycle](http://msdn.microsoft.com/library/621d7101-5237-b239-fcb3-2d942a0329b0%28Office.15%29.aspx)|
|[DataChange](http://msdn.microsoft.com/library/14fd4c9c-eb18-8f4d-ebd9-6f389523c4cf%28Office.15%29.aspx)|
|[DataEntry](http://msdn.microsoft.com/library/0a970904-10f9-d0c3-24d1-0b988725bb38%28Office.15%29.aspx)|
|[DataSetChange](http://msdn.microsoft.com/library/29f7f9a8-4dbd-9f69-7f4c-7f93add9f1b6%28Office.15%29.aspx)|
|[DatasheetAlternateBackColor](http://msdn.microsoft.com/library/d2a63a1f-0604-be80-5eef-67af92104bc2%28Office.15%29.aspx)|
|[DatasheetBackColor](http://msdn.microsoft.com/library/69734522-e570-86a5-f971-ce26ee4f88c3%28Office.15%29.aspx)|
|[DatasheetBorderLineStyle](http://msdn.microsoft.com/library/8a752955-97fe-933a-4130-62f63dbf6566%28Office.15%29.aspx)|
|[DatasheetCellsEffect](http://msdn.microsoft.com/library/3820b218-37b0-d5b5-bae2-8a179cc9b87a%28Office.15%29.aspx)|
|[DatasheetColumnHeaderUnderlineStyle](http://msdn.microsoft.com/library/9e689097-f3ed-bcda-9cc5-d423a3b92806%28Office.15%29.aspx)|
|[DatasheetFontHeight](http://msdn.microsoft.com/library/5cfcf818-eda0-f7ec-f224-ee52ae7d39c9%28Office.15%29.aspx)|
|[DatasheetFontItalic](http://msdn.microsoft.com/library/32fe51fa-ee36-2fc3-bb72-e61a4b43c19c%28Office.15%29.aspx)|
|[DatasheetFontName](http://msdn.microsoft.com/library/e6b963ca-7162-912e-e63d-1437904ec8f1%28Office.15%29.aspx)|
|[DatasheetFontUnderline](http://msdn.microsoft.com/library/a232a1a8-b537-4935-bd64-138548241c7c%28Office.15%29.aspx)|
|[DatasheetFontWeight](http://msdn.microsoft.com/library/6dd2c6d3-1f27-8b86-abf5-f5581fbe7d23%28Office.15%29.aspx)|
|[DatasheetForeColor](http://msdn.microsoft.com/library/9756ff09-67bf-edb9-d4b5-d414ec7c1e2a%28Office.15%29.aspx)|
|[DatasheetGridlinesBehavior](http://msdn.microsoft.com/library/692268ab-69f2-4891-e460-f091b43af962%28Office.15%29.aspx)|
|[DatasheetGridlinesColor](http://msdn.microsoft.com/library/92d07c1c-fc47-0049-7da3-a34ee56fbc83%28Office.15%29.aspx)|
|[DefaultControl](http://msdn.microsoft.com/library/f6444b54-cf68-0ec6-ebd0-041caba21d74%28Office.15%29.aspx)|
|[DefaultView](http://msdn.microsoft.com/library/bb44eca9-1576-794a-0558-f67e2d37559b%28Office.15%29.aspx)|
|[Dirty](http://msdn.microsoft.com/library/5806283f-7947-9e13-d6c3-49d519a8b521%28Office.15%29.aspx)|
|[DisplayOnSharePointSite](http://msdn.microsoft.com/library/f741a5df-5924-2756-409a-94a8fcf21809%28Office.15%29.aspx)|
|[DividingLines](http://msdn.microsoft.com/library/f8c62451-ccde-43f9-91f6-cdef38571c54%28Office.15%29.aspx)|
|[FastLaserPrinting](http://msdn.microsoft.com/library/a64775e5-174d-0349-d3f3-0009798d6462%28Office.15%29.aspx)|
|[FetchDefaults](http://msdn.microsoft.com/library/3bbe8c57-e9ff-419a-d2b4-93cb966d6f30%28Office.15%29.aspx)|
|[Filter](http://msdn.microsoft.com/library/5eb49f82-8519-981c-a663-9862736ac95f%28Office.15%29.aspx)|
|[FilterOn](http://msdn.microsoft.com/library/6ff59ffc-844b-24fc-925f-0331cfcf01ec%28Office.15%29.aspx)|
|[FilterOnLoad](http://msdn.microsoft.com/library/546f367f-fbe5-355f-ad48-784ac5f28c8d%28Office.15%29.aspx)|
|[FitToScreen](http://msdn.microsoft.com/library/5ef37719-ff3b-1f3d-1521-423633ceccc0%28Office.15%29.aspx)|
|[Form](http://msdn.microsoft.com/library/5e18dd48-f288-2b75-f42c-3a8b42f75b33%28Office.15%29.aspx)|
|[FrozenColumns](http://msdn.microsoft.com/library/5b595c5e-6a2e-e3d8-1ae8-a2f224eb5516%28Office.15%29.aspx)|
|[GridX](http://msdn.microsoft.com/library/ebc6a4d9-2f73-cf55-504f-a83aff1fecd4%28Office.15%29.aspx)|
|[GridY](http://msdn.microsoft.com/library/d767e7de-e3eb-0523-8782-26770f22a013%28Office.15%29.aspx)|
|[HasModule](http://msdn.microsoft.com/library/ba43a8c8-89f2-e744-ed99-082510dc8f3a%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/20cd50e1-5ac9-9739-d9e4-e5214706c61d%28Office.15%29.aspx)|
|[HelpFile](http://msdn.microsoft.com/library/72b416b1-5257-9560-ebc0-625abc3f7e85%28Office.15%29.aspx)|
|[HorizontalDatasheetGridlineStyle](http://msdn.microsoft.com/library/31467913-382f-031e-b030-68181a71d5e0%28Office.15%29.aspx)|
|[Hwnd](http://msdn.microsoft.com/library/894b2d6d-b11d-c674-e1e5-21ff33aeca36%28Office.15%29.aspx)|
|[InputParameters](http://msdn.microsoft.com/library/fc3e17a7-f62a-a6bb-c44a-f3a9d7efe6ac%28Office.15%29.aspx)|
|[InsideHeight](http://msdn.microsoft.com/library/7a49b4b4-1bbf-c0ad-d873-ff81f8b99929%28Office.15%29.aspx)|
|[InsideWidth](http://msdn.microsoft.com/library/c92954cd-0b8b-94d8-8826-684e886da0a2%28Office.15%29.aspx)|
|[KeyPreview](http://msdn.microsoft.com/library/f9153ec0-8b6e-60d5-8541-100e2ad1705e%28Office.15%29.aspx)|
|[LayoutForPrint](http://msdn.microsoft.com/library/fd8c8112-186a-3f77-06ef-783bf48a7052%28Office.15%29.aspx)|
|[MaxRecButton](http://msdn.microsoft.com/library/6f5ea968-1f79-1fbc-86e1-fff034dcc827%28Office.15%29.aspx)|
|[MaxRecords](http://msdn.microsoft.com/library/1c1ea306-7ab0-8818-2fb6-8ac377f73484%28Office.15%29.aspx)|
|[MenuBar](http://msdn.microsoft.com/library/b9e6b6f6-5e60-271d-67c4-6697cb294671%28Office.15%29.aspx)|
|[MinMaxButtons](http://msdn.microsoft.com/library/12f2a0b1-1f45-544b-b116-8d5aa51d6897%28Office.15%29.aspx)|
|[Modal](http://msdn.microsoft.com/library/a36b42f6-9d97-acea-cda3-2f380a3270c2%28Office.15%29.aspx)|
|[Module](http://msdn.microsoft.com/library/f4583bc6-a412-811e-a428-cfa10a911d35%28Office.15%29.aspx)|
|[MouseWheel](http://msdn.microsoft.com/library/364f7854-d7d5-5fe2-effa-6154e86376b4%28Office.15%29.aspx)|
|[Moveable](http://msdn.microsoft.com/library/ad0db2eb-9905-15d9-7a96-e61cefd12842%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/29cd22a8-7f38-9559-8c34-d6319a32adab%28Office.15%29.aspx)|
|[NavigationButtons](http://msdn.microsoft.com/library/23af1adc-67e9-b39d-772b-ddecf159f861%28Office.15%29.aspx)|
|[NavigationCaption](http://msdn.microsoft.com/library/0801ef4c-3f0c-6d45-d1f1-4ed46163586e%28Office.15%29.aspx)|
|[NewRecord](http://msdn.microsoft.com/library/9e30b019-1c1d-31eb-cc8d-cab030861ddc%28Office.15%29.aspx)|
|[OnActivate](http://msdn.microsoft.com/library/ab9899de-e0dc-7884-e293-e031098d644c%28Office.15%29.aspx)|
|[OnApplyFilter](http://msdn.microsoft.com/library/5e147a50-5516-f6d3-c1c9-e2c4522cb804%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/a02c677d-d96b-754f-3ca8-0089a27a7e84%28Office.15%29.aspx)|
|[OnClose](http://msdn.microsoft.com/library/af4a7532-f12a-5194-9636-a09f9221f465%28Office.15%29.aspx)|
|[OnConnect](http://msdn.microsoft.com/library/de181e49-ccba-52fa-f521-3e55f3ed78d2%28Office.15%29.aspx)|
|[OnCurrent](http://msdn.microsoft.com/library/bb7eb7be-7bb6-8fdd-6a48-f5b33ad7dc14%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/8e38c442-f2b2-b125-e006-b507765fefd4%28Office.15%29.aspx)|
|[OnDeactivate](http://msdn.microsoft.com/library/c241c3cc-377b-7407-87f3-3003edb3ff8f%28Office.15%29.aspx)|
|[OnDelete](http://msdn.microsoft.com/library/97cfb9eb-e1c7-a879-a8aa-d26ff337efbb%28Office.15%29.aspx)|
|[OnDirty](http://msdn.microsoft.com/library/e1b14d73-a5f6-a393-ea29-4b98cc7bfdd4%28Office.15%29.aspx)|
|[OnDisconnect](http://msdn.microsoft.com/library/8f6514c7-8f61-2ae7-0859-8299523609ca%28Office.15%29.aspx)|
|[OnError](http://msdn.microsoft.com/library/f89366ad-7d68-cb0f-0b17-c6b4f4eb3f3c%28Office.15%29.aspx)|
|[OnFilter](http://msdn.microsoft.com/library/4d1b52cb-0f79-d8e9-05b3-a7a1da0a7a62%28Office.15%29.aspx)|
|[OnGotFocus](http://msdn.microsoft.com/library/04f2e6e2-aaa3-eb05-16ff-32d5a252df94%28Office.15%29.aspx)|
|[OnInsert](http://msdn.microsoft.com/library/26c0ceb7-f345-2ca8-eb0c-744c60cf5340%28Office.15%29.aspx)|
|[OnKeyDown](http://msdn.microsoft.com/library/73302cbd-25bc-4ae1-8df9-7813d0a67b65%28Office.15%29.aspx)|
|[OnKeyPress](http://msdn.microsoft.com/library/1ba311c2-15f2-1756-b35c-18df7cf7f858%28Office.15%29.aspx)|
|[OnKeyUp](http://msdn.microsoft.com/library/18cc6617-082d-584d-518b-f89e4c71f8eb%28Office.15%29.aspx)|
|[OnLoad](http://msdn.microsoft.com/library/8614f8a8-b5ca-6fa6-46b2-7e88d8a8137d%28Office.15%29.aspx)|
|[OnLostFocus](http://msdn.microsoft.com/library/2bbc34d2-e4e6-7133-ef9e-d112514ace92%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/cb812cbf-8ec3-e4b2-ebf3-882c8b21df7f%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/c0d5b1f8-edbc-d319-6edd-e48768b61f40%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/8aea46bc-b62a-351a-0a02-6afbb1362b2b%28Office.15%29.aspx)|
|[OnOpen](http://msdn.microsoft.com/library/151b9103-a25d-a595-6cab-20b737909fa6%28Office.15%29.aspx)|
|[OnResize](http://msdn.microsoft.com/library/84e6df44-53d2-19c9-e8c5-47681649c7e8%28Office.15%29.aspx)|
|[OnTimer](http://msdn.microsoft.com/library/a7df5020-5163-967b-b59a-0fd8f6fe7a54%28Office.15%29.aspx)|
|[OnUndo](http://msdn.microsoft.com/library/30e36849-e190-3a50-a8ef-cf7aa995607c%28Office.15%29.aspx)|
|[OnUnload](http://msdn.microsoft.com/library/70544311-921c-a610-6fbe-bd3bbef0a6a5%28Office.15%29.aspx)|
|[OpenArgs](http://msdn.microsoft.com/library/f18ed66f-01e0-b8a3-a15b-687e738aafe6%28Office.15%29.aspx)|
|[OrderBy](http://msdn.microsoft.com/library/6ca9c25e-9f16-1f08-1ac3-6f19761f9f55%28Office.15%29.aspx)|
|[OrderByOn](http://msdn.microsoft.com/library/8902a8be-344e-d88f-8ac4-71d94dd0e3f0%28Office.15%29.aspx)|
|[OrderByOnLoad](http://msdn.microsoft.com/library/8acb931e-d0fc-4a17-cd89-1f802af4e4d1%28Office.15%29.aspx)|
|[Orientation](http://msdn.microsoft.com/library/25a13b75-48b7-69bd-4d70-e9aa8a94652e%28Office.15%29.aspx)|
|[Page](http://msdn.microsoft.com/library/0ae576ca-75b2-333e-0303-b2bd1e14e438%28Office.15%29.aspx)|
|[Pages](http://msdn.microsoft.com/library/9494fb79-d080-e2cb-6b55-8194ecd81e9b%28Office.15%29.aspx)|
|[Painting](http://msdn.microsoft.com/library/6fbbd097-8882-b633-bbd6-9dcc0bb31db9%28Office.15%29.aspx)|
|[PaintPalette](http://msdn.microsoft.com/library/161a7bfa-c861-68b9-eaac-05a2d7c24d4a%28Office.15%29.aspx)|
|[PaletteSource](http://msdn.microsoft.com/library/91276931-0aa6-7e54-09eb-1747f036aa7c%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/cd2968d5-1862-b01d-1b96-db2c6a5f2554%28Office.15%29.aspx)|
|[Picture](http://msdn.microsoft.com/library/4a42a443-13f0-e7aa-848d-34faff52c9bd%28Office.15%29.aspx)|
|[PictureAlignment](http://msdn.microsoft.com/library/8e6c09ac-9e2e-14b2-c3cf-09be95cd10b8%28Office.15%29.aspx)|
|[PictureData](http://msdn.microsoft.com/library/09748208-d338-f87d-a53a-4cabee01addb%28Office.15%29.aspx)|
|[PicturePalette](http://msdn.microsoft.com/library/4b7f2c69-37c7-f05a-783d-0b57242253b2%28Office.15%29.aspx)|
|[PictureSizeMode](http://msdn.microsoft.com/library/b2e7646c-a040-0205-b840-0ed5b43982ab%28Office.15%29.aspx)|
|[PictureTiling](http://msdn.microsoft.com/library/9343925c-8184-e9fc-ed62-a272a0bfa0a6%28Office.15%29.aspx)|
|[PictureType](http://msdn.microsoft.com/library/93d3b9e4-ca7d-5f21-81b7-24270532dfa2%28Office.15%29.aspx)|
|[PivotTable](http://msdn.microsoft.com/library/a80edfb5-966b-e1d9-d13e-daefe06c6777%28Office.15%29.aspx)|
|[PivotTableChange](http://msdn.microsoft.com/library/d8d6a7eb-2bc1-e441-95fe-aefaec7fde9d%28Office.15%29.aspx)|
|[PopUp](http://msdn.microsoft.com/library/0ccaa174-80e2-5ca3-9614-93b12dc1bfcd%28Office.15%29.aspx)|
|[Printer](http://msdn.microsoft.com/library/c533271a-c500-57de-f16c-ed384698f829%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/6259b555-293b-2095-eb54-09a2b532e2a3%28Office.15%29.aspx)|
|[PrtDevMode](http://msdn.microsoft.com/library/a20a2dd9-4e5a-6fb7-63ba-8394e654057f%28Office.15%29.aspx)|
|[PrtDevNames](http://msdn.microsoft.com/library/0befbc54-6536-9f51-62da-aa01b5b96961%28Office.15%29.aspx)|
|[PrtMip](http://msdn.microsoft.com/library/0b87f955-638c-5cd2-95b1-5aec870350ff%28Office.15%29.aspx)|
|[Query](http://msdn.microsoft.com/library/fcef59f9-f405-0a05-f986-b29c2b0528de%28Office.15%29.aspx)|
|[RecordLocks](http://msdn.microsoft.com/library/9080f7dd-259e-8b13-9648-3269bc7321d3%28Office.15%29.aspx)|
|[RecordSelectors](http://msdn.microsoft.com/library/7700f0c5-621f-5145-57be-777d53228379%28Office.15%29.aspx)|
|[Recordset](http://msdn.microsoft.com/library/baf6c8c4-b4ac-8618-ecbf-4444ae5e88d4%28Office.15%29.aspx)|
|[RecordsetClone](http://msdn.microsoft.com/library/d73ef798-477d-9c36-6e29-82b22352c60b%28Office.15%29.aspx)|
|[RecordsetType](http://msdn.microsoft.com/library/29690204-1014-961d-a969-25c44ca5fc6e%28Office.15%29.aspx)|
|[RecordSource](http://msdn.microsoft.com/library/a473695a-7645-744d-bf69-760e1f2b9fb1%28Office.15%29.aspx)|
|[RecordSourceQualifier](http://msdn.microsoft.com/library/e4c94bb5-b1e4-bfeb-c5f1-b21ae27762b2%28Office.15%29.aspx)|
|[ResyncCommand](http://msdn.microsoft.com/library/0df53ea9-5771-0ccd-07ef-f33ad1082a61%28Office.15%29.aspx)|
|[RibbonName](http://msdn.microsoft.com/library/e352711e-a43d-2dd2-d6db-2bbec7c99e74%28Office.15%29.aspx)|
|[RowHeight](http://msdn.microsoft.com/library/1575cb30-54ab-d45b-bb64-844f12336eca%28Office.15%29.aspx)|
|[ScrollBars](http://msdn.microsoft.com/library/d35e3e88-10ce-20f8-d4b1-305b27992395%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/df8d00af-3e1e-86f8-17f4-dd5792193d03%28Office.15%29.aspx)|
|[SelectionChange](http://msdn.microsoft.com/library/e31876fc-103a-d231-a6fa-7cb026a343e1%28Office.15%29.aspx)|
|[SelHeight](http://msdn.microsoft.com/library/c8808132-ab4d-77f1-fbf7-121d37e188fe%28Office.15%29.aspx)|
|[SelLeft](http://msdn.microsoft.com/library/ddc05c0a-3132-5380-33c9-96fa2f92571d%28Office.15%29.aspx)|
|[SelTop](http://msdn.microsoft.com/library/5503187c-09ea-222e-5db2-f3c2298f34dc%28Office.15%29.aspx)|
|[SelWidth](http://msdn.microsoft.com/library/a5ce22e3-af69-209c-f988-16cf4f77fd62%28Office.15%29.aspx)|
|[ServerFilter](http://msdn.microsoft.com/library/18385de5-bc0d-9d2c-f97c-5b42e3689b45%28Office.15%29.aspx)|
|[ServerFilterByForm](http://msdn.microsoft.com/library/f9f8f28e-b67e-1f4e-a70b-c66169fca250%28Office.15%29.aspx)|
|[ShortcutMenu](http://msdn.microsoft.com/library/ec652f43-4dc8-4970-19ad-d117c3193528%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/b45a1673-292e-8ae0-8936-7d3f7b052d1f%28Office.15%29.aspx)|
|[SplitFormDatasheet](http://msdn.microsoft.com/library/765eacb5-ef79-3b1d-6511-501ec0def22c%28Office.15%29.aspx)|
|[SplitFormOrientation](http://msdn.microsoft.com/library/ddf2228f-6973-ae6c-1477-41a07557b7a2%28Office.15%29.aspx)|
|[SplitFormPrinting](http://msdn.microsoft.com/library/0542af4f-c778-9038-0058-74aa187d0fc7%28Office.15%29.aspx)|
|[SplitFormSize](http://msdn.microsoft.com/library/2fb63076-aebe-23ef-2a11-1c7b1b82ccb1%28Office.15%29.aspx)|
|[SplitFormSplitterBar](http://msdn.microsoft.com/library/80b7c812-2382-ea12-9aff-fb83e5baa7ea%28Office.15%29.aspx)|
|[SplitFormSplitterBarSave](http://msdn.microsoft.com/library/70bd37de-9b8c-0e47-80a8-83e53290e04c%28Office.15%29.aspx)|
|[SubdatasheetExpanded](http://msdn.microsoft.com/library/543f2398-ca70-5261-0f9f-e1d864c442e0%28Office.15%29.aspx)|
|[SubdatasheetHeight](http://msdn.microsoft.com/library/0db2e4b5-e64b-6f55-ebfa-bcce98734491%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/7fec664b-b82e-8cd1-93ff-5162c72fb036%28Office.15%29.aspx)|
|[TimerInterval](http://msdn.microsoft.com/library/ee56bcf8-20cb-9d86-ed17-3b85ac88f6f1%28Office.15%29.aspx)|
|[Toolbar](http://msdn.microsoft.com/library/a004200c-5404-c3ba-f00d-591c0f0a545d%28Office.15%29.aspx)|
|[UniqueTable](http://msdn.microsoft.com/library/25f543fd-d636-db47-ef83-97f4409e74c2%28Office.15%29.aspx)|
|[UseDefaultPrinter](http://msdn.microsoft.com/library/bdb7f428-ee00-5a76-e723-6d1858a6172c%28Office.15%29.aspx)|
|[VerticalDatasheetGridlineStyle](http://msdn.microsoft.com/library/b0174311-f03b-aa6a-b15a-697f6be1b2ac%28Office.15%29.aspx)|
|[ViewChange](http://msdn.microsoft.com/library/f8a8fe82-6983-5632-b779-879faf228ac2%28Office.15%29.aspx)|
|[ViewsAllowed](http://msdn.microsoft.com/library/2aa001e0-ea0d-4ef3-f8d2-fdd301502c96%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/85567666-297a-3380-2d08-864d44b637a1%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/3f417a95-16a7-fdfa-8661-e71992c757cc%28Office.15%29.aspx)|
|[WindowHeight](http://msdn.microsoft.com/library/53af4131-a174-c0c3-db29-f0f0cabb4b05%28Office.15%29.aspx)|
|[WindowLeft](http://msdn.microsoft.com/library/f9e90b5e-6008-675d-9168-6dd932559b6d%28Office.15%29.aspx)|
|[WindowTop](http://msdn.microsoft.com/library/1257fe21-3983-bd51-4683-e0778b59a975%28Office.15%29.aspx)|
|[WindowWidth](http://msdn.microsoft.com/library/81839600-01e6-0462-3cf0-48de708e3d64%28Office.15%29.aspx)|

## About the Contributors
<a name="AboutContributors"> </a>

Luke Chung is the founder and president of FMS, Inc., a leading provider of custom database solutions and developer tools. 

UtterAccess is the premier Microsoft Access wiki and help forum.  

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 


## See also
[Access Object Model Reference](object-model-access-vba-reference.md)


