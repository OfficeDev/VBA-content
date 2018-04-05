---
title: Report Object (Access)
keywords: vbaac10.chm13901
f1_keywords:
- vbaac10.chm13901
ms.prod: access
api_name:
- Access.Report
ms.assetid: 6f77c1b4-a9ce-7caa-204c-fe0755c6f9df
ms.date: 11/30/2017
---


# Report Object (Access)

A **Report** object refers to a particular Microsoft Access report.


## Remarks

A **Report** object is a member of the **Reports** collection, which is a collection of all currently open reports. Within the **Reports** collection, individual reports are indexed beginning with zero. You can refer to an individual **Report** object in the **Reports** collection either by referring to the report by name, or by referring to its index within the collection. If the report name includes a space, the name must be surrounded by brackets ([ ]).



|**Syntax**|**Example**|
|:-----|:-----|
|**Reports** ! _reportname_|Reports!OrderReport|
|**Reports** ![ _report name_]|Reports![Order Report]|
|**Reports** (" _reportname_")|Reports("OrderReport")|
|**Reports** ( _index_)|Reports(0)|

> [!NOTE]
> Each **Report** object has a **Controls** collection, which contains all controls on the report. You can refer to a control on a report either by implicitly or explicitly referring to the **Controls** collection. Your code will be faster if you refer to the **Controls** collection implicitly. The following examples show two of the ways you might refer to a control named **NewData** on a report called **OrderReport**. 


```
' Implicit reference. 
Reports!OrderReport!NewData
```

<br/>

```
' Explicit reference. 
Reports!OrderReport.Controls!NewData
```


## Example


The following example shows how to use the **NoData** event of a report to prevent the report form opening when there is no data to be displayed.

**Sample code provided by:** The [Microsoft Access 2010 Programmer's Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.html)

```
Private Sub Report_NoData(Cancel As Integer)

    'Add code here that will be executed if no data
    'was returned by the Report's RecordSource
    MsgBox "No customers ordered this product this month. " &amp; _
        "The report will now close."
    Cancel = True

End Sub
```

<br/>

The following example shows how to use the **Page** event to add a watermark to a report before it is printed.

```
Private Sub Report_Page()
    Dim strWatermarkText As String
    Dim sizeHor As Single
    Dim sizeVer As Single

#If RUN_PAGE_EVENT = True Then
    With Me
        '// Print page border
        Me.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), vbBlack, B
    
        '// Print watermark
        strWatermarkText = "Confidential"
        
        .ScaleMode = 3
        .FontName = "Segoe UI"
        .FontSize = 48
        .ForeColor = RGB(255, 0, 0)

        '// Calculate text metrics
        sizeHor = .TextWidth(strWatermarkText)
        sizeVer = .TextHeight(strWatermarkText)
        
        '// Set the print location
        .CurrentX = (.ScaleWidth / 2) - (sizeHor / 2)
        .CurrentY = (.ScaleHeight / 2) - (sizeVer / 2)
    
        '// Print the watermark
        .Print strWatermarkText
    End With
#End If

End Sub
```

<br/>

The following example shows how to set the **BackColor** property of a control based on its value.

```
Private Sub SetControlFormatting()
    If (Me.AvgOfRating >= 8) Then
        Me.AvgOfRating.BackColor = vbGreen
    ElseIf (Me.AvgOfRating >= 5) Then
        Me.AvgOfRating.BackColor = vbYellow
    Else
        Me.AvgOfRating.BackColor = vbRed
    End If
End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    ' size the width of the rectangle
    Dim lngOffset As Long
    lngOffset = (Me.boxInside.Left - Me.boxOutside.Left) * 2
    Me.boxInside.Width = (Me.boxOutside.Width * (Me.AvgOfRating / 10)) - lngOffset
    
    ' do conditional formatting for the control in print preview
    SetControlFormatting
End Sub

Private Sub Detail_Paint()
    ' do conditional formatting for the control in report view
    SetControlFormatting
End Sub
```

<br/>

The following example shows how to format a report to show progress bars. The example uses a pair of rectangle controls, **boxInside** and **boxOutside**, to create a progress bar based on the value of **AvgOfRating**. The progress bars are visible only when the report is opened in **Print Preview** mode or it is printed.

```
Private Sub Report_Load()
    If (Me.CurrentView = AcCurrentView.acCurViewPreview) Then
        Me.boxInside.Visible = True
        Me.boxOutside.Visible = True
    Else
        Me.boxInside.Visible = False
        Me.boxOutside.Visible = False
    End If
End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    ' size the width of the rectangle
    Dim lngOffset As Long
    lngOffset = (Me.boxInside.Left - Me.boxOutside.Left) * 2
    Me.boxInside.Width = (Me.boxOutside.Width * (Me.AvgOfRating / 10)) - lngOffset
    
    ' do conditional formatting for the control in print preview
    SetControlFormatting
End Sub
```


## Events

|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/565cf35c-e7ea-e1ec-b23b-b84a6318fde7%28Office.15%29.aspx)|
|[ApplyFilter](http://msdn.microsoft.com/library/46cbe83d-4395-d9e6-3187-c51152269e62%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/37bd4936-2f66-b434-ae54-5f76dd943c4c%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/529a238b-087c-f70e-e651-2630ef1d427a%28Office.15%29.aspx)|
|[Current](http://msdn.microsoft.com/library/adfdbda0-c3e9-c3c6-8768-415b4bd270d5%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/e4cd6226-8647-1a94-07a4-00ecef1ccde7%28Office.15%29.aspx)|
|[Deactivate](http://msdn.microsoft.com/library/15e96e8a-c5f5-1a68-56cc-0ceaa1dbd407%28Office.15%29.aspx)|
|[Error](http://msdn.microsoft.com/library/06d88711-df19-6453-a7ce-095d3d02674f%28Office.15%29.aspx)|
|[Filter](http://msdn.microsoft.com/library/1344ceff-d3ac-3dc1-0f9c-563d895a77dc%28Office.15%29.aspx)|
|[GotFocus](http://msdn.microsoft.com/library/667b4798-4407-f60f-af3a-7788a0501761%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/b33ecbca-b3a1-19b2-8541-fe4bcbf4acec%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/0c846367-a4b0-d716-dcc3-32c916e09dfb%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/5561cbab-b6bd-ab4e-83a6-fbf7ec9272d1%28Office.15%29.aspx)|
|[Load](http://msdn.microsoft.com/library/966527a0-4c61-9f5e-50ca-791d39bd24ac%28Office.15%29.aspx)|
|[LostFocus](http://msdn.microsoft.com/library/8b80c2bc-8be4-1842-4011-0e6475b3a865%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/fcce0466-9c65-8e76-eb2a-e0a82d299015%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/b7df8ba7-dd10-4aea-1b79-df33e151250d%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/e7b6aa74-1cba-ee10-03d1-11236d14faae%28Office.15%29.aspx)|
|[MouseWheel](http://msdn.microsoft.com/library/9c234923-3459-c45e-8489-146353f59c21%28Office.15%29.aspx)|
|[NoData](http://msdn.microsoft.com/library/fa5f22b1-3695-bd16-2ca3-b2a1cc1f1d94%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/d170b67d-3123-6f51-6cf8-38433736f104%28Office.15%29.aspx)|
|[Page](http://msdn.microsoft.com/library/c3fcce28-0bcd-4ef1-427f-504f0f80d336%28Office.15%29.aspx)|
|[Resize](http://msdn.microsoft.com/library/cd2c1c2a-959b-a5d0-9f75-a7443a9a57f1%28Office.15%29.aspx)|
|[Timer](http://msdn.microsoft.com/library/52e3db7f-a61c-8144-e39b-0f9daf61bd98%28Office.15%29.aspx)|
|[Unload](http://msdn.microsoft.com/library/05f0d51e-8fa0-9547-6b22-e7711754d1a5%28Office.15%29.aspx)|

## Methods

|**Name**|
|:-----|
|[Circle](http://msdn.microsoft.com/library/4f5d24e2-75bf-3586-7e0d-0902adee61a6%28Office.15%29.aspx)|
|[Line](http://msdn.microsoft.com/library/9e640e37-c055-3dc3-b70e-0805cdc13561%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/169ac85a-394f-5db2-7b55-b6ca5fd03546%28Office.15%29.aspx)|
|[Print](http://msdn.microsoft.com/library/6f8523cc-7b17-ec27-e2c9-a7ae3d5a8c3f%28Office.15%29.aspx)|
|[PSet](http://msdn.microsoft.com/library/951a262b-b17b-9b95-b5f2-922d4aff9ce9%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/d078d523-3bbd-fa70-44ac-116cdcedfebd%28Office.15%29.aspx)|
|[Scale](http://msdn.microsoft.com/library/6a261d1d-9474-7374-f399-4d46e404058b%28Office.15%29.aspx)|
|[TextHeight](http://msdn.microsoft.com/library/cac67d4c-e140-06ae-ccbd-961cdee3d087%28Office.15%29.aspx)|
|[TextWidth](http://msdn.microsoft.com/library/98827373-8610-5e48-ab46-2c89f8e2d2a7%28Office.15%29.aspx)|

## Properties

|**Name**|
|:-----|
|[ActiveControl](http://msdn.microsoft.com/library/71599d10-a423-ab75-e12b-03adec04c2bf%28Office.15%29.aspx)|
|[AllowLayoutView](http://msdn.microsoft.com/library/5388fcd8-32fb-781d-538c-ac114f8d5bd8%28Office.15%29.aspx)|
|[AllowReportView](http://msdn.microsoft.com/library/43db97fa-bdc0-883c-7b83-a7bbe7c62c07%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/2fb2ca56-bbc9-e689-08f3-d42fa97f21d2%28Office.15%29.aspx)|
|[AutoCenter](http://msdn.microsoft.com/library/d4a12dac-1000-38cd-e4ed-4f5879dfe4a0%28Office.15%29.aspx)|
|[AutoResize](http://msdn.microsoft.com/library/bf18b1b2-aba6-d4fe-7916-de821c76fbb4%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/0f3f0ac9-5a25-13fb-0227-f0f6384d647b%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/8e662558-755f-0dbe-8803-b0f0ef093172%28Office.15%29.aspx)|
|[CloseButton](http://msdn.microsoft.com/library/dad15f66-4787-a4eb-dbbe-d698faaa0917%28Office.15%29.aspx)|
|[ControlBox](http://msdn.microsoft.com/library/440dd25d-4792-2a92-beac-21dbcf478b62%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/ea1ad090-91ba-d2c8-2a42-83227068548f%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/72848003-cdf4-586b-c059-6c821104fbda%28Office.15%29.aspx)|
|[CurrentRecord](http://msdn.microsoft.com/library/da19aa9e-6192-1e51-3c07-aadd2d8ebe4c%28Office.15%29.aspx)|
|[CurrentView](http://msdn.microsoft.com/library/d1c33390-75f1-4e11-0735-a8860211b4ce%28Office.15%29.aspx)|
|[CurrentX](http://msdn.microsoft.com/library/3b5e7c50-ecb4-606a-6715-4d54ed912c45%28Office.15%29.aspx)|
|[CurrentY](http://msdn.microsoft.com/library/040c0b5d-f7d6-2fa1-e34d-f69799f0b273%28Office.15%29.aspx)|
|[Cycle](http://msdn.microsoft.com/library/031194ca-f058-3a73-3551-f67d4e9bc27a%28Office.15%29.aspx)|
|[DateGrouping](http://msdn.microsoft.com/library/e2495aa7-06e9-8eaf-81d8-182c7d51559c%28Office.15%29.aspx)|
|[DefaultControl](http://msdn.microsoft.com/library/13c06cbc-b6bb-60dc-dc84-d16abdeffe9c%28Office.15%29.aspx)|
|[DefaultView](http://msdn.microsoft.com/library/75eb8fcd-9e28-bda4-d560-a2a5bfca0450%28Office.15%29.aspx)|
|[Dirty](http://msdn.microsoft.com/library/d9d9fe57-4fc5-9700-fc14-71f9eaa4a01b%28Office.15%29.aspx)|
|[DisplayOnSharePointSite](http://msdn.microsoft.com/library/4e13b1e9-3b79-d073-fb51-848fdc2dcada%28Office.15%29.aspx)|
|[DrawMode](http://msdn.microsoft.com/library/773a3c7f-fb59-9614-3363-b417607fbe28%28Office.15%29.aspx)|
|[DrawStyle](http://msdn.microsoft.com/library/0dd2afb9-d310-3637-6ed7-e66c9ad3460d%28Office.15%29.aspx)|
|[DrawWidth](http://msdn.microsoft.com/library/1bda5387-9244-f150-2165-8dba1684ca25%28Office.15%29.aspx)|
|[FastLaserPrinting](http://msdn.microsoft.com/library/b96ec618-de46-8802-0d9e-064fd8835fbd%28Office.15%29.aspx)|
|[FillColor](http://msdn.microsoft.com/library/04fa1376-fddb-a4b3-04fd-d562f0567136%28Office.15%29.aspx)|
|[FillStyle](http://msdn.microsoft.com/library/0fcb840d-4ff6-718a-2267-25cd2622c8d2%28Office.15%29.aspx)|
|[Filter](http://msdn.microsoft.com/library/ce57e05d-c3a8-012a-205e-7dfb2e4dd78d%28Office.15%29.aspx)|
|[FilterOn](http://msdn.microsoft.com/library/94751217-8b8b-0979-b4f8-c9af9f38ae69%28Office.15%29.aspx)|
|[FilterOnLoad](http://msdn.microsoft.com/library/37d0e39d-dfd5-c2b7-e549-9b165a90ceb9%28Office.15%29.aspx)|
|[FitToPage](http://msdn.microsoft.com/library/e2210e28-273b-8eb5-0229-5f6513cf5ae2%28Office.15%29.aspx)|
|[FontBold](http://msdn.microsoft.com/library/0a3589d9-96a3-0a48-03a5-4e08f9da2c74%28Office.15%29.aspx)|
|[FontItalic](http://msdn.microsoft.com/library/e6cc9478-2bbd-6a80-daff-95e160bdcbe6%28Office.15%29.aspx)|
|[FontName](http://msdn.microsoft.com/library/37759316-e5f6-14f6-0423-c5a11e02161f%28Office.15%29.aspx)|
|[FontSize](http://msdn.microsoft.com/library/7fbb96dd-9354-39e3-a62a-0ca0e3532126%28Office.15%29.aspx)|
|[FontUnderline](http://msdn.microsoft.com/library/37f62220-069d-939d-7ad0-e9f25ae6bf36%28Office.15%29.aspx)|
|[ForeColor](http://msdn.microsoft.com/library/048b01a3-f962-d2d3-b546-027fec6a1369%28Office.15%29.aspx)|
|[FormatCount](http://msdn.microsoft.com/library/35fbc0fb-a106-11d6-26db-99d6f0b969a3%28Office.15%29.aspx)|
|[GridX](http://msdn.microsoft.com/library/b932531f-89d3-5f8e-d6cc-43baf1004149%28Office.15%29.aspx)|
|[GridY](http://msdn.microsoft.com/library/e4a13708-fa05-8ac4-af5f-0f78ee15e623%28Office.15%29.aspx)|
|[GroupLevel](http://msdn.microsoft.com/library/8a40502d-84ac-0652-8c07-c4c155ec1242%28Office.15%29.aspx)|
|[GrpKeepTogether](http://msdn.microsoft.com/library/605e8999-d184-b8d9-3f55-9926cd0ceefd%28Office.15%29.aspx)|
|[HasData](http://msdn.microsoft.com/library/e8827477-6877-ec7a-63e5-7f4de972f0bb%28Office.15%29.aspx)|
|[HasModule](http://msdn.microsoft.com/library/a4f33211-aaa8-d082-feed-aea75bda8659%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/14821735-efbb-e831-e1d4-94f34de41ef7%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/3911ba15-a1fd-06a6-659f-b8599bb01931%28Office.15%29.aspx)|
|[HelpFile](http://msdn.microsoft.com/library/cfab3009-c8e1-5f56-020a-e0a972e0de50%28Office.15%29.aspx)|
|[Hwnd](http://msdn.microsoft.com/library/e2d045f4-57bf-8681-0e00-bb5fe287136d%28Office.15%29.aspx)|
|[InputParameters](http://msdn.microsoft.com/library/c544db38-9d31-42ff-3fb7-98a79d9d2fc2%28Office.15%29.aspx)|
|[KeyPreview](http://msdn.microsoft.com/library/49ca195d-bd9e-7a69-1891-455581bcf09a%28Office.15%29.aspx)|
|[LayoutForPrint](http://msdn.microsoft.com/library/f661155f-696b-3acf-5b90-44fba06345ab%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/678601b5-ab80-2c19-9a29-7c5c2d63f792%28Office.15%29.aspx)|
|[MenuBar](http://msdn.microsoft.com/library/008e1d2e-f467-05a4-d246-eba85fd626ba%28Office.15%29.aspx)|
|[MinMaxButtons](http://msdn.microsoft.com/library/8aee0247-804a-e9ee-e11a-11c9c5d37ed6%28Office.15%29.aspx)|
|[Modal](http://msdn.microsoft.com/library/654ff830-c8d9-5bd9-1ec6-61ee6546b4db%28Office.15%29.aspx)|
|[Module](http://msdn.microsoft.com/library/e0cff3db-1697-7b8e-3934-7ead204052fb%28Office.15%29.aspx)|
|[MouseWheel](http://msdn.microsoft.com/library/ea9d6443-abfd-6140-e167-548f4aafd342%28Office.15%29.aspx)|
|[Moveable](http://msdn.microsoft.com/library/77e682a5-7a0f-f55e-a469-2770bb2de844%28Office.15%29.aspx)|
|[MoveLayout](http://msdn.microsoft.com/library/b02ddbda-ea3f-aad7-5f92-3b308dac4e79%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/a5d01749-e127-8807-4c90-a86c2d5dc028%28Office.15%29.aspx)|
|[NextRecord](http://msdn.microsoft.com/library/771508ff-9a2d-6317-2b23-a1c0b012e7ba%28Office.15%29.aspx)|
|[OnActivate](http://msdn.microsoft.com/library/eb7f05e3-edba-ab9e-3708-5c3ee7b2ee18%28Office.15%29.aspx)|
|[OnApplyFilter](http://msdn.microsoft.com/library/18e5b016-19a0-46bb-c552-c4bb8d458ca4%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/45161619-ed2c-ff3d-00a6-26ed802e0480%28Office.15%29.aspx)|
|[OnClose](http://msdn.microsoft.com/library/640b5540-4b0d-6649-0a36-9dd63a437c84%28Office.15%29.aspx)|
|[OnCurrent](http://msdn.microsoft.com/library/593fdb6c-017a-986f-22ef-cc9e66aaaf01%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/b92a1b2b-4f27-4f45-959c-6f1aec557004%28Office.15%29.aspx)|
|[OnDeactivate](http://msdn.microsoft.com/library/2b15bb7c-a307-6e2b-c933-b7a069ff99d0%28Office.15%29.aspx)|
|[OnError](http://msdn.microsoft.com/library/28436e0e-a37e-8acd-6c3c-1f6d96c63e23%28Office.15%29.aspx)|
|[OnFilter](http://msdn.microsoft.com/library/72af402e-8e37-328e-b0f4-89f54f59bce0%28Office.15%29.aspx)|
|[OnGotFocus](http://msdn.microsoft.com/library/259d14b1-cd39-722e-b4d7-28742fefd831%28Office.15%29.aspx)|
|[OnKeyDown](http://msdn.microsoft.com/library/22be1d11-abbd-81ff-d83c-66aa2884560a%28Office.15%29.aspx)|
|[OnKeyPress](http://msdn.microsoft.com/library/9f6dcc2e-b2b1-56bc-2c3a-c7be498eda72%28Office.15%29.aspx)|
|[OnKeyUp](http://msdn.microsoft.com/library/a31122bb-3f5a-4021-a2b5-16487aa0ce7c%28Office.15%29.aspx)|
|[OnLoad](http://msdn.microsoft.com/library/b9ce7eaf-3f52-4cdf-a8eb-74f242c6b526%28Office.15%29.aspx)|
|[OnLostFocus](http://msdn.microsoft.com/library/5a0e6b1d-ad2b-f28e-a565-dddeff9659c6%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/8b66aa47-d409-4cc6-2441-6c959f7120a4%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/260c8b56-2985-1da4-7c3f-1398b54666b3%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/0fc68ad5-6738-ba57-2f31-40d3d3e93cb3%28Office.15%29.aspx)|
|[OnNoData](http://msdn.microsoft.com/library/5d3cfec5-1b57-625c-c350-0d7e475be2d2%28Office.15%29.aspx)|
|[OnOpen](http://msdn.microsoft.com/library/e381f9a5-c409-7ae5-e266-cb3a046eb919%28Office.15%29.aspx)|
|[OnPage](http://msdn.microsoft.com/library/d72bab5d-fdb8-99f5-5d27-8227bc0136ec%28Office.15%29.aspx)|
|[OnResize](http://msdn.microsoft.com/library/336eceb4-7f78-b0b0-cb8f-a6a35c8bea76%28Office.15%29.aspx)|
|[OnTimer](http://msdn.microsoft.com/library/ef7ac956-ffa4-da79-0d39-9c505409b4af%28Office.15%29.aspx)|
|[OnUnload](http://msdn.microsoft.com/library/0ebc34b7-3541-4d35-fc9b-ac0feb41b873%28Office.15%29.aspx)|
|[OpenArgs](http://msdn.microsoft.com/library/91dcbf42-6bb8-73e5-744c-de82d8668f9c%28Office.15%29.aspx)|
|[OrderBy](http://msdn.microsoft.com/library/1939157c-12ad-2e58-bf4c-22c04a6c4366%28Office.15%29.aspx)|
|[OrderByOn](http://msdn.microsoft.com/library/8784e57f-e4f1-a606-36b0-1200d6f17b89%28Office.15%29.aspx)|
|[OrderByOnLoad](http://msdn.microsoft.com/library/28c05775-7090-a699-c7be-8a17b43210b0%28Office.15%29.aspx)|
|[Orientation](http://msdn.microsoft.com/library/46687f4f-77e2-d9c3-ed12-5df0a8abc2bb%28Office.15%29.aspx)|
|[Page](http://msdn.microsoft.com/library/6d1dd330-ecd8-3b5c-c851-26bf7e431f98%28Office.15%29.aspx)|
|[PageFooter](http://msdn.microsoft.com/library/82cd1c0f-2823-9b61-a1fd-66c02c6aaadf%28Office.15%29.aspx)|
|[PageHeader](http://msdn.microsoft.com/library/9f9fe114-b5a5-39c7-d2c0-39453948ace6%28Office.15%29.aspx)|
|[Pages](http://msdn.microsoft.com/library/b97a6878-0a2c-3834-8f3d-6f4460dab3bd%28Office.15%29.aspx)|
|[Painting](http://msdn.microsoft.com/library/82c5a5e6-9d87-7293-e0f5-8ee950f3b85f%28Office.15%29.aspx)|
|[PaintPalette](http://msdn.microsoft.com/library/d4c05c71-52da-6185-89d6-a69c7a883e0a%28Office.15%29.aspx)|
|[PaletteSource](http://msdn.microsoft.com/library/9dc324a1-dc31-b0c5-edca-c4bc1674155a%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/8ad25142-21e4-f0ae-d1c6-621dee5edc69%28Office.15%29.aspx)|
|[Picture](http://msdn.microsoft.com/library/18c914c4-0c6d-6ab3-49e0-0e68a9b60ce0%28Office.15%29.aspx)|
|[PictureAlignment](http://msdn.microsoft.com/library/d038e65b-c258-b6b7-ce53-87b9a60e74e3%28Office.15%29.aspx)|
|[PictureData](http://msdn.microsoft.com/library/b9100f5e-5734-ca30-1cbf-45f8afaadd75%28Office.15%29.aspx)|
|[PicturePages](http://msdn.microsoft.com/library/a1266a43-3e1c-33f3-ae18-a7306723cc11%28Office.15%29.aspx)|
|[PicturePalette](http://msdn.microsoft.com/library/55f8363a-de60-c92f-6330-2cd9f6a16785%28Office.15%29.aspx)|
|[PictureSizeMode](http://msdn.microsoft.com/library/7343ec48-b15e-632e-7493-776d8c9cd456%28Office.15%29.aspx)|
|[PictureTiling](http://msdn.microsoft.com/library/44927121-1ec4-1edf-b3ca-3e00022fab08%28Office.15%29.aspx)|
|[PictureType](http://msdn.microsoft.com/library/96a8ab1c-42d2-2322-927f-4b2cf8822c56%28Office.15%29.aspx)|
|[PopUp](http://msdn.microsoft.com/library/76e82181-c5d5-01b2-c7ce-b2c78f237a75%28Office.15%29.aspx)|
|[PrintCount](http://msdn.microsoft.com/library/9228d6eb-872c-db58-b316-78bff8b375dc%28Office.15%29.aspx)|
|[Printer](http://msdn.microsoft.com/library/9e21b583-5539-bc24-49a0-c248e7f9aafb%28Office.15%29.aspx)|
|[PrintSection](http://msdn.microsoft.com/library/745f4624-557b-0a4c-d4f4-9f0ba4113a61%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/0711a5a9-7e41-66c9-f5a1-fe16fb6842c5%28Office.15%29.aspx)|
|[PrtDevMode](http://msdn.microsoft.com/library/a0c74bb7-7c9d-d978-c9de-de893e935899%28Office.15%29.aspx)|
|[PrtDevNames](http://msdn.microsoft.com/library/96a3437b-3655-5a87-9a1f-722116c82708%28Office.15%29.aspx)|
|[PrtMip](http://msdn.microsoft.com/library/f2a3eb10-04d5-c1fc-5ca3-0dc588db18ff%28Office.15%29.aspx)|
|[RecordLocks](http://msdn.microsoft.com/library/21f8d145-e417-a7a1-e697-b1e07434c760%28Office.15%29.aspx)|
|[Recordset](http://msdn.microsoft.com/library/8f37dfcd-ee53-c3f1-0edc-b3c38f263686%28Office.15%29.aspx)|
|[RecordSource](http://msdn.microsoft.com/library/aa3b31cc-21a6-5d56-8361-9fc232ffae97%28Office.15%29.aspx)|
|[RecordSourceQualifier](http://msdn.microsoft.com/library/8ebf77b6-69c8-e386-2bd5-687e46a872fb%28Office.15%29.aspx)|
|[Report](http://msdn.microsoft.com/library/0cacc875-2083-159a-423f-757ab19e5839%28Office.15%29.aspx)|
|[RibbonName](http://msdn.microsoft.com/library/598dc161-1d90-8339-a214-95d6e9d6396a%28Office.15%29.aspx)|
|[ScaleHeight](http://msdn.microsoft.com/library/b150ece7-b285-669f-8677-f28d6899454b%28Office.15%29.aspx)|
|[ScaleLeft](http://msdn.microsoft.com/library/1e20b9ca-5b5b-2b05-431e-1957f5c70524%28Office.15%29.aspx)|
|[ScaleMode](http://msdn.microsoft.com/library/e3955e48-80bb-989e-2992-cd5a541b468b%28Office.15%29.aspx)|
|[ScaleTop](http://msdn.microsoft.com/library/2f148587-6da0-a6d3-414a-82f97d94a615%28Office.15%29.aspx)|
|[ScaleWidth](http://msdn.microsoft.com/library/b6bdab85-d0d0-99d1-af59-b0b0fe48ab1e%28Office.15%29.aspx)|
|[ScrollBars](http://msdn.microsoft.com/library/12693642-6288-4f21-40cd-5aa1d6886cca%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/3baad974-8869-30b5-abe3-8cf754a225b3%28Office.15%29.aspx)|
|[ServerFilter](http://msdn.microsoft.com/library/e73ad797-8c76-705f-080b-2d0f3423cb39%28Office.15%29.aspx)|
|[Shape](http://msdn.microsoft.com/library/1fd2eb3c-5473-b239-d0c6-4e0ded950df6%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/2773dcb2-b122-6502-66f6-ecd412fb75d0%28Office.15%29.aspx)|
|[ShowPageMargins](http://msdn.microsoft.com/library/7001d6ae-40db-ca7b-5276-0f299890ff9f%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/7e67170b-0058-bdd8-161b-806f732fbca4%28Office.15%29.aspx)|
|[TimerInterval](http://msdn.microsoft.com/library/272fb1f6-2aca-60c2-1f0f-d901e0da91ac%28Office.15%29.aspx)|
|[Toolbar](http://msdn.microsoft.com/library/e897d294-2d8d-aca7-9aed-4bd2ebd23552%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/badaa1a0-44ef-c2cd-64fa-8450add21d69%28Office.15%29.aspx)|
|[UseDefaultPrinter](http://msdn.microsoft.com/library/a7edf38e-181b-3822-bdb4-fb74ec18d40a%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/b860f01b-3a3e-14ab-686b-402fef0027f9%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/f6631a47-79a0-3b8e-e0f0-11aab5b1b477%28Office.15%29.aspx)|
|[WindowHeight](http://msdn.microsoft.com/library/316bce4d-1b5b-e30e-5e80-a4bc31c41d11%28Office.15%29.aspx)|
|[WindowLeft](http://msdn.microsoft.com/library/839ca3d7-4d53-c9e8-b47f-34f94eb5083f%28Office.15%29.aspx)|
|[WindowTop](http://msdn.microsoft.com/library/99d1bec5-f6ac-bf5b-39d0-869a565e0572%28Office.15%29.aspx)|
|[WindowWidth](http://msdn.microsoft.com/library/55d2354d-1a7a-2432-f9ab-bef3f1920aa4%28Office.15%29.aspx)|

## About the contributors
<a name="AboutContributors"> </a>

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 


## See also

[Report Object Members](http://msdn.microsoft.com/library/73370a33-1ca0-da4d-9e36-88011bc2b93e%28Office.15%29.aspx)
[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
