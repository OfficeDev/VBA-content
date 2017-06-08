---
title: Attachment Object (Access)
keywords: vbaac10.chm14036
f1_keywords:
- vbaac10.chm14036
ms.prod: access
api_name:
- Access.Attachment
ms.assetid: b0756145-9012-f9b9-7df9-e168defed3bf
ms.date: 06/08/2017
---


# Attachment Object (Access)

This object corresponds to an attachment control. Use an attachment control when you want to manipulate the contents fields of the attachment data type.


## Remarks


 **Note**  You can attach files only to databases that you create in Office Access 2007 and that use the new .accdb file format. You cannot share attachments between a Office Access 2007 (.accdb) database and a database in the earlier (.mdb) file format.

You can attach a maximum of two gigabytes of data (the maximum size for an Access database). Individual files cannot exceed 256 megabytes in size.


### Supported image file formats

Office Access 2007 supports the following graphic file formats natively, meaning the attachment control renders them without the need for additional software.


- BMP (Windows Bitmap)
    
- RLE (Run Length Encoded Bitmap)
    
- DIB (Device Independent Bitmap)
    
- GIF (Graphics Interchange Format)
    
- JPEG, JPG, JPE (Joint Photographic Experts Group)
    
- EXIF (Exchangeable File Format)
    
- PNG (Portable Network Graphics)
    
- TIFF, TIF (Tagged Image File Format)
    
- ICON, ICO (Icon)
    
- WMF (Windows Metafile)
    
- EMF (Enhanced Metafile)
    

### Supported formats for documents and other files

As a rule, you can attach any file that was created with one of the 2007 Microsoft Office system programs. You can also attach log files (.log), text files (.text, .txt), and compressed .zip files.


### File-naming conventions

The names of your attached files can contain any Unicode character supported by the NTFS file system used in Microsoft Windows NT (NTFS). In addition, file names must conform to these guidelines:


- Names must not exceed 255 characters, including the file name extensions.
    
- Names cannot contain the following characters: question marks (?), quotation marks ("), forward or backward slashes (/ \), opening or closing brackets (< >), asterisks (*), vertical bars or pipes (|), colons (:), or paragraph marks (?).
    

### Types of files that Access compresses

Access will compress your attached files unless those files are compressed natively. For example, JPEG files are compressed by the graphics program that created them, so Access does not compress them. the following table lists some supported file types and whether or not Access compresses them.



|**File Extension**|**Compressed?**|**Reason**|
|:-----|:-----|:-----|
|.jpg, .jpeg|No|Already compressed|
|.gif|No|Already compressed|
|.png|No|Already compressed|
|.tif, .tiff|Yes||
|.exif|Yes||
| .bmp|Yes||
|.emf|Yes||
|.wmf|Yes||
|.ico|Yes||
|.zip|No|Already compressed|
|.cab|No|Already compressed|
|.docx|No|Already compressed|
|.xlsx|No|Already compressed|
|.xlsb|No|Already compressed|
|.pptx|No|Already compressed|

### Blocked file formats

Office Access 2007 blocks the following types of attached files. At this time, you cannot unblock any of the file types listed here.


|||||
|:-----|:-----|:-----|:-----|
|.ade|.ins|.mda|.scr|
|.adp|.isp|.mdb|.sct|
|.app|.its|.mde|.shb|
|.asp|.js|.mdt|.shs|
|.bas|.jse|.mdw|.tmp|
|.bat|.ksh|.mdz|.url|
|.cer|.lnk|.msc|.vb|
|.chm|.mad|.msi|.vbe|
|.cmd|.maf|.msp|.vbs|
|.com|.mag|.mst|.vsmacros|
|.cpl|.mam|.ops|.vss|
|.crt|.maq|.pcd|.vst|
|.csh|.mar|.pif|.vsw|
|.exe|.mas|.prf|.ws|
|.fxp|.mat|.prg|.wsc|
|.hlp|.mau|.pst|.wsf|
|.hta|.mav|.reg|.wsh|
|.inf|.maw|.scf||

## Events



|**Name**|
|:-----|
|[AfterUpdate](http://msdn.microsoft.com/library/09dfe871-0e56-38fc-46d2-c517ea795907%28Office.15%29.aspx)|
|[AttachmentCurrent](http://msdn.microsoft.com/library/4b81608a-d591-7ce2-0075-8d841a825a9f%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/0437e831-b96f-60b6-1a7c-3e1f720394b7%28Office.15%29.aspx)|
|[Change](http://msdn.microsoft.com/library/5b34517d-f3a8-a10d-1bc3-ed3bc8ecc484%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/cdeff1db-5d95-dab5-79ae-d02ac25d5659%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/abc31523-5154-2d91-67c0-03cc0e73e957%28Office.15%29.aspx)|
|[Dirty](http://msdn.microsoft.com/library/d211238b-cbe4-f0ef-471b-33c1ced1aa9b%28Office.15%29.aspx)|
|[Enter](http://msdn.microsoft.com/library/0ca691d8-aace-3240-c7c7-acfb69960f4a%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/a083d56d-7a57-6874-14e6-c830f598a950%28Office.15%29.aspx)|
|[GotFocus](http://msdn.microsoft.com/library/9c841973-cb31-2ec6-d593-97ad8803250b%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/91a000e2-0a4e-4dd0-2715-b1987eb7212a%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/fc54afea-35ca-e354-1223-c7f3d5cf00b0%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/4b13f772-12e7-b840-029a-3736df1a9645%28Office.15%29.aspx)|
|[LostFocus](http://msdn.microsoft.com/library/b2a680bb-faec-bc7d-c568-3c827ee5d6b1%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/45056b32-a019-1284-35e4-fefab6ba2e3e%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/61ec0bdb-6e39-a4a7-92aa-45d543e35109%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/af4d03e6-af13-d91f-168f-70e90783aa2a%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Back](http://msdn.microsoft.com/library/96a8625a-2565-134b-e46e-52567ab08690%28Office.15%29.aspx)|
|[Forward](http://msdn.microsoft.com/library/0fec305d-b2b9-29a4-c756-2f3e59679316%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/cd807ce2-79b8-0873-c035-7927bc91967d%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/6af04ea8-02cb-9eda-439d-6c69cd772891%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/53e9c9f3-f1b8-f68d-8e9a-8b15ab4a3e83%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/9e9b8a65-79ba-9fda-08d8-9b5444678228%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddColon](http://msdn.microsoft.com/library/8eeb0085-e880-50ff-1e9f-3ae48d5bc6de%28Office.15%29.aspx)|
|[AfterUpdate](http://msdn.microsoft.com/library/556fc6d2-3936-5cc7-0c4f-03274f00cfc2%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/db88250d-da59-300c-6f0c-3768c1bb8a7f%28Office.15%29.aspx)|
|[AttachmentCount](http://msdn.microsoft.com/library/30c3bc2b-d6d5-8f83-8154-d451ab3a32ed%28Office.15%29.aspx)|
|[AutoLabel](http://msdn.microsoft.com/library/09007508-f7b4-3fa6-2548-a78afd34bd0c%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/7a96f656-4ca5-ebf8-47d9-7fe1f4939517%28Office.15%29.aspx)|
|[BackShade](http://msdn.microsoft.com/library/23a28b72-b30c-4b2c-77c9-51bb0099efe9%28Office.15%29.aspx)|
|[BackStyle](http://msdn.microsoft.com/library/7e86f99d-a74a-8153-64ef-fe7cea81d218%28Office.15%29.aspx)|
|[BackThemeColorIndex](http://msdn.microsoft.com/library/c1f88ca4-825e-4a35-2896-60d982a36819%28Office.15%29.aspx)|
|[BackTint](http://msdn.microsoft.com/library/f81ef313-0b84-a061-c58d-e433b01167f4%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/44a17114-bbb6-8ec9-89b5-db09cf60de98%28Office.15%29.aspx)|
|[BorderColor](http://msdn.microsoft.com/library/cd43f030-f832-c58a-a374-67a349c3d499%28Office.15%29.aspx)|
|[BorderShade](http://msdn.microsoft.com/library/39792e3b-e10d-98e8-4fcc-cb95fac69ce1%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/a1156f6c-5649-ddef-619d-d15a57bb581a%28Office.15%29.aspx)|
|[BorderThemeColorIndex](http://msdn.microsoft.com/library/a1ee1ca4-74d4-5e8e-e2b7-fb44cd7f3617%28Office.15%29.aspx)|
|[BorderTint](http://msdn.microsoft.com/library/819768fc-1473-5f7e-c320-b2d25d1b83d3%28Office.15%29.aspx)|
|[BorderWidth](http://msdn.microsoft.com/library/e72672a1-3b17-ad1b-ff7d-96e3652a9f35%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/b84a725d-0a4a-b105-ef2b-7355601181ec%28Office.15%29.aspx)|
|[ColumnHidden](http://msdn.microsoft.com/library/772c644e-b180-25ad-5566-c0b5dc6dbc41%28Office.15%29.aspx)|
|[ColumnOrder](http://msdn.microsoft.com/library/e11872da-df82-83e0-0c6f-8716989622dd%28Office.15%29.aspx)|
|[ColumnWidth](http://msdn.microsoft.com/library/1827efbf-f481-7e26-0638-775a522b2c46%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/767fd173-4a85-48ac-820a-9235776b7b00%28Office.15%29.aspx)|
|[ControlSource](http://msdn.microsoft.com/library/6c1f2351-5671-51dd-0ff7-964719d91b9c%28Office.15%29.aspx)|
|[ControlTipText](http://msdn.microsoft.com/library/c5dd9325-b545-d25e-10bf-7d58f7806e04%28Office.15%29.aspx)|
|[ControlType](http://msdn.microsoft.com/library/f660ca13-59f0-efae-8e6b-7449662a15c2%28Office.15%29.aspx)|
|[CurrentAttachment](http://msdn.microsoft.com/library/894b6b94-2fa2-66b9-8e18-925c77241fef%28Office.15%29.aspx)|
|[DefaultPicture](http://msdn.microsoft.com/library/98bc9637-50c9-5831-8170-a32abe5915bc%28Office.15%29.aspx)|
|[DefaultPictureType](http://msdn.microsoft.com/library/77032908-5b98-7072-1e53-520485580746%28Office.15%29.aspx)|
|[DisplayAs](http://msdn.microsoft.com/library/a8813925-8062-501a-a985-27084c2033f4%28Office.15%29.aspx)|
|[DisplayWhen](http://msdn.microsoft.com/library/7029a8ef-6672-7a30-deb4-581f4f66ce7f%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/d0838624-4ed8-6099-8aac-ea947de2f56e%28Office.15%29.aspx)|
|[EventProcPrefix](http://msdn.microsoft.com/library/f58670ff-b42c-69eb-0561-90ce5cc40d19%28Office.15%29.aspx)|
|[FileName](http://msdn.microsoft.com/library/4d39020e-b21f-35b5-f5dc-4d8d5b4fdc88%28Office.15%29.aspx)|
|[FileType](http://msdn.microsoft.com/library/0e22ddf6-695a-f6bc-58c8-f6af77912306%28Office.15%29.aspx)|
|[FileURL](http://msdn.microsoft.com/library/661ce36f-77f8-be34-845f-a3c450b878bf%28Office.15%29.aspx)|
|[GridlineColor](http://msdn.microsoft.com/library/ee74a374-ad6b-e002-cc02-41861192923c%28Office.15%29.aspx)|
|[GridlineShade](http://msdn.microsoft.com/library/24b5e8fa-7416-b312-7d2f-75c3b60e4617%28Office.15%29.aspx)|
|[GridlineStyleBottom](http://msdn.microsoft.com/library/337fb2fd-0f4b-f113-826a-661a03333085%28Office.15%29.aspx)|
|[GridlineStyleLeft](http://msdn.microsoft.com/library/0a67119c-035e-157d-f47d-4f5cd3f356c8%28Office.15%29.aspx)|
|[GridlineStyleRight](http://msdn.microsoft.com/library/c91f1804-656b-1c5d-84c5-3ac51a57ec20%28Office.15%29.aspx)|
|[GridlineStyleTop](http://msdn.microsoft.com/library/bf54b66f-f16f-195a-9fcc-37cfa6b69de3%28Office.15%29.aspx)|
|[GridlineThemeColorIndex](http://msdn.microsoft.com/library/195122c2-c972-8d39-aea6-bf2b531b1f84%28Office.15%29.aspx)|
|[GridlineTint](http://msdn.microsoft.com/library/c1730e7b-88ae-3810-1a6c-9a0ff17b95b1%28Office.15%29.aspx)|
|[GridlineWidthBottom](http://msdn.microsoft.com/library/47465946-1888-d2f5-a577-44e5c2fa80c2%28Office.15%29.aspx)|
|[GridlineWidthLeft](http://msdn.microsoft.com/library/2b6bce3c-e1b2-0ce8-c4d6-0c3e160e50cb%28Office.15%29.aspx)|
|[GridlineWidthRight](http://msdn.microsoft.com/library/743ab25e-44ae-b5d1-7c7b-4f31a91a8f17%28Office.15%29.aspx)|
|[GridlineWidthTop](http://msdn.microsoft.com/library/734f4aef-7233-7fd1-f0e2-bb782b7b6262%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/377565ec-9e10-2a3f-5d05-e1440707dc9c%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/a9eceafb-48b4-8bcd-bec1-6a16c71b4850%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/0d2aec7f-caa7-4779-fe39-4abe9f1465c6%28Office.15%29.aspx)|
|[InSelection](http://msdn.microsoft.com/library/e17586b5-0619-e713-e1fa-f27c9e26b561%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/5f8e4bcc-f304-09df-de50-ca994bb07420%28Office.15%29.aspx)|
|[LabelAlign](http://msdn.microsoft.com/library/328832bf-303b-1988-11b9-4e9505fe80de%28Office.15%29.aspx)|
|[LabelX](http://msdn.microsoft.com/library/6786c91f-32e6-39b1-b9d7-105463a7c103%28Office.15%29.aspx)|
|[LabelY](http://msdn.microsoft.com/library/4ac59667-61bf-925c-a70a-0857fabcf2e1%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/1256c89d-96d6-20de-1a37-31c92e5e6ddb%28Office.15%29.aspx)|
|[LayoutID](http://msdn.microsoft.com/library/e66ced6f-59de-b7ec-6b15-28825f154992%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/cee21215-a0b0-9247-976d-9f7899287e54%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/6b6d0829-1c61-db95-f955-863df4827972%28Office.15%29.aspx)|
|[Locked](http://msdn.microsoft.com/library/d35da857-2f8a-9d7f-19d2-6d7fbe029c76%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/5f4eaa10-4f7c-70ee-f408-23f3b4135ce2%28Office.15%29.aspx)|
|[OldBorderStyle](http://msdn.microsoft.com/library/abbc1a8d-d9cc-b917-026d-a1847739c362%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/f3c20114-fc68-98ab-03de-0e023aacaaf1%28Office.15%29.aspx)|
|[OnAttachmentCurrent](http://msdn.microsoft.com/library/7987943b-5283-e9dc-17a6-5f4b54c90d4d%28Office.15%29.aspx)|
|[OnChange](http://msdn.microsoft.com/library/c2c12032-463a-2e3e-f434-defce71c8138%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/c1d1ddcb-db80-e0e1-4318-0cf9477d7316%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/5bfe9633-dd3a-d1d5-450b-eafbc1a607c1%28Office.15%29.aspx)|
|[OnDirty](http://msdn.microsoft.com/library/a3f0e108-3abe-23b2-6c7d-e528432fc3d9%28Office.15%29.aspx)|
|[OnEnter](http://msdn.microsoft.com/library/5aae3814-5fff-2011-c86d-3765f2a3615d%28Office.15%29.aspx)|
|[OnExit](http://msdn.microsoft.com/library/5ca25e6f-1fc3-826a-9111-b899e324ef99%28Office.15%29.aspx)|
|[OnGotFocus](http://msdn.microsoft.com/library/a25aa4f5-8ac6-86e9-d8de-725072a77007%28Office.15%29.aspx)|
|[OnKeyDown](http://msdn.microsoft.com/library/8135c3e5-e7d0-bafa-3eef-740b6ee73edd%28Office.15%29.aspx)|
|[OnKeyPress](http://msdn.microsoft.com/library/9f3213eb-9c37-f550-6c14-e6dd85d030a5%28Office.15%29.aspx)|
|[OnKeyUp](http://msdn.microsoft.com/library/56e5a246-5907-f537-0c89-a746beab0865%28Office.15%29.aspx)|
|[OnLostFocus](http://msdn.microsoft.com/library/546d0491-ddb8-87d4-9f97-d68cfd96070c%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/71ba8a45-7814-4939-b8cf-9b07d9e04b4d%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/4bf67a8a-1c54-d67c-b93d-1cfd98e59e70%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/077568b6-2053-7ddb-9afe-503b8a9850a5%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/a1fe6219-650c-4a2b-4121-0de42109dc3f%28Office.15%29.aspx)|
|[PictureAlignment](http://msdn.microsoft.com/library/505daae0-8321-cce0-028a-ff6c2ac16245%28Office.15%29.aspx)|
|[PictureSizeMode](http://msdn.microsoft.com/library/07d268ad-d4ba-c9ba-1ef4-7b3e7911ebba%28Office.15%29.aspx)|
|[PictureTiling](http://msdn.microsoft.com/library/d7eb8047-ea1d-e864-d2d7-51cd340cbc63%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/38e9513e-3297-6f82-9072-7e03c2e3e22e%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/ade6bde4-ebea-36af-f0ad-f071260dbf00%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/fb08e629-6056-85ac-4eae-2d7ab88916b9%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/be4ce61e-c4a9-9e3b-e2f4-187b77451f67%28Office.15%29.aspx)|
|[SpecialEffect](http://msdn.microsoft.com/library/4c2a07d1-99b4-1558-7956-d4a8d4fa157d%28Office.15%29.aspx)|
|[StatusBarText](http://msdn.microsoft.com/library/9d947d11-edb1-947a-df0c-727ef9b1599a%28Office.15%29.aspx)|
|[TabIndex](http://msdn.microsoft.com/library/bca68c61-a795-34d9-9e42-97113f1d4387%28Office.15%29.aspx)|
|[TabStop](http://msdn.microsoft.com/library/918d81a6-a9a2-ab4e-6fb3-ad78233b6e7f%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/6d3e9f35-1986-e6b4-5f35-2652123c007c%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/051c785a-7e71-fb5c-c00c-c86bdaf7194b%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/18c89f2e-e276-6c9f-b317-5fa931dd7003%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/bbb588c4-ec99-1352-4f1b-fd166d67df33%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/15606b3e-dffb-f179-021a-5bf8087003a7%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/708c9f0d-deaa-1149-2ce7-53f0b5ec8c25%28Office.15%29.aspx)|

## See also


#### Other resources


[Attachment Object Members](http://msdn.microsoft.com/library/4294b913-7691-5f45-2c20-5137c2320620%28Office.15%29.aspx)
[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
