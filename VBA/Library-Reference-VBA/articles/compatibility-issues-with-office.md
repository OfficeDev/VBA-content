# Compatibility issues with Office

Get more information about issues displayed in the telemetry log about possible compatibility issues in Office products.

The following tables list information about issues presented in the telemetry log. For more information about the telemetry log, see [Troubleshooting Office files and custom solutions with the telemetry log](https://msdn.microsoft.com/en-us/library/office/jj230106.aspx).

For information about features that have been changed or removed since Office 2013, see [Changes in Office 2016 for Windows](https://technet.microsoft.com/library/mt715497(v=office.16).aspx).

## Controls

These messages can appear if the file contains a control that may not be supported in Office or on the computer operating system.

Table 1. Issues displayed in the telemetry log about controls

<table>
  <tr>
    <th scope="col">
      <p>Event ID</p>
    </th>
    <th scope="col">
      <p>Introduced in version</p>
    </th>
    <th scope="col">
      <p>Applications affected</p>
    </th>
    <th scope="col">
      <p>Additional information</p>
    </th>
    <th scope="col">
      <p>Title</p>
    </th>
    <th scope="col">
      <p>Description</p>
    </th>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10000</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>All Office 2013</p>
    </td>
    <td data-th="Additional information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>Warning: Visual Basic 6.0 Controls</p>
    </td>
    <td data-th="Description">
      <p>The file uses a Visual Basic 6.0 control that does not work in 64-bit versions of Office or in 32-bit versions of Office that are running on a device that uses an ARM processor. Replace the control with a supported control if you want it to be available to Office applications in those environments.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10001</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>All Office 2013</p>
    </td>
    <td data-th="Additional information">
      <p>
        <a href="https://msdn.microsoft.com/en-us/vbasic/ms788708.aspx" target="_blank">Link</a>
      </p>
    </td>
    <td data-th="Title">
      <p>Controls: Visual Basic 6.0 Control on 64-bit OS</p>
    </td>
    <td data-th="Description">
      <p>The file uses a Visual Basic 6.0 control that does not work in 64-bit versions of Office. Visual Basic 6.0 runtime files are 32-bit and are supported in the 32-bit OS or in WOW emulation environments only.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10002</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>All Office 2013</p>
    </td>
    <td data-th="Additional information">
      <p>
        <a href="https://msdn.microsoft.com/en-us/vbasic/ms788708.aspx" target="_blank">Link</a>
      </p>
    </td>
    <td data-th="Title">
      <p>Controls: Visual Basic 6.0 Controls on Device with ARM Processor</p>
    </td>
    <td data-th="Description">
      <p>The file uses a Visual Basic 6.0 control that does not work on a device that uses an ARM processor.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10003</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>All Office 2013</p>
    </td>
    <td data-th="Additional information">
      <p>
        <a href="https://technet.microsoft.com/en-us/library/cc179181.aspx" target="_blank">Link</a>
      </p>
    </td>
    <td data-th="Title">
      <p>Controls: Microsoft Calendar Control</p>
    </td>
    <td data-th="Description">
      <p>The file uses the Microsoft Calendar control (Mscal.ocx), a feature of previous versions of Access that is not available in Office 2013. The control will not work because it is not installed on the host computer. Use other date picker controls as an alternative, like the <span class="input">Date Picker Content Control</span> (in Word 2013) or the Windows <span class="input">DatePicker</span> control (in the Windows Common Controls). </p>
      <p>For more information, see <a href="https://msdn.microsoft.com/library/dc6ba80d-b1fa-4596-b484-5e729cae4d70" target="_blank">Replacing the Calendar Control in Access 2010 Applications</a>.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10004</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>All Office 2013</p>
    </td>
    <td data-th="Additional information">
      <p>
        <a href="http://support.microsoft.com/kb/972129" target="_blank">Link</a>
      </p>
    </td>
    <td data-th="Title">
      <p>Office Web Components</p>
    </td>
    <td data-th="Description">
      <p>The file uses an Office Web Components (MSOWC.dll) control. The control will not work because the Office Web Components are not installed on this computer and are not included with Office 2013. To use this control, install the Office Web Components separately.</p>
      <p>For more information, see <a href="http://support.microsoft.com/kb/319793" target="_blank">HOW TO: Find Office Web Components Programming Documentation and Samples</a></p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10005</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>All Office 2013</p>
    </td>
    <td data-th="Additional information">
      <p>
        <a href="http://office.microsoft.com/en-us/access-help/embedded-object-and-activex-control-policy-settings-error-HA101825674.aspx?CTT=1" target="_blank">Link</a>
      </p>
    </td>
    <td data-th="Title">
      <p>Controls: Unregistered ActiveX Control</p>
    </td>
    <td data-th="Description">
      <p>The file uses ActiveX controls that are not registered on the host computer. To use the control, register it on the host computer.</p>
    </td>
  </tr>
</table>

## Removed and deprecated members in the Object Model

These messages can appear if the add-in or macro-enabled document code uses an object, member, collection, enumeration, or constant that has been removed from the application’s object model.

Table 2. Issues displayed in the telemetry log about removed and deprecated members

<table Responsive="true" summary="table">
  <tr Responsive="true">
    <th scope="col">
      <p>Event ID</p>
    </th>
    <th scope="col">
      <p>Introduced in version</p>
    </th>
    <th scope="col">
      <p>Applications affected</p>
    </th>
    <th scope="col">
      <p>Additional Information</p>
    </th>
    <th scope="col">
      <p>Title</p>
    </th>
    <th scope="col">
      <p>Description</p>
    </th>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10103</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013, Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>
        <a href="http://support.microsoft.com/kb/2445062" target="_blank">Link</a>
      </p>
    </td>
    <td data-th="Title">
      <p>OM Removed: Custom XML Feature</p>
    </td>
    <td data-th="Description">
      <p>The Custom XML feature is removed from Word. The following methods and properties are hidden, and if accessed, return run-time error: </p>
      <ul>
        <li>
          <p>
            <span class="input">XMLNodes.Add</span> method</p>
        </li>
        <li>
          <p>
            <span class="input">Document.XMLHideNamespaces</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">Document.XMLSaveDataOnly</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">Document.XMLSchemaViolations</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">XMLSchemaViolations</span> object and all its members</p>
        </li>
        <li>
          <p>
            <span class="input">XMLSchemaViolation</span> object and all its members</p>
        </li>
        <li>
          <p>
            <span class="input">Application.TaskPanes</span>, if the <span class="input">wdTaskPaneXMLStructure</span> constant (5) of the <span class="input">WdTaskPanes</span> enumeration is specified</p>
        </li>
        <li>
          <p>
            <span class="input">Options.PrintXMLTag</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">View.ShowXMLMarkup</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">XMLChildNodeSuggestions</span> collection and all its members</p>
        </li>
        <li>
          <p>
            <span class="input">XMLChildNodeSuggestion</span> object and all its members</p>
        </li>
        <li>
          <p>
            <span class="input">Selection.XMLParentNode</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">Range.XMLParentNode</span> property</p>
        </li>
      </ul>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10113</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013, Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Removed: Smart Tag Feature</p>
    </td>
    <td data-th="Description">
      <p>The SmartTags feature is removed from Word. The following objects, methods, and properties are hidden, and if accessed, return a runtime error: </p>
      <ul>
        <li>
          <p>
            <span class="input">SmartTag</span> object and members</p>
        </li>
        <li>
          <p>
            <span class="input">SmartTags</span> collection and members</p>
        </li>
        <li>
          <p>
            <span class="input">SmartTagAction</span> object and members</p>
        </li>
        <li>
          <p>
            <span class="input">SmartTagActions</span> collection and members</p>
        </li>
        <li>
          <p>
            <span class="input">SmartTagType</span> object and members</p>
        </li>
        <li>
          <p>
            <span class="input">SmartTagTypes</span> collection and members</p>
        </li>
        <li>
          <p>
            <span class="input">XMLNode.SmartTag</span> property</p>
        </li>
      </ul>
      <p>The following methods are hidden, and if accessed, fail silently: </p>
      <ul>
        <li>
          <p>
            <span class="input">Document.CheckNewSmartTags</span> method</p>
        </li>
        <li>
          <p>
            <span class="input">Document.RecheckSmartTags</span> method</p>
        </li>
        <li>
          <p>
            <span class="input">Document.RemoveSmartTags</span> method</p>
        </li>
      </ul>
      <p>The following properties are hidden, and if accessed, always returns FALSE: </p>
      <ul>
        <li>
          <p>
            <span class="input">Document.EmbedSmartTags</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">Document.SmartTagsAsXMLProps</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">Options.LabelSmartTags</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">Options.DisplaySmartTagButtons</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">EmailOptions.EmbedSmartTag</span> property</p>
        </li>
      </ul>
      <p>The following property is hidden, and if accessed, always returns true: </p>
      <ul>
        <li>
          <p>
            <span class="input">View.DisplaySmartTags</span> property</p>
        </li>
      </ul>
      <p>The following properties are hidden, and if accessed, always returns an empty collection: </p>
      <ul>
        <li>
          <p>
            <span class="input">Application.SmartTagTypes</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">Document.SmartTags</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">Range.SmartTags</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">Selection.SmartTags</span> property</p>
        </li>
      </ul>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10115</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013, Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Removed: AutoSummary Feature</p>
    </td>
    <td data-th="Description">
      <p>The AutoSummary feature is removed from Word. The following method and properties are hidden, and if accessed, return a runtime error: </p>
      <ul>
        <li>
          <p>
            <span class="input">Document.AutoSummarize</span> method</p>
        </li>
        <li>
          <p>
            <span class="input">Document.ShowSummary</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">Document.SummaryViewMode</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">Document.SummaryLength</span> property</p>
        </li>
      </ul>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10116</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013, Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Removed: Barcode Feature</p>
    </td>
    <td data-th="Description">
      <p>The barcode feature for envelopes is removed from Word. The following  properties are hidden, and if accessed, always return FALSE: </p>
      <ul>
        <li>
          <p>
            <span class="input">Envelope.DefaultPrintBarCode</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">MailingLabel.DefaultPrintBarCode</span> property</p>
        </li>
      </ul>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10117</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013, Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Removed: Window.DocumentMapPercentWidth Property</p>
    </td>
    <td data-th="Description">
      <p>The <span class="input">Window.DocumentMapPercentWidth</span> property is hidden in Word. If accessed, the property raises a run-time error.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10122</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013, Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Removed: Application.FileSearch</p>
    </td>
    <td data-th="Description">
      <p>The <span class="input">Application.FileSearch</span> was removed in Office 2007. If accessed, this property will return an error. To work around this issue, use the <a href="https://msdn.microsoft.com/en-us/library/office/gg278516.aspx">FileSystemObject</a> to recursively search directories to find specific files.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10145</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Excel 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Removed: Application.FileSearch</p>
    </td>
    <td data-th="Description">
      <p>The <span class="input">Application.FileSearch</span> property was removed in Office 2007. If accessed, this property will return an error. To work around this issue, use the <a href="https://msdn.microsoft.com/en-us/library/office/gg278516.aspx">FileSystemObject</a> to recursively search directories to find specific files.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10154</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Excel 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Removed: Smart Tag Feature</p>
    </td>
    <td data-th="Description">
      <p>The SmartTags feature is removed from Excel. The following properties are hidden, and if accessed, always returns FALSE: </p>
      <ul>
        <li>
          <p>
            <span class="input">Application.SmartTagRecognizers</span> property</p>
        </li>
      </ul>
      <p>The following objects, methods, and properties are hidden, and if accessed, return a runtime error: </p>
      <ul>
        <li>
          <p>
            <span class="input">SmartTag</span> object and members</p>
        </li>
        <li>
          <p>
            <span class="input">SmartTags</span> collection and members</p>
        </li>
        <li>
          <p>
            <span class="input">SmartTagAction</span> object and members</p>
        </li>
        <li>
          <p>
            <span class="input">SmartTagActions</span> collection and members</p>
        </li>
        <li>
          <p>
            <span class="input">SmartTagOptions</span> collection and members</p>
        </li>
        <li>
          <p>
            <span class="input">SmartTagRecognizer</span> object and members</p>
        </li>
        <li>
          <p>
            <span class="input">SmartTagRecognizers</span> collection and members</p>
        </li>
      </ul>
      <p>The following methods are hidden, and if accessed, fail silently: </p>
      <ul>
        <li>
          <p>
            <span class="input">Workbook.RecheckSmartTags</span> method</p>
        </li>
      </ul>
      <p>The following properties are hidden, and if accessed, always returns an empty collection: </p>
      <ul>
        <li>
          <p>
            <span class="input">Workbook.SmartTagOptions</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">Worksheet.SmartTags</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">Range.SmartTags</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">IRange.SmartTags</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">DialogSheet.SmartTags</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">IDialogSheet.SmartTags</span> property</p>
        </li>
      </ul>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10155</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>All Office 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Removed: ToolbarButton.Edit Method</p>
    </td>
    <td data-th="Description">
      <p>The CommandBar Button Editor has been removed. If called, the method silently fails. Custom images can be applied to legacy CommandBar buttons using the <a href="https://msdn.microsoft.com/en-us/library/office/ff860599.aspx">CommandBarButton.PasteFace</a> method, or the <a href="https://msdn.microsoft.com/en-us/library/office/ff864041.aspx">CommandBarButton.Picture</a> and the <a href="https://msdn.microsoft.com/en-us/library/office/ff864960.aspx">CommandBarButton.Mask</a> properties.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10159</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2016</p>
    </td>
    <td data-th="Applications affected">
      <p>Word</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Deprecated: SkyDriveSignInOption</p>
    </td>
    <td data-th="Description">
      <p>SkyDriveSignInOption has been deprecated. Use CloudSignInOption instead.</p>
    </td>
  </tr>
</table>

## Behavior changes in the Object Model

These messages can appear if the add-in or macro-enabled document code uses an object, member, collection, enumeration, or constant that behaves differently from previous versions of Office.

Table 3. Issues displayed in the telemetry log about behavior changes

<table Responsive="true" summary="table">
  <tr Responsive="true">
    <th scope="col">
      <p>Event ID</p>
    </th>
    <th scope="col">
      <p>Introduced in version</p>
    </th>
    <th scope="col">
      <p>Applications affected</p>
    </th>
    <th scope="col">
      <p>Additional Information</p>
    </th>
    <th scope="col">
      <p>Title</p>
    </th>
    <th scope="col">
      <p>Description</p>
    </th>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10156</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2016</p>
    </td>
    <td data-th="Applications affected">
      <p>Word</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Behavior Change: Use of save events detected</p>
    </td>
    <td data-th="Description">
      <p>The compatibility checker has detected use of save events which may cause an undesirable experience during real-time co-authoring. Your solution may not work as intended during real time co-authoring sessions due to the higher save frequency during those scenarios. We recommend adjusting the solution to throttle during frequent saves. Alternatively, disable real time co-authoring by using Group Policy.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10160</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2016</p>
    </td>
    <td data-th="Applications affected">
      <p>Word, Excel, PowerPoint</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Behavior Change: Application.DisplayDocumentInformationPanel</p>
    </td>
    <td data-th="Description">
      <p>The Document Information Panel has been deprecated as part of InfoPath product deprecation. Querying this property will always return false. Setting this property varies by application. Setting it to true will show the Property Panel for Word and PowerPoint and do nothing in Excel. Setting it to false does nothing in all apps.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10161</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2016</p>
    </td>
    <td data-th="Applications affected">
      <p>Word</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Behavior Change: ContentControl.DropdownListEntries</p>
    </td>
    <td data-th="Description">
      <p>The Document Information Panel has been deprecated as part of InfoPath product deprecation. When acting against SharePoint lookup properties the behavior of this API is no longer supported. It works as expected with other types of list entries.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10157</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2016</p>
    </td>
    <td data-th="Applications affected">
      <p>PowerPoint</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Behavior Change: Presentation.InMergeMode Property</p>
    </td>
    <td data-th="Description">
      <p>The old merge mode that appears in the document window when co-authoring has been replaced with a new conflict resolution window. If accessed in this situation, the Presentation.InMergeMode property will return False.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10106</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Excel 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Behavior Change: Application.FormulaBarHeight Property</p>
    </td>
    <td data-th="Description">
      <p>The <span><a href="https://msdn.microsoft.com/en-us/library/office/ff841264.aspx">Application.FormulaBarHeight Property (Excel)</a></span> property is changed. If accessed, the property reads and writes the height of the formula bar associated with the active window in Excel. To change formula bar height for another window in Excel, set the <span class="input">Application.FormulaBarHeight</span> property after activating the window.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10107</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Excel 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Behavior Change: Workbook.Protect Method</p>
    </td>
    <td data-th="Description">
      <p>Window structure (height, width, minimized or maximized state) cannot be protected in Excel. If called, the <span><a href="https://msdn.microsoft.com/en-us/library/office/ff193800.aspx">Workbook.Protect Method (Excel)</a></span> method does not protect the workbook window structure regardless of the value of the Windows parameter.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10140</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013, Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Behavior Change: Table.AllowPageBreaks</p>
    </td>
    <td data-th="Description">
      <p>The <span class="input">Table.AllowPageBreaks</span> property is hidden, and always returns TRUE. To achieve the same behavior, use the <span><a href="https://msdn.microsoft.com/en-us/library/office/ff821554.aspx">ParagraphFormat.KeepTogether Property (Word)</a></span> and <span><a href="https://msdn.microsoft.com/en-us/library/office/ff196927.aspx">ParagraphFormat.KeepWithNext Property (Word)</a></span> properties.</p>
    </td>
  </tr>
</table>

## Hidden members in the Object Model

These messages can appear if the add-in or macro-enabled document code uses an object, member, collection, enumeration, or constant that has been hidden in the application’s object model.

Table 4. Issues displayed in the telemetry log about hidden members

<table Responsive="true" summary="table">
  <tr Responsive="true">
    <th scope="col">
      <p>Event ID</p>
    </th>
    <th scope="col">
      <p>Introduced in version</p>
    </th>
    <th scope="col">
      <p>Applications affected</p>
    </th>
    <th scope="col">
      <p>Additional Information</p>
    </th>
    <th scope="col">
      <p>Title</p>
    </th>
    <th scope="col">
      <p>Description</p>
    </th>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10158</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2016</p>
    </td>
    <td data-th="Applications affected">
      <p>Excel</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Presentation.WorksheetFunction.Forecast (All) Method</p>
    </td>
    <td data-th="Description">
      <p>WorksheetFunction.Forecast method is hidden. If called, the method behaves similarly as it does in Excel 2013. It remains part of the object model for backward compatibility, but you should use WorksheetFunction.Forecast_Linear in new applications.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10109</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013, Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Document.UpdateSummaryProperties Method</p>
    </td>
    <td data-th="Description">
      <p>The AutoSummary feature is removed from Word. If called, the <span class="input">Document.UpdateSummaryProperties</span> method raises a run-time error.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10110</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013, Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Comment.Delete Method</p>
    </td>
    <td data-th="Description">
      <p>Commenters can reply directly to other comments in Word. If called, the <span class="input">Comment.Delete</span> method functions similar to previous versions of Office by deleting a single comment and leaving all replies in the document. To remove an entire thread of comments, use the <span class="input">Comment.DeleteRecursively</span> method. To reply to a comment, use the <span class="input">Comment.Replies.Add</span> method.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10111</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013, Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Comment.Author Property</p>
    </td>
    <td data-th="Description">
      <p>Comments in Word are now associated with contacts. If accessed, the <span class="input">Comment.Author</span> property behaves similarly to previous versions of Office. To access the name of a commenter, use the Name property of the <span class="input">Contact</span> object associated with the comment.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10112</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013, Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Comment.Initial Property</p>
    </td>
    <td data-th="Description">
      <p>Initials of commenters are not displayed with comments in Word by default. If accessed, the <span class="input">Comment.Initial</span> property behaves similar to previous versions of Office. However, printed documents still display initials for comments.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10114</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013, Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Comment.ShowTip Property</p>
    </td>
    <td data-th="Description">
      <p>ScreenTips associated with comments in Word are shown by default. If accessed, the <span class="input">Comment.ShowTip</span> property always returns FALSE.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10118</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013, Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Options.BackgroundOpen Property</p>
    </td>
    <td data-th="Description">
      <p>Large web documents cannot be opened in the background in Word. If accessed, the <span><a href="https://msdn.microsoft.com/en-us/library/office/ff840248.aspx">Options.BackgroundOpen Property (Word)</a></span> property always returns FALSE and cannot be set to any other value.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10119</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013, Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Document.ApplyQuickStyleSet Method</p>
    </td>
    <td data-th="Description">
      <p>The <span class="input">Document.ApplyQuickStyleSet</span> method is hidden in Word. If called, the method continues to function the same as it did in Office 2007 by changing the Style Set for the document. To use the new features of Office 2010 and above, replace with the <span><a href="https://msdn.microsoft.com/en-us/library/office/ff821672.aspx">Document.ApplyQuickStyleSet2 Method (Word)</a></span> method.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10120</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013, Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Document.SaveAs Method</p>
    </td>
    <td data-th="Description">
      <p>The Save As feature behaves similarly to previous versions of Word. If called, the <span class="input">Document.SaveAs</span> method behaves similarly as it does in Office 2007. And SaveAs2 method is added to the Document object that contains the properties introduced in Office 2010. To use the new features of Office 2010 and above, replace the <span class="input">Document.SaveAs</span> method with the <span><a href="https://msdn.microsoft.com/en-us/library/office/ff836084.aspx">Document.SaveAs2 Method (Word)</a></span>.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10121</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013, Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Assistant and AnswerWizard Features</p>
    </td>
    <td data-th="Description">
      <p>The Assistant and AnswerWizard features have been hidden in Word. </p>
      <p>The following properties are hidden but remain part of the object model for backward compatibility. It is not recommended to use them in new Office solutions: </p>
      <ul>
        <li>
          <p>
            <span class="input">Application.Assistant</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">Application.AnswerWizard</span> property</p>
        </li>
      </ul>
      <p>The following properties are hidden. If accessed, they return a run-time error.</p>
      <ul>
        <li>
          <p>
            <span class="input">Global.Assistant</span> property</p>
        </li>
        <li>
          <p>
            <span class="input">Global.AnswerWizard</span> property</p>
        </li>
      </ul>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10123</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013, Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Options.WPHelp</p>
    </td>
    <td data-th="Description">
      <p>The <span class="input">Options.WPHelp</span> property is hidden.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10124</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013,  Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Options.SetWPHelpOptions</p>
    </td>
    <td data-th="Description">
      <p>The <span class="input">Options.SetWPHelpOptions</span> property is hidden. If accessed, the property returns an error.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10125</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013, Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Options.WPDocNavKeys</p>
    </td>
    <td data-th="Description">
      <p>The <span class="input">Options.WPDocNavKeys</span> property is hidden. If accessed, the property always returns FALSE.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10126</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013,  Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden:Options.BlueScreen</p>
    </td>
    <td data-th="Description">
      <p>The <span class="input">Options.BlueScreen</span> property is hidden. If accessed, the property always returns FALSE.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10127</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013,  Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Options.AllowFastSave</p>
    </td>
    <td data-th="Description">
      <p>The <span class="input">Options.AllowFastSave</span> is hidden. If accessed, the property always returns FALSE.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10128</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013,  Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Application.DisplayStatusBar</p>
    </td>
    <td data-th="Description">
      <p>The <span class="input">Application.DisplayStatusBar</span> property is hidden. Use <span class="input">Application.CommandBars("Status Bar")</span>Visible instead.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10129</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Document.HTMLProject</p>
    </td>
    <td data-th="Description">
      <p>The <span class="input">Document.HTMLProject</span> is hidden. If accessed, the property returns an error.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10130</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013,  Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Document.Versions</p>
    </td>
    <td data-th="Description">
      <p>The Versions feature is removed, and as a result, the <span class="input">Document.Versions</span> property is hidden. If accessed, the property returns an error.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10131</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013,  Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Document.Route</p>
    </td>
    <td data-th="Description">
      <p>The Routing Slip feature is removed, and as a result, the <span class="input">Document.Route</span> method is hidden. If accessed, this method returns an error.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10132</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013,  Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Document.HasRoutingSlip</p>
    </td>
    <td data-th="Description">
      <p>The Routing Slip feature is removed, and as a result, the <span class="input">Document.HasRoutingSlip</span> property is hidden. If accessed, the property returns an error.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10133</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013,  Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Document.Routed</p>
    </td>
    <td data-th="Description">
      <p>The Routing Slip feature is removed, and as a result, the <span class="input">Document.Routed</span> property is hidden. If accessed, the property returns an error.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10134</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013,  Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Document.RoutingSlip</p>
    </td>
    <td data-th="Description">
      <p>The Routing Slip feature is removed, and as a result, the <span class="input">Document.RoutingSlip</span> property is hidden. If accessed, the property returns an error.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10135</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013,  Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Diagram OM</p>
    </td>
    <td data-th="Description">
      <p>The <span class="input">Diagram</span> object and the properties and methods associated with the <span class="input">Diagram</span> object have been hidden. If accessed, the following members generate errors: </p>
      <ul>
        <li>
          <p>
            <span class="input">Shapes.AddDiagram</span>
          </p>
        </li>
        <li>
          <p>
            <span class="input">Shape.Diagram</span>
          </p>
        </li>
        <li>
          <p>
            <span class="input">Shape.DiagramNode</span>
          </p>
        </li>
        <li>
          <p>
            <span class="input">Shape.HasDiagram</span>
          </p>
        </li>
        <li>
          <p>
            <span class="input">ShapeHasDiagramNode</span>
          </p>
        </li>
        <li>
          <p>
            <span class="input">ShapeRange.DiagramNode</span>
          </p>
        </li>
        <li>
          <p>
            <span class="input">ShapeRange.HasDiagram</span>
          </p>
        </li>
        <li>
          <p>
            <span class="input">ShapeRange.HasDiagramNode</span>
          </p>
        </li>
      </ul>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10136</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013,  Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: ShapeRange.Activate</p>
    </td>
    <td data-th="Description">
      <p>The Word Picture object is hidden, and as a result, the methods used to convert a picture to a Word Picture object were also hidden. These methods included the following: </p>
      <ul>
        <li>
          <p>
            <span class="input">InlineShape.Activate</span>
          </p>
        </li>
        <li>
          <p>
            <span class="input">Shape.Activate</span>
          </p>
        </li>
        <li>
          <p>
            <span class="input">ShapeRange.Activate</span>
          </p>
        </li>
      </ul>
      <p>If used, these methods generate an error.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10137</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013,  Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Shape.Activate</p>
    </td>
    <td data-th="Description">
      <p>The Word Picture object is hidden, and as a result, the methods used to convert a picture to a Word Picture object were also hidden. These methods included the following: </p>
      <ul>
        <li>
          <p>
            <span class="input">InlineShape.Activate</span>
          </p>
        </li>
        <li>
          <p>
            <span class="input">Shape.Activate</span>
          </p>
        </li>
        <li>
          <p>
            <span class="input">ShapeRange.Activate</span>
          </p>
        </li>
      </ul>
      <p>If used, these methods generate an error.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10138</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013,  Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: InlineShape.Activate</p>
    </td>
    <td data-th="Description">
      <p>The Word Picture object is hidden, and as a result, the methods used to convert a picture to a Word Picture object were also hidden. These methods included the following: </p>
      <ul>
        <li>
          <p>
            <span class="input">InlineShape.Activate</span>
          </p>
        </li>
        <li>
          <p>
            <span class="input">Shape.Activate</span>
          </p>
        </li>
        <li>
          <p>
            <span class="input">ShapeRange.Activate</span>
          </p>
        </li>
      </ul>
      <p>If used, these methods generate an error.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10139</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Shapes.AddChart</p>
    </td>
    <td data-th="Description">
      <p>The <span class="input">Shapes.AddChart</span> method is hidden. It remains part of the object model for backward compatibility, but you should not use it in new applications. Use the <span class="input">Shapes.AddChart2</span> method instead.</p>
      <div class="alert">
        <table summary="table">
          <tr>
            <th align="left" scope="col">
              <img id="alert_note" alt="Note" src="https://i-msdn.sec.s-msft.com/dynimg/IC589958.gif" title="Note" xmlns="" />
              <strong>Note</strong>
            </th>
          </tr>
          <tr>
            <td>
              <p>The <span class="input">Shapes.AddChart2</span> method applies a default title to the new chart. If you need to change the title of the chart after it has been added to the file, use the <span class="input">Chart.ChartTitle</span> property or edit the title manually. </p>
            </td>
          </tr>
        </table>
      </div>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10141</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013,  Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Application.ShowWindowsInTaskbar</p>
    </td>
    <td data-th="Description">
      <p>The <span class="input">Application.ShowWindowinTaskbar</span> property is hidden. If accessed, the property always returns true.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10142</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013,  Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: HangulHanjaConversionDictionaries.BuiltinDictionary</p>
    </td>
    <td data-th="Description">
      <p>The <span class="input">HangulHanjaConversionDictionaries.BuiltinDictionary</span> property is hidden. If accessed, the property returns NULL.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10143</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013,  Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Template.AutoTextEntries</p>
    </td>
    <td data-th="Description">
      <p>AutoText is now a type of Building Block. You can access Building Blocks by using the <span><a href="https://msdn.microsoft.com/en-us/library/office/ff195119.aspx">Template.BuildingBlockEntries Property (Word)</a></span> or <span><a href="https://msdn.microsoft.com/en-us/library/office/ff834280.aspx">Template.BuildingBlockTypes Property (Word)</a></span> properties. </p>
      <p>By default, AutoText is saved in normal.dotm</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10144</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Word 2013,  Outlook 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: View.RevisionsMode</p>
    </td>
    <td data-th="Description">
      <p>The <span class="input">View.RevisionsMode</span> property is hidden. Instead, use the <span><a href="https://msdn.microsoft.com/en-us/library/office/ff192820.aspx">View.MarkupMode Property (Word)</a></span> property.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10146</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Excel 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: ISlicerCache.ClearManualFilter</p>
    </td>
    <td data-th="Description">
      <p>The method <span class="input">ClearManualFilter</span> of ISlicerCache object was marked as hidden. It remains part of the object model for backward compatibility, but you should not use it in new applications.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10147</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Excel 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: _Application.ShowWindowsInTaskbar</p>
    </td>
    <td data-th="Description">
      <p>The property <span class="input">_Application.ShowWindowsInTaskbar</span> has been hidden. It remains part of the object model for backward compatibility, but you should not use it in new applications.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10148</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Excel 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: _Application.SaveISO8601Dates</p>
    </td>
    <td data-th="Description">
      <p>The property <span class="input">_Application.SaveISO8601Dates</span> has been hidden. It remains part of the object model for backward compatibility, but you should not use it in new applications.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10149</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Excel 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: SlicerCache.ClearManualFilter</p>
    </td>
    <td data-th="Description">
      <p>The method <span class="input">ClearManualFilter</span> of SlicerCache. object was marked as hidden. It remains part of the object model for backward compatibility, but you should not use it in new applications.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10150</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Excel 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: _Application.Assistant</p>
    </td>
    <td data-th="Description">
      <p>The property <span class="input">_Application.Assistant</span> has been hidden. It remains part of the object model for backward compatibility, but you should not use it in new applications.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10151</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Excel 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: _Application.AnswerWizard</p>
    </td>
    <td data-th="Description">
      <p>The property<span class="input"> _Application.Assistant</span> has been hidden. If accessed, the property returns a run-time error.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10152</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Excel 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: _Global.Assistant</p>
    </td>
    <td data-th="Description">
      <p>The property <span class="input">_Global.Assistant</span> has been hidden. It remains part of the object model for backward compatibility, but you should not use it in new applications.</p>
    </td>
  </tr>
  <tr>
    <td data-th="Event ID">
      <p>10153</p>
    </td>
    <td data-th="Introduced in version">
      <p>Office 2013</p>
    </td>
    <td data-th="Applications affected">
      <p>Excel 2013</p>
    </td>
    <td data-th="Additional Information">
      <p>

      </p>
    </td>
    <td data-th="Title">
      <p>OM Hidden: Shapes.AddChart</p>
    </td>
    <td data-th="Description">
      <p>The <span class="input">Shapes.AddChart</span> method is hidden. It remains part of the object model for backward compatibility, but you should not use it in new applications. Use the <span class="input">Shapes.AddChart2</span> method instead.</p>
      <div class="alert">
        <table summary="table">
          <tr>
            <th align="left" scope="col">
              <img id="alert_note" alt="Note" src="https://i-msdn.sec.s-msft.com/dynimg/IC589958.gif" title="Note" xmlns="" />
              <strong>Note</strong>
            </th>
          </tr>
          <tr>
            <td>
              <p>The <span class="input">Shapes.AddChart2</span> method applies a default title to the new chart. If you need to change the title of the chart after it has been added to the file, use the <span class="input">Chart.ChartTitle</span> property or edit the title manually. </p>
            </td>
          </tr>
        </table>
      </div>
    </td>
  </tr>
</table>

## Additional resources

- [Compatibility and telemetry in Office](https://technet.microsoft.com/library/f1a9a3c6-a3d3-44c6-aec8-14cd834ebaeb)
- [Office Developer Center](https://msdn.microsoft.com/en-us/office/aa905340.aspx)
- [Troubleshooting Office files and custom solutions with the telemetry log](https://msdn.microsoft.com/en-us/library/office/jj230106.aspx)
- [Office Application Compatibility Forum](http://social.technet.microsoft.com/forums/officesetupdeploy/threads)

