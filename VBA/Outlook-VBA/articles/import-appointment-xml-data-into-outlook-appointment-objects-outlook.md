---
title: Import Appointment XML Data into Outlook Appointment Objects (Outlook)
ms.prod: outlook
ms.assetid: ecfd3849-877b-01ad-2b76-1a54e980f6e2
ms.date: 06/08/2017
---


# Import Appointment XML Data into Outlook Appointment Objects (Outlook)

This topic shows how to read appointment data formatted in XML, save the data to Microsoft Outlook **[AppointmentItem](appointmentitem-object-outlook.md)** objects in the default calendar, and return the appointment objects in an array.



|
![MVP logo](./images/MVPLogo_Small_ZA10349011.jpg)

|Helmut Obertanner provided the following code samples. Helmut is a  [Microsoft Most Valuable Professional](https://mvp.microsoft.com/en-us/default.aspx
) with expertise in Microsoft Office development tools in Microsoft Visual Studio and Microsoft Office Outlook.|



The following managed code samples are written in C# and Visual Basic. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code samples in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.
The following code samples contain the  `CreateAppointmentsFromXml` method of the `Sample` class, implemented as part of an Outlook add-in project. Each project adds a reference to the Outlook PIA, which is based on the **Microsoft.Office.Interop.Outlook** namespace.
The  `CreateAppointmentsFromXml` method accepts two input parameters, _application_ and _xml_:

-  _application_ is a trusted Outlook **[Application](application-object-outlook.md)** object.
    
-  _xml_ is either an XML string, or a string that represents a path to a valid XML file. For the purpose of the following code samples, the XML delimits appointment data by using the following XML tags:
    

|**Appointment data**|**Delimiting XML tag**|
|:-----|:-----|
|Entire set of appointment data|<appointments>|
|Each appointment in the set|<appointment>|
|Start time of an appointment|<starttime>|
|End time of an appointment|<endtime>|
|Title of an appointment|<subject>|
|Location of an appointment|<location>|
|Details of an appointment|<body>|

The following example shows input data for the  _xml_ parameter.



```XML
<?xml version="1.0" encoding="utf-8" ?>  
<appointments> 
    <appointment> 
        <starttime>2009-06-01T15:00:00</starttime> 
        <endtime>2009-06-01T16:15:00</endtime> 
        <subject>This is a Test-Appointment</subject> 
        <location>At your Desk</location> 
        <body>Here is the Bodytext</body> 
    </appointment> 
    <appointment> 
        <starttime>2009-06-01T17:00:00</starttime> 
        <endtime>2009-06-01T17:15:00</endtime> 
        <subject>This is a second Test-Appointment</subject> 
        <location>At your Desk</location> 
        <body>Here is the Bodytext</body> 
    </appointment> 
    <appointment> 
        <starttime>2009-06-01T17:00:00</starttime> 
        <endtime>2009-06-01T18:15:00</endtime> 
        <subject>This is a third Test-Appointment</subject> 
        <location>At your Desk</location> 
        <body>Here is the Bodytext</body> 
    </appointment> 
</appointments> 

```

 The `CreateAppointmentsFromXml` method uses the Microsoft COM implementation of the XML Document Object Model (DOM) to load and process the XML data that _xml_ provides. `CreateAppointmentsFromXml` first checks whether _xml_ specifies a valid source of XML data. If so, it loads the data into an XML document, **DOMDocument**. Otherwise,  `CreateAppointmentsFromXml` throws an exception. For more information about the XML DOM, see [DOM](http://msdn.microsoft.com/library/e9da2722-7879-4e48-869c-7f16714e2824%28Office.15%29.aspx).
For each appointment child node delimited by the <appointment> tag in the XML data,  `CreateAppointmentsFromXml` looks for specific tags, uses the DOM to extract the data, and assigns the data to corresponding properties of an **AppointmentItem** object: **[Start](appointmentitem-start-property-outlook.md)**,  **[End](appointmentitem-end-property-outlook.md)**,  **[Subject](appointmentitem-subject-property-outlook.md)**,  **[Location](appointmentitem-location-property-outlook.md)**, and  **[Body](appointmentitem-body-property-outlook.md)**.  `CreateAppointmentsFromXml` then saves the appointment to the default calendar.
 `CreateAppointmentsFromXml` uses the ** [Add](http://msdn.microsoft.com/library/frlrfSystemCollectionsGenericList1ClassAddTopic%28Office.15%29.aspx)** method of the **List( _type_)** class in the **System.Collections.Generic** namespace to aggregate these **AppointmentItem** objects. When the method has processed all the appointments in the XML data, it returns the **AppointmentItem** objects in an array.
The following is the C# code sample.



```C#
using System; 
using System.Collections.Generic; 
using System.IO; 
using System.Text; 
using System.Xml; 
using Outlook = Microsoft.Office.Interop.Outlook; 
 
namespace OutlookAddIn1 
{ 
    class Sample 
    { 
        Outlook.AppointmentItem[] CreateAppointmentsFromXml(Outlook.Application application,  
                                                            string xml) 
        { 
            // Create a list of appointment objects. 
            List<Outlook.AppointmentItem> appointments = new  
                List<Microsoft.Office.Interop.Outlook.AppointmentItem>(); 
            XmlDocument xmlDoc = new XmlDocument(); 
 
            // If xml is an XML string, create the document directly.  
            if (xml.StartsWith("<?xml")) 
            { 
                xmlDoc.LoadXml(xml); 
            } 
            else if (File.Exists(xml)) 
            { 
                xmlDoc.Load(xml); 
            } 
            else 
            { 
                throw new Exception( 
                    "The input string is not valid XML data or the specified file doesn't exist."); 
            } 
 
            // Select all appointment nodes under the root appointements node. 
            XmlNodeList appointmentNodes = xmlDoc.SelectNodes("appointments/appointment"); 
            foreach (XmlNode appointmentNode in appointmentNodes) 
            { 
 
                // Create a new AppointmentItem object. 
                Outlook.AppointmentItem newAppointment =  
                    (Outlook.AppointmentItem)application.CreateItem(Outlook.OlItemType.olAppointmentItem); 
 
                // Loop over all child nodes, check the node name, and import the data into the  
                // appointment fields. 
                foreach (XmlNode node in appointmentNode.ChildNodes) 
                { 
                    switch (node.Name) 
                    { 
 
                        case "starttime": 
                            newAppointment.Start = DateTime.Parse(node.InnerText); 
                            break; 
 
                        case "endtime": 
                            newAppointment.End = DateTime.Parse(node.InnerText); 
                            break; 
 
                        case "subject": 
                            newAppointment.Subject = node.InnerText; 
                            break; 
 
                        case "location": 
                            newAppointment.Location = node.InnerText; 
                            break; 
 
                        case "body": 
                            newAppointment.Body = node.InnerText; 
                            break; 
 
                    } 
                } 
 
                // Save the item in the default calendar. 
                newAppointment.Save(); 
                appointments.Add(newAppointment); 
            } 
 
            // Return an array of new appointments. 
            return appointments.ToArray(); 
        } 
 
    } 
}
```

The following is the Visual Basic code sample.



```VB.net
Imports System.IO 
Imports System.Xml 
Imports Outlook = Microsoft.Office.Interop.Outlook 
 
Namespace OutlookAddIn2 
    Class Sample 
        Function CreateAppointmentsFromXml(ByVal application As Outlook.Application, _ 
            ByVal xml As String) As Outlook.AppointmentItem() 
 
            Dim appointments As New List(Of Outlook.AppointmentItem) 
            Dim xmlDoc As New XmlDocument() 
 
            If xml is an XML string, create the XML document directly. 
            If xml.StartsWith("<?xml") Then 
                xmlDoc.LoadXml(xml) 
            ElseIf (File.Exists(xml)) Then 
                xmlDoc.Load(xml) 
            Else 
                Throw New Exception("The input string is not valid XML data or the specified file doesn't exist.") 
            End If 
 
 
            ' Select all appointment nodes under the root appointements node. 
            Dim appointmentNodes As XmlNodeList = xmlDoc.SelectNodes("appointments/appointment") 
 
            For Each appointmentNode As XmlNode In appointmentNodes 
 
                ' Create a new AppointmentItem object. 
                Dim newAppointment As Outlook.AppointmentItem = _ 
                    DirectCast(application.CreateItem(Outlook.OlItemType.olAppointmentItem), _ 
                    Outlook.AppointmentItem) 
 
                ' Loop over all child nodes, check the node name, and import the data into the appointment fields. 
 
                For Each node As XmlNode In appointmentNode.ChildNodes 
                    Select Case (node.Name) 
 
                        Case "starttime" 
                            newAppointment.Start = DateTime.Parse(node.InnerText) 
 
 
                        Case "endtime" 
                            newAppointment.End = DateTime.Parse(node.InnerText) 
 
 
                        Case "subject" 
                            newAppointment.Subject = node.InnerText 
 
 
                        Case "location" 
                            newAppointment.Location = node.InnerText 
 
 
                        Case "body" 
                            newAppointment.Body = node.InnerText 
 
 
                    End Select 
                Next 
 
                ' Save the item in the default calendar. 
                newAppointment.Save() 
                appointments.Add(newAppointment) 
            Next 
 
            ' Return an array of new appointments. 
            Return appointments.ToArray() 
        End Function 
 
 
    End Class 
End Namespace
```


