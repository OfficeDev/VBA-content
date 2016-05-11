
# Configuring RDS on Windows 2000

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

If you experience difficulties getting RDS to function properly after upgrading to Windows 2000, follow the steps below to troubleshoot the issue.


1. Make sure that the World Wide Web Publishing Service is running first by navigating to http:// _server_ using Internet Explorer. If you are unable to access the web server this way, go to a command prompt and enter the following command, "NET START W3SVC".
    
2. From the Start menu, select Run. Type msdfmap.ini and click OK to open the msdfmap.ini file in Notepad. Check the [CONNECT DEFAULT] section, and if the ACCESS parameter is set to NOACCESS, change it to READONLY.
    
3. Using the RegEdit utility, navigate to "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\DataFactory\HandlerInfo" and make sure  **HandlerRequired** is set to 0 and **DefaultHandler** is "" (Null string).
    
     **Note**  If you make any changes to this section of the registry, you must stop and restart the World Wide Web Publishing Service by entering the following commands at a command prompt: "NET STOP W3SVC" and "NET START W3SVC".
4. Using the RegEdit utility, navigate in the registry to "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W3SVC\Parameters\ADCLaunch" and verify that there is a key called  **RDSServer.Datafactory**. If not, create it.
    
5. Using Internet Services Manager, go to the Default Web Site and view the properties of the MSADC virtual root. Inspect the Directory Security/IP Address and Domain Name Restrictions. If the "Access is Denied" is checked then select "Granted".
    
Be sure to try rebooting the server if the changes to do not appear to solve the problem at first.
