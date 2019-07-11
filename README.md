# Citrix Provisioning Services (PVS) Health Check Script

_(The following was copied from the PVS Assessment Script ReadMe RTF)_

Creates a basic health check of a Citrix PVS 5.x, 6.x or 7.x farm.

## Prerequisites

### SQL Permissions

The account used to run this script must have at least Read access to the SQL Server that holds the Citrix Provisioning databases.

### PVS Console PowerShell Snap-in

Before we can start using PowerShell to assess a PVS farm, let us ensure we have the necessary requirements. The following code samples show how to install the (DLL-based) PowerShell Snap-in particular to various versions of Windows Operating System

For versions of Windows prior to Windows 8 and Server 2012, run:

- For 32-bit:

```PowerShell
%systemroot%\Microsoft.NET\Framework\v2.0.50727\installutil.exe "%ProgramFiles%\Citrix\Provisioning Services Console\McliPSSnapIn.dll"
```

- For 64-bit:

```PowerShell
%systemroot%\Microsoft.NET\Framework64\v2.0.50727\installutil.exe "%ProgramFiles%\Citrix\Provisioning Services Console\McliPSSnapIn.dll"
```

For Windows 8.1 and later, Server 2012 and later, run:

- For 32-bit:

```PowerShell
%systemroot%\Microsoft.NET\Framework\v4.0.30319\installutil.exe "%ProgramFiles%\Citrix\Provisioning Services Console\McliPSSnapIn.dll"
```

- For 64-bit:

```PowerShell
%systemroot%\Microsoft.NET\Framework64\v4.0.30319\installutil.exe "%ProgramFiles%\Citrix\Provisioning Services Console\McliPSSnapIn.dll"
```

_Note_: All code snippet lines above are intended to be run as one line. If you copy and paste into a command prompt, watch for the smart quotes. If you get an “invalid URL” error, paste the text into Notepad and change the smart quotes to regular double quotes.

_Note_: If you are running 64-bit Windows Server, you will need to run both commands so the snap-in is registered for both 32-bit and 64-bit PowerShell.

### Script Usage

How to use this script?

1. Save the script as PVS_Assessment.ps1 in your PowerShell scripts folder.
2. From a PowerShell prompt, change to your PowerShell scripts folder (where the script was saved to in step 1).
3. From the PowerShell prompt, run the script.

    ```PowerShell
    .\PVS_Assessment.ps1
    ```

4. A text file is created named for the PVS Farm.
5. This script is designed to be run directly on a PVS server since only a text file is created.
6. Word is not used or required.

Full help text is available.

```PowerShell
Get-Help .\PVS_Assessment.ps1 –full
```

See also: [PowerShell_Help](PowerShell_Help.md)

### Supported Environments

I have tested this script with PVS 5.6 SP3 through 1903. I have tested the script running on and from Windows 7, Windows 8.x, Windows 10, Server 2003 R2, Server 2008 R2, Server 2012, Server 2012 R2, and Server 2016.

To run the script remotely, please read the following articles:

- http://carlwebster.com/using-my-citrix-pvs-powershell-documentation-script-with-remoting/

- http://carlwebster.com/pvs-v2-documentation-script-has-been-updated-17-jun-2013/

- http://carlwebster.com/error-in-the-provisioning-services-7-0-powershell-programmer-guide-for-windows-8-and-server-2012/
