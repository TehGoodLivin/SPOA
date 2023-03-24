# MIT License

Copyright (c) 2023 Austin L

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

# SPOA
A PnP PowerShell Tool used to assist with SharePoint Online Adminstive Actions.

## Install PnP PowerShell
The script will automatically run this for you and check for updates.

1.	Open “Windows PowerShell ISE”. You can type this in the windows search bar to find the exact program.

2.	Run the following command: Install-Module -Name PnP.PowerShell -Scope CurrentUser

3.	When prompted: "NuGet provider is required to continue PowerShell Get requires NuGet provider version '2.8.5.201' or newer to interact with NuGet-based repositories. The NuGet provider must be available in 'C:\Program Files\PackageManagement\ProviderAssemblies' or 'C:\Users\EDIPI\AppData\Local\PackageManagement\ProviderAssemblies'. You can also install the NuGet provider by running 'Install-PackageProvider –Name NuGet –Minimum Version 2.8.5.201 -Force'. Do you want PowerShell Get to install and import the NuGet provider now?

    [Y] Yes [N] No [S] Suspend [?] Help (default is "Y"):" 

    Select Yes (Y).

4.	When prompted: "Untrusted repository you are installing the modules from an untrusted repository. If you trust this repository, change its InstallationPolicy value by running the Set-PSRepository cmdlet. Are you sure you want to install the modules from 'PSGallery'?

    [Y] Yes [A] Yes to All [N] No [L] No to All [S] Suspend [?] Help (default is "N"):" 

    Select Yes to All (A).

5.	Verify Installation of PnP PowerShell 1.8.0 or higher: Get-Module PnP.PowerShell* -ListAvailable | Select-Object Name,Version | Sort-Object Version –Descending
