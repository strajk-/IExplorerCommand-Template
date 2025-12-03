## Overview
This repo contains the bare minimum required for creating your own Windows 11 context menu extension with IExplorerCommand without having to install any dependencies on target machines you intend to deploy this.

It is based on [xandfis's](https://github.com/xandfis) [W11ContextMenuDemo](https://github.com/xandfis/W11ContextMenuDemo) 

Everything is automated in the solution as Post-Build steps, all you have to do is compile the solution and it should output everything into a Debug or Release folder in the Solution's Root path.

It also creates a Self Signed Cert and signs all files in the root output directory, you can adjust the Post-Build step or Sign-Output.ps1 to a certificate you already have installed.

- CustomShellManager Application handles the installation and uninstall of both the Context Menu Extension and the Root CA on the target machine.

- CustomShell Library contains the implementation of the Extension.

- CustomShellPackage contains all assets and configuration files for the extension and MSIX package

You can learn more about the Windows 11 context menu on xandif's [post](https://blogs.windows.com/windowsdeveloper/2021/07/19/extending-the-context-menu-and-share-dialog-in-windows-11/) on the Windows Developer blog.

## Dependencies

* [Windows SDK Kit v10.0.20348.3330](https://go.microsoft.com/fwlink/?linkid=2331862), an SDK archive can be found [here](https://developer.microsoft.com/en-us/windows/downloads/sdk-archive/). You can use newer Versions, if you do so you will have to adjust the paths in the Post-Build Steps and Sign-Output.ps1 script to match what you have.
* [Visual Studio Build Tools 2019 (v142)](https://gist.github.com/Mr-Precise/9967e3fcf03f2df0282b76841d2f3876), newer Versions are compatible, just adjust it in the CustomShell Library if you end up using something else.
* [.NET Framework 4.8](https://dotnet.microsoft.com/en-us/download/dotnet-framework/net48), which is integrated into every Windows installation since Windows 10 v20H2, older Versions require it to be installed first, you could target 4.5 which is integrated since Windows 8.

## What is different from xandfis's demo
- No need to install .NET Core since it is build with .NET Framework 4.8.
- All in one Solution Build, so no need to manually copy files, create/apply certificates, edit configurations or manually create MSIX packages, everything is done automatically as long all dependencies are present.
- CustomShellManager Application that makes it very easy to handle deployment of your extension, run it with the /help parameter to see a list of commands you can use.
- CustomShellPackage folder has its resources.pri file re-created every Build with makepri.exe, to ensure that any changes to its structure is recognized by the MSIX Package.
- Minimal amount of files as resources, only shipping what is needed for the context menu, up to you if you want to add more, head over to the AppxManifest.xml and adjust it at your own leasure as you add more resources into the CustomShellPackage folder.
- CustomShell Library which contains the implementation of this extension, has some helper functions to help you start off with creating your own commands, such as:
	- GetExecutableIcon, extracts the icon being used on an executable you're trying to run as a command, which then shows up in the context menu, if none is found it defaults to what is defined in Resource.rc.
	- GetFirstSelectedFilePath, extracts the path of the file that was right clicked.
	- LaunchApplication, launches the application that you intend to run, it can either be the full path where the executable is located at, or just the application name if the folder it is located at is in the %PATH% Environmental Variable.
- Easy to apply your own branding, just use a tool such as AstroGrep, search for files and folders containing "CustomShell" as a name or content, rename and replace all instances it finds to what you want and you're good to go.

## Build steps 
In Visual Studio 2019 or newer if all dependencies are present:
 1. Clone repo
 2. Select Build Configuration (Debug/Release)
 3. Build IExplorerCommand-Template.sln
 4. Navigate to your Solution root directory, a Debug or Release folder should have been created with all necessary signed files to copy to your target machine.

## Installing Context Menu and Certificate
1. Run CustomShellManager.exe with /i parameter
Console output should tell you that it successfully installed both the certificate the executable is signed with and the context menu extension
## Uninstalling Context Menu and Certificate
1. Run CustomShellManager.exe with /u parameter to uninstall the extension
2. (Optional) Run CustomShellManager.exe with /uc parameter to uninstall the certificate
