---
title: PowerShell Gallery - Download
sidebar: mydoc_sidebar
permalink: mydoc_psgallery_download.html
folder: mydoc
---

## Download from PowerShell Gallery - Offline Install

The DSCEA PowerShell module needs to be installed only on the server acting as the management system that will be used to connect to remote systems to evaluate their compliance.

If the server acting as your management system does not have internet access, you can download the DSCEA module from the PowerShell Gallery using an internet connected system and copy the module files over to your management system.

Details on downloading DSCEA from the PowerShell Gallery can be found at the link below:

[https://www.powershellgallery.com/packages/DSCEA](https://www.powershellgallery.com/packages/DSCEA)

To download the module, use the Save-Module function and copy the module files to C:\Program Files\WindowsPowerShell\Modules or any other location defined in the PSModulePath variable on the management system.