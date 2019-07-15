[![License: GNU ](https://img.shields.io/badge/License-GPL%20v3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0.html)
[![Release](https://img.shields.io/github/release-pre/steamer/docguard.svg)](https://github.com/SteAmeR/DocGuard/releases/tag/1.0)
[![Stars](https://img.shields.io/github/stars/SteAmeR/DocGuard.svg)]()
[![Issue](https://img.shields.io/github/issues/SteAmeR/DocGuard.svg)](https://github.com/SteAmeR/DocGuard/issues)

![Logo](https://raw.githubusercontent.com/SteAmeR/DocGuard/master/images/DocGuard_Logo_Small.png)

# DocGuard Document Analyzer Solution Kit

## Overview
DocGuard started as an educational project that aims to teach the Windows User-Mode environment to newbie's which get in the cyber-security-job. The development idea of this project came along while inspect on [EvilClippy](https://github.com/outflanknl/EvilClippy) project and so I thought, "Why I don't try a solution that checks Office files for malicious attacks such as VBA Stomping, DDE or Macro?" 

Since the project was developed for educational purposes, I wanted to have a comprehensive solution that covers a lot of Office and Windows capabilities while design phase. Project consists of one main class and also six module which some of them are Add-Ons, Web and Windows Service. 

I'm sure this project has dozens of bugs that need fixing. However for now they have to wait. Me and other participants (I hope :) will continue to add new features and bug fix to/for this project as long as we have the opportunity. 

[![DocGuard_Excel(https://raw.githubusercontent.com/SteAmeR/DocGuard/master/images/DocGuard_Excel_small.jpg)](https://raw.githubusercontent.com/SteAmeR/DocGuard/master/images/DocGuard_Excel_big.jpg)

## General features

* Supported files types: Doc, Docx, Docm, Dot, Xls, Xlsx, Xlsm 
* Detect Obfuscated VBA Code (Shannon Entropy)
* Detect DDE Vulnerabilities
* Detect Stomping VBA Code
* Detect Random Module Names
* Detect Blacklist Api Usage
* Detect Hide Module
* Detect Unviewable Protection

[![WebApi(https://raw.githubusercontent.com/SteAmeR/DocGuard/master/images/WebApi_small.jpg)](https://raw.githubusercontent.com/SteAmeR/DocGuard/master/images/WebApi_big.jpg)

## Modules
These components are as follows;

* DocGuard Audit - This is the main component that serves all other projects. Its main task is to perform file checks against malicious attacks such as VBA Stomping, Malicious Macro and DDE Vulnerabilities.

* DocGuard Outlook - This [VSTO module](https://docs.microsoft.com/en-us/visualstudio/vsto/create-vsto-add-ins-for-office-by-using-visual-studio?view=vs-2019) work as an Outlook Add-in and enable then sends the attached file to DocGuard-Audit for analysis when a new message arrives.

* DocGuard Word - This [VSTO module](https://docs.microsoft.com/en-us/visualstudio/vsto/create-vsto-add-ins-for-office-by-using-visual-studio?view=vs-2019) work as an Word Add-in and enable then sends the opened file to DocGuard-Audit for analysis when open a word file. 

* DocGuard Excel - This [VSTO module](https://docs.microsoft.com/en-us/visualstudio/vsto/create-vsto-add-ins-for-office-by-using-visual-studio?view=vs-2019) work as an Excel Add-in and enable then sends the opened file to DocGuard-Audit for analysis when open a excel file. 

* DocGuard Service - This module work as a Windows Service and then sends to DocGuard-Audit for analysis when the read i/o process for supported office files on the file system. 

* DocGuard ShellExt - This module work as a Windows Shell Extension and then send to DocGuard-Audit which file that selected through right click menu.

* DocGuard WebApi - This module work as a Web Api and then send to DocGuard-Audit which file that uploaded through Browser, Postman, Fiddler etc...   

## To-Do

* DocGuard Scanner

* DocGuard FS Filter-Driver

* PDF, EXE, Php, Asp, Jsp support for another malicious file base attack types

## Third Party Credits

* [OpenMCDF](https://github.com/ironfede/openmcdf) to play with OLE components on DocGuard_Audit
* [Kavod](https://github.com/rossknudsen/Kavod.Vba.Compression) compress/decompress stuffs on VBA things
* [SharpShell](https://github.com/dwmkerr/sharpshell) ContextMenuHandler for DocGuard_ShellExt
* [StrongNamer](https://github.com/dsplaisted/strongnamer) build helpers for DocGuard_ShellExt

