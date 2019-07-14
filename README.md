# DocGuard Document Analyzer Solution Kit

## Overview
DocGuard started as an educational project that aims to teach the Windows User-Mode environment to newbie's which get in the cyber-security-job. The development idea of this project came along while inspect on [EvilClippy](https://github.com/outflanknl/EvilClippy) project and so I thought, "Why I don't try a solution that checks Office files for malicious attacks such as VBA Stomping, DDE or Macro?" 

Since the project was developed for educational purposes, I wanted to have a comprehensive solution that covers a lot of Office and Windows capabilities while design phase. Project consists of one main class and also six module which some of them are Add-Ons, Web and Windows Service. 

## Modules
These components are as follows;

* DocGuard Audit - This is the main component that serves all other projects. Its main task is to perform file checks against malicious attacks such as VBA Stomping, Malicious Macro and DDE Vulnerabilities.

* DocGuard Outlook - This [VSTO module](https://docs.microsoft.com/en-us/visualstudio/vsto/create-vsto-add-ins-for-office-by-using-visual-studio?view=vs-2019) work as an Outlook Add-in and enable then sends the attached file to DocGuard-Audit for analysis when a new message arrives.

* DocGuard Word - This [VSTO module](https://docs.microsoft.com/en-us/visualstudio/vsto/create-vsto-add-ins-for-office-by-using-visual-studio?view=vs-2019) work as an Word Add-in and enable then sends the opened file to DocGuard-Audit for analysis when open a word file. 

* DocGuard Excel - This [VSTO module](https://docs.microsoft.com/en-us/visualstudio/vsto/create-vsto-add-ins-for-office-by-using-visual-studio?view=vs-2019) work as an Excel Add-in and enable then sends the opened file to DocGuard-Audit for analysis when open a excel file. 

* DocGuard Service - This module work as a Windows Service and then sends to DocGuard-Audit for analysis when the read i/o process for supported office files on the file system. 

* DocGuard ShellExt - This module work as a Windows Shell Extension and then send to DocGuard-Audit which file that selected through right click menu.

* DocGuard WebApi - This module work as a Web Api and then send to DocGuard-Audit which file that uploaded through Browser, Postman, Fiddler etc...   