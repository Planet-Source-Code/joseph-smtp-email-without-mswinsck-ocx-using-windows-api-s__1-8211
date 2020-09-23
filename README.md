<div align="center">

## SMTP email without mswinsck\.ocx using Windows API's


</div>

### Description

SMTP email without mswinsck.ocx using Windows api's

Hai friends at psc...

This is the implementation of smtp (Simple Mail Transfer Protocol) using the winsock.dll. Till now i have seen only implentations of the above protocol using the bulky activex control mswinsck.ocx(106 kb). All you have to do is to include the module in your program and call the smtp( ...) function. It will return a 1 if succesful, 0 if an error occured during transmission and -1 if a serious error ocuured.

Your mail enabled program will be just about 30kb...Wonderful...No need of any support ocx files for your program to work. This program uses just pure windows api's only..Send in your comments and bug reports to josephninan@crosswinds.net.. The problem is that i have done extensive tests with the loopback ip address(127.0.0.1) running a mail server on my machine...Hadnt got the time to check out live at the net.

Also do mail me the latest version of the winsock.dll's basic module if you have...Thanks in advance
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2000-05-23 10:03:28
**By**             |[Joseph](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/joseph.md)
**Level**          |Advanced
**User Rating**    |4.8 (198 globes from 41 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD60205232000\.zip](https://github.com/Planet-Source-Code/joseph-smtp-email-without-mswinsck-ocx-using-windows-api-s__1-8211/archive/master.zip)








