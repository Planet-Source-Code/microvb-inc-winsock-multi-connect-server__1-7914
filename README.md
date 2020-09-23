<div align="center">

## Winsock \- Multi\-Connect Server


</div>

### Description

A full basic winsock server allowing for 1000 and more client connections. Very easy to expand on for any type of server that needs the ability to allow multiple socket connections.
 
### More Info
 
Run the 'Server' and then run as many 'Client' as you want. The socket client that is included connects to the server on launch. If you read the code, it is well commented, and not very complex. This makes winsock easy and fun. =) .... Enjoy.



----

SERVER: I tried to trap for Error 340, but somehow, it ignores it sometimes... (you will understand when you run it)... all this error means is that the control is not available... I unload them as they are no longer needed to free up memory, and use error handling to try and trap it, but it would seem this is another microsoft bug.

----

CLIENT: Further, on the client side, I tried to add some code to the unload event that would send a disconnect request to the server prior to unloading if the socket was still connected, It would seem that this too is sadly ignored.

----

= Any input on this matter would be appreciated.=

----




<span>             |<span>
---                |---
**Submitted On**   |2000-05-07 06:38:12
**By**             |[MicroVB INC](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/microvb-inc.md)
**Level**          |Advanced
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD5560572000\.zip](https://github.com/Planet-Source-Code/microvb-inc-winsock-multi-connect-server__1-7914/archive/master.zip)








