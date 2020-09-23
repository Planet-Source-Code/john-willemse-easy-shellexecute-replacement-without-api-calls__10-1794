<div align="center">

## Easy ShellExecute replacement without API calls


</div>

### Description

With only a few lines of code you will be able to open any document with the associated executable file or, if no association present, you will be prompted to open the file with the program you choose. This can also be used to execute an executable file. No API calls used!
 
### More Info
 
Simply pass the full path to the file in a string to function.

To use this function, simply put the code in a class and call it, e.g.:

Dim retVal As Boolean

retVal = ExecuteFile("c:\temp\myword.doc")

If retVal = False then

MsgBox("Something went wrong!")

Else

MsgBox("Succeeded!")

End If

That's all! Check the additional properties of a Process() for more options, like waiting for the process to finish, etc.

Returns true if succeeded, false in case of an error.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Willemse](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-willemse.md)
**Level**          |Beginner
**User Rating**    |4.8 (67 globes from 14 users)
**Compatibility**  |VB\.NET, ASP\.NET
**Category**       |[System Services/ Functions](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/system-services-functions__10-23.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-willemse-easy-shellexecute-replacement-without-api-calls__10-1794/archive/master.zip)

### API Declarations

Free to use, just vote for me ;)


### Source Code

```
Private Function ExecuteFile(ByVal FileName As String) As Boolean
 Dim myProcess As New Process()
 myProcess.StartInfo.FileName = FileName
 myProcess.StartInfo.UseShellExecute = True
 myProcess.StartInfo.RedirectStandardOutput = False
 myProcess.Start()
End Function
```

