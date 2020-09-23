<div align="center">

## Proper Way To Use WinXP Manifest


</div>

### Description

After attempting to use various manifest creators and uilities found here and finding out it either doesn't work at all or works but with errors I decided to add minor fixes and explain exactly how to implement windows xp theme use and not get an error... Well I hope you don't get an error cause mine doesn't generate errors. Now don't get me wrong there are a few manifest programs I found that worked fine in displaying the content but I was getting various errors from them when using other components in the program and this seems to have worked with most of the components I have.

I didn't create this coding but I did figure out how to avoid errors, at least on my own machine and if this helps you please toss a vote just for the sake of all this typing and rambling. *lol*

Additional Note I Left Out: You do not need to add Reunion.exe for your app file name it can be just reunion without the exe (of course reunion probably iisn't the name of your exe but mine is so just for ease of the explanation)

A piece of advice when doing xp theming, if you place the controls within a frame they will form a ugly black outline so make your controls on the form not inside of a frame then move the controls to the frame.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[PorkNBeans](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/porknbeans.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/porknbeans-proper-way-to-use-winxp-manifest__1-29985/archive/master.zip)





### Source Code

In your form add the following:
<BR>
Option Explicit
<br>
Private Declare Function InitCommonControls
Lib "Comctl32.dll" () As Long
<p>
Private Sub Form_Initialize()
<br>
Dim XP As Long
<br>
XP = InitCommonControls
<br>
<br>
End Sub
<p>
Now in a good text editor geared towards coding languages (such as edit plus) you add this.
<br>
#?xml version="1.0" encoding="UTF-8" standalone="yes"?#
<br>
#assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0"#
<br>
#assemblyIdentity
version="1.0.0.0"
processorArchitecture="X86"
name="PorkORamaInc.ClassReunionCompanion.Reunion.exe"
type="win32"
/#
<br>
#description#WindowsExecutable#/description# "
<br>
#dependency#
<br>
#dependentAssembly#
<br>
#assemblyIdentity
type="win32"
name="Microsoft.Windows.Common-Controls"
version="6.0.0.0"
processorArchitecture="X86"
publicKeyToken="6595b64144ccf1df"
language="*"
/#
<br>
 #/dependentAssembly#
<br>
#/dependency#
<br>
#/assembly#
<p>
Now for some explanations...
this line here :<br>
name="PorkORamaInc.ClassReunionCompanion.Reunion.exe"
<br>
PorkORamaInc would be your company name
ClassReunionCompanion Would Be Your Product Name
Reunion.exe of course would be your programs exe name
<p>
This line here:
#description#WindowsExecutable#/description#
well thats basically what it appears to be, it's your files description.
when done save the file as the following
Reunion.exe.manifest
<br>
reunion.exe would be replaced by the actual name of your programs executable file.
Due to psc treating the code as html i had to replace the < and > with # so just replace them when using this code.

