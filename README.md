<div align="center">

## A basic DDE Server that accepts LinkExecute commands


</div>

### Description

Demonstrates how to write a VB DDE Server that accepts DDEExecute command.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chui Tey](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chui-tey.md)
**Level**          |Beginner
**User Rating**    |4.8 (43 globes from 9 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[DDE](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/dde__1-28.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chui-tey-a-basic-dde-server-that-accepts-linkexecute-commands__1-36271/archive/master.zip)





### Source Code

<p><b>VB DDE Server</b>
<ul>
<li>Create new VB project and name it Project1.
<li>Set Form1.LinkTopic to "SYSTEM" and
<li>Set Form1.LinkMode to "1 - Source"
<li>Add the following code to the form
</ul>
 <code><pre>
Option Explicit
Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
 MsgBox "Received " & CmdStr, vbInformation
 Cancel = False
End Sub
</pre></code>
<p>Run the project
<p><b>VB DDE Client</b>
<p>We'll now code up a client that calls DDE Execute on the server.
<ul>
<li>Create new VB project and name it Project2.
<li>Add a textbox call "Text1"
<li>Add a command button called "Command1" and
<li>Add the following code to the form
</ul>
<code><pre>
Option Explicit
Private Sub Command1_Click()
 Text1.LinkMode = vbLinkNone
 Text1.LinkTopic = "Project1|SYSTEM"
 Text1.LinkMode = vbLinkManual
 Text1.LinkExecute "Hello World"
End Sub
</pre></code>
<p><b>Testing</b></p>
<p>Run the project.</p>
<p>When the form appears click on the button.
 Notice that the VB DDE Server has received
 a string from the client and shown it on a
 textbox.
</p>
<p><b>Comments</b>
<p>DDE is a mechanism that is supported by a lot of applications. It can only handle simple strings, and is not as flexible or powerful as COM. However, some older applications only support DDE, and it is still necessary to code
DDE Servers.
</p>
<p><b>Note on link topic</b></p>
<p>The link topic by default is the EXE name.
However, you can change this in the VBP file,
if you have a look at the key TITLE="...", this
key can be modified to use a different link topic.
</p>

