<div align="center">

## Dealing with C DLLs


</div>

### Description

Ever stumbled upon the error "Bad DLL Calling Convention"? Was it when you tried to declare a function contained in a DLL? Then this might be a solution for you.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2006-12-14 16:30:02
**By**             |[Arne Elster](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/arne-elster.md)
**Level**          |Advanced
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Dealing\_wi20376712142006\.zip](https://github.com/Planet-Source-Code/arne-elster-dealing-with-c-dlls__1-67384/archive/master.zip)





### Source Code

<style type="text/css">
<!--
.VBCodeFont {
	font-family: "Courier New", Courier, monospace;
	font-size: 12px;
}
.VBKeyWordColor {color: #0033FF}
-->
</style>
<h2>Calling Conventions and VB </h2>
<p>Suppose you have written some functions in C/C++ and you exported them to a DLL.<br>
 If not explicitly changed, they will have the calling convetion "cdecl".
A calling convention specifies in which order arguments <br>
are pushed on the stack and who removes them from it after the called function returns, either the caller or the callee.<br>
VB only supports the calling convention "stdcall" (actually, cdecl too, we will get back to that), so how would you call these functions,<br>
if they aren't compatible to VB?<br>
If you have the source code to the C functions, you change the calling convention, look into the manual of your compiler, or search for<br>
something like &quot;stdcall C++&quot;.<br>
But what to do if you do not? In this case you have various options.</p>
<ol>
 <li> Write a Wrapper DLL in C/C++, where each function has the calling convention stdcall and internally calls the cdecl function. </li>
 <li> Write a Wrapper DLL in VB! </li>
 <li> Use some precompiled machine code to call the functions.<br>
 </li>
</ol>
<h3>1. Writing a Wrapper DLL in C/C++ </h3>
<p>This is a trivial task if you know some C.</p>
<h3>2. Writing a Wrapper DLL in VB</h3>
<p>Yes, this is actually possible. VB also supports cdecl!<br>
 But only compiled to native code.<br>
 You can create a new Active-X DLL with an empty class, write your declares, and add some functions which call these functions.<br>
 For example:</p>
<p class="VBCodeFont"><span class="VBKeyWordColor">Option Explicit</span><br>
 <br>
 <span class="VBKeyWordColor">Private Declare Function</span> my_c_function <span class="VBKeyWordColor">Lib</span> &quot;my_c.dll&quot; (<span class="VBKeyWordColor">ByVal</span> param1 <span class="VBKeyWordColor">As Long</span>, <span class="VBKeyWordColor">ByVal</span> param2 <span class="VBKeyWordColor">As Long</span>)<span class="VBKeyWordColor"> As Long</span><br>
 <br>
 <span class="VBKeyWordColor">Public Function</span> my_c_function_wrapper<span class="VBKeyWordColor">(ByVal</span> param1 <span class="VBKeyWordColor">As Long</span>, <span class="VBKeyWordColor">ByVal</span> param2 <span class="VBKeyWordColor">As Long) As Long</span><br>
&nbsp;&nbsp;&nbsp;&nbsp;my_c_function_wrapper = my_c_function(param1, param2)<br>
<span class="VBKeyWordColor">End Function</span></p>
<p>After you're finished, compile the DLL (make sure you compile to Native Code, as P-Code won't work!) and reference it <br>
 in a new Standard Exe project. Now you can call the functions without any problems.<br>
 But beware! VB does not support cdecl callbacks!<br>
This can be solved with:</p>
<h3>3. Assembler for cdecl calls and callbacks </h3>
<p>The way of of including machine code I used is pretty easy. You reserve executable memory with VirtualAlloc,<br>
 write your machine code into it, and call it using some function which allows the execution of a function pointer,<br>
 e.g. CallWindowProc().<br>
The bad thing about this way is, you have to know how your arguments look on the stack.<br>
The good thing is, you don't need additional components.<br>
Anyways, you can also combine my function CreateCdeclCbWrap() with method #2 (writing a Wrapper DLL in VB)<br>
and just ignore CallCdecl(). For some code look at the attachment (the example shows how to use the function &quot;qsort&quot;<br>
of the Microsoft Visual C Runtime (msvcrt.dll)).
</p>

