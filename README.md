<div align="center">

## Beginners guide \[2\]  to ASP filled with examples\!


</div>

### Description

Here is my second tutorial in basic ASP scripting. Leave comments if you want. Or have questions.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[snowboardr](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/snowboardr.md)
**Level**          |Beginner
**User Rating**    |4.3 (43 globes from 10 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__4-33.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/snowboardr-beginners-guide-2-to-asp-filled-with-examples__4-7324/archive/master.zip)





### Source Code


<html>
<head>
<title>Jason's Tutorial..</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#CCCCCC" text="#000000">
<table width="43%" border="0" cellspacing="0" cellpadding="0" height="25" align="center">
 <tr>
  <td height="13"><font face="Verdana, Arial, Helvetica, sans-serif"><b>ASP
   - Active Server Pages [<i>Tutorial 2</i>] - <a href="http://www.irideforlife.com/go?x=view">Tutorial
   1</a></b></font></td>
 </tr>
</table>
<table width="84%" border="1" cellspacing="0" cellpadding="0" height="94" align="center" bordercolor="#999999" bgcolor="#CCCCCC">
 <tr>
  <td bgcolor="#CCCCCC" height="23"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Getting
   Scriptname </b></font></td>
 </tr>
 <tr>
  <td bgcolor="#eeeeee" height="69">
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&lt;%</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#009933">'This
    is very useful, using this in your code will allow you to change the name
    of your page as much as you want without having to chang any source</font></p>
   <p> <font size="2" face="Verdana, Arial, Helvetica, sans-serif">ScriptName
    = Request.ServerVariables(&quot;SCRIPT_NAME&quot;)</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#993333">Response.write</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">(ScriptName)</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">%&gt;</font></p>
  </td>
 </tr>
 <tr>
  <td bgcolor="#FFFFFF"> default.asp</td>
 </tr>
 <tr>
  <td bgcolor="#FFFFFF">Live Example at <a href="http://www.irideforlife.com/go/default2.asp" target="_blank">http://www.irideforlife.com/go/default2.asp</a></td>
 </tr>
</table>
<p>&nbsp;</p>
<table width="83%" border="1" cellspacing="0" cellpadding="0" height="107" align="center" bordercolor="#999999" bgcolor="#CCCCCC">
 <tr>
  <td bgcolor="#CCCCCC" height="23"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Redirect
   to url..</b></font></td>
 </tr>
 <tr>
  <td bgcolor="#eeeeee">
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&lt;%</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Response.Redirect
    &quot;http://www.planetsourcecode.com&quot;</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">%&gt;</font></p>
  </td>
 </tr>
</table>
<p>&nbsp;</p>
<table width="84%" border="1" cellspacing="0" cellpadding="0" height="107" align="center" bordercolor="#999999" bgcolor="#CCCCCC">
 <tr>
  <td bgcolor="#CCCCCC" height="23"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Stop
   writing the page..</b></font></td>
 </tr>
 <tr>
  <td bgcolor="#eeeeee">
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&lt;%</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><br>
    </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#993333">Response.write</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
    &quot;Hello there, this page needs to...&quot;</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Response.End<br>
    </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#993333">Response.write</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
    &quot;Have an ending?&quot;</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">%&gt;</font></p>
  </td>
 </tr>
 <tr>
  <td bgcolor="#eeeeee"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Hello
   there, this page needs to...</font> </td>
 </tr>
 <tr>
  <td bgcolor="#eeeeee">Live Example at <a href="http://www.irideforlife.com/go/default2.asp" target="_blank">http://www.irideforlife.com/go/default2.asp</a></td>
 </tr>
</table>
<p>&nbsp;</p>
<table width="82%" border="1" cellspacing="0" cellpadding="0" height="94" align="center" bordercolor="#999999" bgcolor="#CCCCCC">
 <tr>
  <td bgcolor="#CCCCCC" height="23"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Count
   the number of chr's in your string</b></font></td>
 </tr>
 <tr>
  <td bgcolor="#eeeeee" height="69">
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&lt;%</font></p>
   <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif">MyString
    = &quot;1234567&quot;</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><br>
    </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#993333">Response.write</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
    len(MyString)</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">%&gt;</font></p>
  </td>
 </tr>
 <tr>
  <td bgcolor="#FFFFFF"> 7</td>
 </tr>
 <tr>
  <td bgcolor="#FFFFFF">Live Example at <a href="http://www.irideforlife.com/go/default2.asp" target="_blank">http://www.irideforlife.com/go/default2.asp</a></td>
 </tr>
</table>
<p>&nbsp;</p>
<table width="80%" border="1" cellspacing="0" cellpadding="0" height="94" align="center" bordercolor="#999999" bgcolor="#CCCCCC">
 <tr>
  <td bgcolor="#CCCCCC" height="23"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Select
   Case </b></font></td>
 </tr>
 <tr>
  <td bgcolor="#eeeeee" height="345">
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&lt;%</font></p>
   <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif">MyCase =
    &quot;3&quot;</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Select <font color="#0000FF">Case</font>
    MyCase </font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#0000FF">Case</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
    1</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#993333">Response.write</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
    &quot;its 1&quot;</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#0000FF">Case</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
    2</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#993333">Response.write</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
    &quot;its 2&quot;</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#0000FF">Case</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
    3</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#993333">Response.write</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
    &quot;its 3&quot;</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#0000FF">End</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
    <font color="#0000FF">Select</font></font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">%&gt;</font></p>
  </td>
 </tr>
 <tr>
  <td bgcolor="#FFFFFF">
   <p>Live Example at <a href="http://www.irideforlife.com/go/default2.asp" target="_blank">http://www.irideforlife.com/go/default2.asp</a></p>
   </td>
 </tr>
</table>
<p>&nbsp;</p>
<table width="80%" border="1" cellspacing="0" cellpadding="0" height="107" align="center" bordercolor="#999999" bgcolor="#CCCCCC">
 <tr>
  <td bgcolor="#CCCCCC" height="23"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><a name="case"></a>Select
   Case QueryString</b></font></td>
 </tr>
 <tr>
  <td bgcolor="#eeeeee">
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&lt;%</font></p>
   <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#006633">'&lt;%=
    ScriptName &amp; &quot;?case=one&quot; %&gt;<br>
    '&lt;%= ScriptName &amp; &quot;?case=two&quot; %&gt;<br>
    '&lt;%= ScriptName &amp; &quot;?case=three&quot;%&gt;</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><br>
    </font></p>
   <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif">TheCase =
    Request.QueryString(&quot;case&quot;)</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Select <font color="#0000FF">Case</font>
    TheCase</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#0000FF">Case</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
    &quot;one&quot;</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#993333">Response.write</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
    &quot;its 1&quot; </font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#0000FF">Case</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
    &quot;two&quot;</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#993333">Response.write</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
    &quot;its 2&quot;</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#0000FF">Case</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
    &quot;three&quot;</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#993333">Response.write</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
    &quot;its 3&quot;</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#0000FF">End</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
    <font color="#0000FF">Select</font></font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">%&gt;</font></p>
  </td>
 </tr>
 <tr>
  <td bgcolor="#FFFFFF"> Live Example at <a href="http://www.irideforlife.com/go/default2.asp" target="_blank">http://www.irideforlife.com/go/default2.asp</a></td>
 </tr>
</table>
<p>&nbsp;</p>
<table width="82%" border="1" cellspacing="0" cellpadding="0" height="107" align="center" bordercolor="#999999" bgcolor="#CCCCCC">
 <tr>
  <td bgcolor="#CCCCCC" height="23"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Writing
   m / dd/yyyy</b></font></td>
 </tr>
 <tr>
  <td bgcolor="#eeeeee">
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&lt;%</font></p>
   <p><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">TheMonth
    = month(date)</font></p>
   <p><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">TheDay
    = day(date)</font></p>
   <p><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">TheYear
    = year(date)</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#993333">Response.write</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">(TheMonth)
    &amp; &quot;/&quot; &amp; TheDay &amp; &quot;/&quot; &amp; TheYear</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#339933">'
    &amp; Seperates each thing.</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">%&gt;</font></p>
  </td>
 </tr>
 <tr>
  <td bgcolor="#FFFFFF"> Live Example at <a href="http://www.irideforlife.com/go/default2.asp" target="_blank">http://www.irideforlife.com/go/default2.asp</a></td>
 </tr>
</table>
<p>&nbsp;</p>
<table width="84%" border="1" cellspacing="0" cellpadding="0" height="107" align="center" bordercolor="#999999" bgcolor="#CCCCCC">
 <tr>
  <td bgcolor="#CCCCCC" height="23"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Writing
   HH:MM:SS</b></font></td>
 </tr>
 <tr>
  <td bgcolor="#eeeeee">
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&lt;%</font></p>
   <p><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">TheHour
    = hour(time)</font></p>
   <p><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">TheMinute
    = minute(time)</font></p>
   <p><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">TheSecond
    = second(time)</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#993333">Response.write</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">(TheHour)
    &amp; &quot;:&quot; &amp; TheMinute &amp; &quot;:&quot; &amp; TheSecond</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#339933">'
    &amp; Seperates each thing.</font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">%&gt;</font></p>
  </td>
 </tr>
 <tr>
  <td bgcolor="#FFFFFF"> Live Example at <a href="http://www.irideforlife.com/go/default2.asp" target="_blank">http://www.irideforlife.com/go/default2.asp</a></td>
 </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
<table width="87%" border="1" cellspacing="0" cellpadding="0" height="107" align="center" bordercolor="#999999" bgcolor="#CCCCCC">
 <tr>
  <td bgcolor="#CCCCCC" height="23"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><a name="age"></a>Getting
   Form Results</b></font></td>
 </tr>
 <tr>
  <td bgcolor="#eeeeee">
   <p><font size="2">&lt;% </font></p>
   <p><font size="2">Dim <b>age</b></font></p>
   <p><font size="2"><b>age</b> = Request.Form(&quot;<b>age</b>&quot;)</font></p>
   <p><font size="2">%&gt;</font></p>
   <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&lt;form
    name=&quot;form1&quot; method=&quot;post&quot; action=&quot;&lt;%=ScriptName
    &amp; &quot;?submit=form&quot; %&gt;&quot;&gt;<br>
    &lt;table width=&quot;9%&quot; border=&quot;0&quot; cellspacing=&quot;0&quot;
    cellpadding=&quot;0&quot;&gt;<br>
    &lt;tr&gt; <br>
    &lt;td&gt;&lt;font size=&quot;2&quot; face=&quot;Verdana, Arial, Helvetica,
    sans-serif&quot;&gt;&amp;lt;18&lt;/font&gt;&lt;/td&gt;<br>
    &lt;td&gt;&lt;font size=&quot;2&quot; face=&quot;Verdana, Arial, Helvetica,
    sans-serif&quot;&gt; <br>
    &lt;input type=&quot;radio&quot; name=&quot;<b>age</b>&quot; value=&quot;<b>&lt;18</b>&quot;
    <font color="#FF0000">&lt;% if <b>age</b> = &quot;&lt;18&quot; then <font color="#993300">Response.write</font>
    &quot;checked&quot; %&gt;</font> &gt;<br>
    &lt;/font&gt;&lt;/td&gt;<br>
    &lt;/tr&gt;<br>
    &lt;tr&gt; <br>
    &lt;td&gt;&lt;font size=&quot;2&quot; face=&quot;Verdana, Arial, Helvetica,
    sans-serif&quot;&gt;18-21&lt;/font&gt;&lt;/td&gt;<br>
    &lt;td&gt;&lt;font size=&quot;2&quot; face=&quot;Verdana, Arial, Helvetica,
    sans-serif&quot;&gt; <br>
    &lt;input type=&quot;radio&quot; name=&quot;<b>age</b>&quot; value=&quot;<b>18-21</b>&quot;
    <font color="#FF0000">&lt;% if <b>age</b> = &quot;18-21&quot; then Response.write
    &quot;checked&quot; %&gt;</font> &gt;<br>
    &lt;/font&gt;&lt;/td&gt;<br>
    &lt;/tr&gt;<br>
    &lt;tr&gt; <br>
    &lt;td&gt;&lt;font face=&quot;Verdana, Arial, Helvetica, sans-serif&quot;
    size=&quot;2&quot;&gt;22 &amp;gt;&lt;/font&gt;&lt;/td&gt;<br>
    &lt;td&gt;&lt;font size=&quot;2&quot; face=&quot;Verdana, Arial, Helvetica,
    sans-serif&quot;&gt; <br>
    &lt;input type=&quot;radio&quot; name=&quot;<b>age</b>&quot; value=&quot;<b>22&gt;</b>&quot;
    <font color="#FF0000">&lt;% if <b>age</b> = &quot;22&gt;&quot; then Response.write
    &quot;checked&quot; %&gt;</font> &gt;<br>
    &lt;/font&gt;&lt;/td&gt;<br>
    &lt;/tr&gt;<br>
    &lt;tr&gt; <br>
    &lt;td colspan=&quot;2&quot;&gt;<br>
    &lt;input type=&quot;submit&quot; name=&quot;Submit&quot; value=&quot;Submit&quot;&gt;<br>
    <font color="#FF0000">&lt;%</font></font></p>
   <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FF0000">If
    Request.QueryString(&quot;submit&quot;) = &quot;form&quot;</font></p>
   <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#0000FF">If</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
    NOT age = &quot;&quot; Then <font color="#009933">' Lets make sure the
    form wasn't submited with nothing checked.</font></font></p>
   <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><font color="#FF0000">
    <font color="#990000">Response.write</font> &quot;You are in the &quot;
    &amp; age &amp; &quot; age range! </font></font></p>
   <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#0000FF">End
    If</font></p>
   <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#0000FF">End
    if</font></p>
   <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><font color="#FF0000">%&gt;</font></font></p>
   <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&lt;/td&gt;<br>
    &lt;/tr&gt;<br>
    &lt;/table&gt;<br>
    &lt;/form&gt;</font></p>
  </td>
 </tr>
 <tr>
  <td bgcolor="#FFFFFF" height="28"> Live Example at <a href="http://www.irideforlife.com/go/default2.asp" target="_blank">http://www.irideforlife.com/go/default2.asp</a></td>
 </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<table width="43%" border="1" cellspacing="0" cellpadding="0" height="94" align="center" bordercolor="#999999" bgcolor="#CCCCCC">
 <tr>
  <td bgcolor="#CCCCCC" height="23"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
 </tr>
 <tr>
  <td bgcolor="#eeeeee" height="161">
   <p align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Please
    post comments below and vote if this helped...thanks!</font></p>
   <p align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">-Jason</font></p>
   </td>
 </tr>
 <tr>
  <td bgcolor="#FFFFFF" height="1">
   <p>&nbsp; </p>
  </td>
 </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
</body>
</html>

