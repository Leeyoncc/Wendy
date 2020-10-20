<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<%
act=Request("act")
If act="save" Then 
FileFolder=trim(request.form("FileFolder"))
FileType=trim(request.form("FileType"))
FileSize=trim(request.form("FileSize"))
'FileNameType=trim(request.form("FileNameType"))
FileTime=now()


set rs=server.createobject("adodb.recordset")
sql="select * from web_FileSetting"
rs.open(sql),cn,1,3
OldFolderDir=rs("FileFolder")
rs("FileFolder")=FileFolder
rs("FileType")=FileType
rs("FileSize")=FileSize
'rs("FileNameType")=FileNameType
rs("FileTime")=FileTime
rs.update
rs.close
set rs=nothing

'检测原文件夹是否存在，否则创建
Set Fso=Server.CreateObject("Scripting.FileSystemObject") 
If Fso.FolderExists(Server.MapPath("/"&OldFolderDir))=false Then
NewFolderDir="/"&OldFolderDir
call CreateFolderB(NewFolderDir)
end if
'检测新文件夹是否与原文件夹不同，是则更名。
if FileFolder<>OldFolderDir  then
NewFolderDir="/"&FileFolder
call renamefolder("/"&OldFolderDir,NewFolderDir) 
end if


response.Write "<script language='javascript'>alert('修改成功！')</script>"

end if
 %>

	<%
Call header()

%>
<%
set rs2=server.createobject("adodb.recordset")
sql="select * from web_FileSetting "
rs2.open(sql),cn,1,3
if not rs2.eof and not rs2.bof then
%>
  <form id="form1" name="form1" method="post" action="?act=save">
         <script language='javascript'>
function checksignup1() {
if ( document.form1.FileFolder.value == '' ) {
window.alert('请输入文件存位置^_^');
document.form1.FileFolder.focus();
return false;}

if ( document.form1.FileType.value == '' ) {
window.alert('请输入允许上传文件类型^_^');
document.form1.FileType.focus();
return false;}

if ( document.form1.FileSize.value == '' ) {
window.alert('请输入允许上传文件大小^_^');
document.form1.FileSize.focus();
return false;}

if(document.form1.FileSize.value.search(/^([0-9]*)([.]?)([0-9]*)$/)   ==   -1)   
      {   
  window.alert("文件大小只能是数字^_^");   
document.form1.FileSize.focus();
return false;}


return true;}
</script>
	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>

	<tr>
	  <th class='tableHeaderText' colspan=2 height=25>上传设置</th>
	<tr>
	<tr>
	  <td height=23 colspan="2" class='forumRow'><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="20" class='TipTitle'>&nbsp;√ 操作提示</td>
        </tr>
        <tr>
          <td height="30" valign="top" class="TipWords"><p>1、“文件存放位置”即指你上传的文件放置的地方，一般位于你的空间的根目录下的某个文件夹。</p>
            <p>2、“允许上传文件类型”将决定哪些类型的文件是可以上传的，建议不要设置过多的和陌生的文件类型，以确保系统安全。</p>
			<p>3、“允许上传文件大小”建议不要设置过大，上传太大的文件可能导致超时或无法上传，另外上传文件的大小还可能会受到空间的限制。</p>
			<p>4、对于常见的办公文档如WORD,EXCEL,POWERPOINT等文件一般不会太大，采用系统默认的2M限制已经够用。</p>
            </td>
        </tr>
        <tr>
          <td height="10">&nbsp;</td>
        </tr>
      </table></td>
	  </tr>
	<td width="15%" height=23 class='forumRowHighLight'>文件存放位置</td>
	<td class='forumRowHighLight'><input name='FileFolder' type='text' id='FileFolder' size='40'  value="<%=rs2("FileFolder")%>" >
	  &nbsp;将在系统根目录下建立文件夹存放你上传的文件</td>
	</tr>
	  <tr>
	    <td class='forumRow' height=23>允许上传文件类型</td>
	    <td class='forumRow'><input name='FileType' type='text' id='FileType' value="<%=rs2("FileType")%>" size='80'>
        &nbsp;多个文件扩展名以 / 分开。</td>
      </tr>
	  <tr>
	    <td class='forumRowHighLight' height=23>允许上传文件大小</td>
	    <td class='forumRowHighLight'><input name='FileSize' type='text' id='FileSize' value="<%=rs2("FileSize")%>" size='20'>KB，系统默认限制大小为2MB。1MB=1024KB。</td>
      </tr>	  

	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='提交' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>
<%
end if
rs2.close
set rs2=nothing
%>
<%
Call DbconnEnd()
 %>