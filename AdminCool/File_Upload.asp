<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<%
act=Request("act")
If act="save" Then 
FileName=trim(request.form("FileName"))
FileSize=trim(request.form("FileSize"))
FileMemo=trim(request.form("FileMemo"))
FileTime=now()

set rs=server.createobject("adodb.recordset")
sql="select * from web_Files"
rs.open(sql),cn,1,3
rs.addnew
rs("FileName")=FileName
rs("FileSize")=FileSize
rs("FileMemo")=FileMemo
rs("FileTime")=FileTime
rs.update
rs.close
set rs=nothing

response.Write "<script language='javascript'>alert('添加成功！');location.href='File_List.asp';</script>"

end if
 %>

	<%
Call header()

%>
<%
set rs2=server.createobject("adodb.recordset")
sql="select * from web_FileSetting "
rs2.open(sql),cn,1,1
if not rs2.eof  then
FileFolder=rs2("FileFolder")
FileType=rs2("FileType")
FileSize=rs2("FileSize")
end if
rs2.close
set rs2=nothing
%>
  <form id="form1" name="form1" method="post" action="?act=save">
         <script language='javascript'>
function checksignup1() {
if ( document.form1.FileName.value == '' ) {
window.alert('请选择文件上传^_^');
document.form1.FileName.focus();
return false;}

if ( document.form1.FileSize.value == '' ) {
window.alert('请输入文件大小^_^');
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
	  <th class='tableHeaderText' colspan=2 height=25>上传文件</th>
	<tr>
	<tr>
	  <td height=23 colspan="2" class='forumRow'><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td height="20" class='TipTitle'>&nbsp;√ 操作提示</td>
        </tr>
        <tr>
          <td height="30" valign="top" class="TipWords"><p>1、目前系统允许上传文件大小最大为<%=FileSize%>KB。</p>
            <p>2、目前系统允许上传 <%=replace(FileType,"/"," 、")%>等扩展名的文件。</p>
			 <p>3、如果文件比较大，上传可能会耗时过长，请耐心等待，不要连续点击。</p>
			<p>4、文件如果无法上传可能有以下几个原因：(1)你的空间不支持FSO组件；(2)你的空间写入权限未打开;(3)上传文件类型不支持;(4)上传文件超过大小;(5)文件存放文件夹不存在；(6)你的空间已满；(7)你的空间速度过低；(8)黑客入侵了。</p>
			<p>5、如果你确认以上情况都没有出现的话，那么可以联系程序作者了。</p></td>
        </tr>
        <tr>
          <td height="10">&nbsp;</td>
        </tr>
      </table></td>
	  </tr>
	<td width="15%" height=23 class='forumRowHighLight'>上传文件</td>
	<td class='forumRowHighLight'><input name='FileName' type='text' id='FileName' size='30'  value="" >
	  &nbsp;<iframe frameborder="0" width="330" height="23" scrolling="No" src="Upload_File.asp?Action=upFile&Field=FileName&FieldSize=FileSize&FF=<%=FileFolder%>&FS=<%=FileSize%>&FT=<%=FileType%>"></iframe></td>
	</tr>
	  <tr>
	    <td class='forumRow' height=23>文件大小</td>
	    <td class='forumRow'><input name='FileSize' type='text' id='FileSize'  size='20'>KB，系统自动检测文件大小，无需修改。</td>
      </tr>	  	
<tr>
	  <td class='forumRow' height=11>备注</td>
	  <td class='forumRow'><textarea name='FileMemo'  cols="100" rows="6" id="FileMemo" ></textarea></td>
	</tr>	  
	<tr><td height="50" colspan=2  class='forumRow'><div align="center">
	  <INPUT type=submit value='提交' onClick='javascript:return checksignup1()' name=Submit>
	  </div></td></tr>
	</table>
</form>
<%
Call DbconnEnd()
 %>