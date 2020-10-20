<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->

	<%
Call header()
%>
<%
set rs2=server.createobject("adodb.recordset")
sql="select * from web_FileSetting "
rs2.open(sql),cn,1,1
if not rs2.eof  then
FileFolder=rs2("FileFolder")
end if
rs2.close
set rs2=nothing
%>

	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th width="100%" height=25 class='tableHeaderText'>删除文件</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
	    <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" bgcolor="#B1CFF8"><div align="center"></div></td>
          </tr>
          <tr>
            <td height="100">
			<%page=request.querystring("page")
			act=request.querystring("act")
			keywords=request.querystring("keywords")
			article_id=cint(request.querystring("id"))
			set rs=server.createobject("adodb.recordset")
sql="select * from web_Files where id="&article_id&""
rs.open(sql),cn,1,3
FileName=rs("FileName")
rs.delete
rs.close
set rs=nothing

'先判断文件是否存在，否则删除
Set fso=Server.CreateObject("Scripting.FileSystemObject")
If fso.FileExists(Server.MapPath("/"&FileFolder&"/"&FileName)) then
FilePath="/"&FileFolder&"/"&FileName
call DelFile(FilePath)
end if

response.Write "<script language='javascript'>alert('删除成功！');location.href='File_list.asp?page="&page&"&act="&act&"&keywords="&keywords&"';</script>"
			%></td>
          </tr>
        </table>
	    </td>
	</tr>
	</table>


<%
Call DbconnEnd()
 %>