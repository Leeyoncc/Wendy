<!--#include file="../inc/access.asp"  -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/x_to_html/index_to_html.asp" -->
<!-- #include file="../inc/x_to_html/Model_to_html.asp" -->
<!-- #include file="page_next.asp" -->

<% '����ģ��
act=request.querystring("act")
keywords=trim(request.form("keywords"))
cid=request("cid")


if act="search" then
s_sql="select * from web_theme where [name]  like '%"&keywords&"%'  order by [time] desc"
else
s_sql="select * from web_theme order by [time] desc"
end if

%>

<% '���⼤��ģ��
action1=request.querystring("action")
ThemeFolder=request.querystring("ThemeFolder")
ThemeID=request.querystring("ThemeID")
if action1="Edit" then
set rs1=server.createobject("adodb.recordset")
sql="select web_theme,web_ThemeID from web_settings "
rs1.open(sql),cn,1,3
rs1("web_theme")=ThemeFolder
rs1("web_ThemeID")=ThemeID
rs1.update
rs1.close
set rs1=nothing

'���ɸ�����ģ���ļ�
set rs_create=server.createobject("adodb.recordset")
sql="select [id],ModelType,ModelTheme from web_models where  ModelTheme="&ThemeID
rs_create.open(sql),cn,1,1
Do While not rs_create.eof 
l_id=rs_create("id")
ModelType=rs_create("ModelType")
ModelTheme=rs_create("ModelTheme")
Call Model_to_html(l_id)
rs_create.movenext
loop
rs_create.close
set rs_create=nothing

'��������ҳЧ��
call index_to_html()

response.Write "<script language='javascript'>alert('���������óɹ����Ѹ�����ҳ������ҳ���뵽�����ɹ���������������Ŀ���͡��������ݡ��Ż���£�');location.href='ThemeSetting.asp';</script>"
end if
%>
<script language="JavaScript">
<!--
function ask(msg) {
	if( msg=='' ) {
		msg='���棺ɾ���󽫲��ɻָ�������������벻�������';
	}
	if (confirm(msg)) {
		return true;
	} else {
		return false;
	}
}
//-->
</script>
	<%
Call header()
%>

	<table cellpadding='3' cellspacing='1' border='0' class='tableBorder' align=center>
	<tr>
	  <th width="100%" height=25 class='tableHeaderText'>�����б�</th>
	
	<tr><td height="400" valign="top"  class='forumRow'><br>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='TipTitle'>&nbsp;�� ������ʾ</td>
          </tr>
          <tr>
            <td height="30" valign="top" class="TipWords"><p>1������վ���⡱����˼��һ����ҳ���������ݲ��������£�ҳ����ɫ�����ֵı仯��һ�ֱ仯��һ�����⡣��������Ϥ�����˲��;Ϳ������ö������⡣</p>
            <p>2��������ĳ�������Ĭ��ֻ���Զ�������վ��ҳ������ҳ����Ҫ�ֶ���"���ɹ���"��<a href="html_items.asp">������Ŀ</a>��<a href="html_article.asp">��������</a>�Żῴ���޸ĺ��Ч����</p>
            <p>3����������޷����£�����������ԭ��(1)��Ŀռ�д��Ȩ��δ��;(2)������������⣬ҳ����Ҫ��ˢ���²Żῴ��;(3)�����������û�е����ɹ�������������ҳ�档</p>
			 <p>4��ϵͳ�Դ��˶�����⣬���и����������𲽿����У�����ʱ��ע���ɸ��˲��͹ٷ���վ(www.hitux.com)��</p></td>
          </tr>
          <tr>
            <td height="10" ></td>
          </tr>
        </table>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="25" class='forumRowHighLight'>&nbsp;| <a href="Theme_add.asp">�����µ�����</a></td>
          </tr>
          <tr>
            <td height="30"></td>
          </tr>
        </table>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="2">
          <tr>
            <td width="5%" height="30" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">���</div></td>
            <td width="14%" height="30" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">��������<br>(���������ļ���)</div></td>
            <td width="26%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">����Ԥ��</div></td>
            <td width="18%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">Ԥ����ҳ<br>(�롰���á�����Ԥ��)</div></td>
            <td width="10%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">�Ƿ�����</div></td>
            <td width="18%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">����ʱ��</div></td>
            <td width="9%" class="TitleHighlight"><div align="center" style="font-weight: bold;color:#ffffff">����</div></td>
          </tr>
<% '�����б�ģ��
strFileName="ThemeSetting.asp" 
pageno=5
set rs = server.CreateObject("adodb.recordset")
rs.Open (s_sql),cn,1,1
rscount=rs.recordcount
if not rs.eof and not rs.bof then
call showsql(pageno)
rs.move(rsno)
for p_i=1 to loopno
%>
<% if p_i mod 2 =0 then
class_style="forumRow"
else
class_style="forumRowHighLight"
end if%>
            <form name="form1" method="post" action="?action=edit&id=<%=rs("id")%>">
          <tr >
            <td   height="176" class='<%=class_style%>'><div align="center"><%=rs("id")%></div></td>
           <td class='<%=class_style%>' ><div align="center"><%=rs("name")%><br>(<%=rs("folder")%>)</div></td>
            <td height="176" class='<%=class_style%>' ><div align="center"><img src="/images/up_images/<%=rs("image")%>" width="200" height="143" border="0"></div></td>

            <td class='<%=class_style%>' ><div align="center"><a href="/" target="_blank"><br />
            Ԥ����ҳ</a></div></td>
            <td class='<%=class_style%>' ><div align="center">
<%
set rs_theme=server.createobject("adodb.recordset")
sql="select web_theme from web_settings"
rs_theme.open(sql),cn,1,1
if  rs_theme("web_theme")=rs("folder") then
response.write "<img src='images/use_no.jpg' border='0' title='�������Ѿ����óɹ�'>"
else
response.write "<a href='?Action=Edit&ThemeFolder="&rs("folder")&"&ThemeID="&rs("id")&"' title='������ø�����'><img src='images/use_yes.jpg' border='0'></a>"
end if
rs_theme.close
set rs_theme=nothing
%></div></td>
            <td class='<%=class_style%>' ><div align="center"><%=rs("time")%></div></td>
            <td class='<%=class_style%>' ><div align="center"><a href="Theme_edit.asp?id=<%=rs("id")%>&amp;page=<%=page%>&amp;act=<%=act%>&amp;keywords=<%=keywords%>">�޸�</a> | <a href="javascript:if(ask('���棺ɾ���󽫲��ɻָ�������������벻�������')) location.href='Theme_del.asp?id=<%=rs("id")%>&amp;page=<%=page%>&amp;act=<%=act%>&amp;keywords=<%=keywords%>';">ɾ��</a> </div></td>
          </tr></form>
		  		  <%
		  rs.movenext
		  next
		  else
response.write "<div align='center'><span style='color: #FF0000'>�������ӣ�</span></div>"
		  end if 
		  rs.close
		  set rs=nothing
		  %>
		    <tr  >
              <td height="35"  colspan="9" ><div align="center">
                <%call showpage(strFileName,rscount,pageno,false,true,"")%>
           </div></td>
		    </tr>
      </table>
	    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="20" class='forumRow'>&nbsp;</td>
          </tr>
          <tr>
            <td height="25" class='forumRowHighLight'>&nbsp;| ��������</td>
          </tr>
          <tr>
            <td height="70"><form name="form1" method="post" action="?act=search"><div align="center">
<input name="keywords" type="text"  size="35" maxlength="40">
                <label>
                       &nbsp;
                       <input type="submit" name="Submit" value="�� ��">
                </label>
              </div>
            </form>
            </td>
          </tr>
      </table>
	    <br></td>
	</tr>
	</table>

<%
Call DbconnEnd()
 %>