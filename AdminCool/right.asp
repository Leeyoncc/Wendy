<!-- #include file="../inc/access.asp" -->
<!-- #include file="inc/functions.asp" -->
<!-- #include file="../inc/html_clear.asp" -->
<link href="images/skin.css" rel="stylesheet" type="text/css" />
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />

  <style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background-color: #EEF2FB;
}
-->
</style>
  <body>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="17" valign="top" background="images/mail_leftbg.gif"><img src="images/left-top-right.gif" width="17" height="29" /></td>
    <td valign="top" background="images/content-bg.gif"><table width="100%" height="31" border="0" cellpadding="0" cellspacing="0" class="left_topbg" id="table2">
      <tr>
        <td height="31"><div class="titlebt">��ӭ����</div></td>
      </tr>
    </table></td>
    <td width="16" valign="top" background="images/mail_rightbg.gif"><img src="images/nav-right-bg.gif" width="16" height="29" /></td>
  </tr>
  <tr>
    <td valign="middle" background="images/mail_leftbg.gif">&nbsp;</td>
    <td valign="top" bgcolor="#F7F8F9"><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="2" valign="top">&nbsp;</td>
        <td>&nbsp;</td>
        <td valign="top">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="2" valign="top"><span class="left_bt"><%=session("log_name")%>����ӭ����<%=gdb("select web_name from web_settings ")%>��վ��̨����ϵͳ</span>��
          <span class="left_txt"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;������ʹ�õ���Britar Yao������һ��ר��Ϊ��׼���ĸ��˲���ϵͳ��������(Hitux)���˲��͡�ͨ����ϵͳ�������Ĳ��ͻ����Ǹ�����վ���������׾١�����Ҫ�߱���ôרҵ����ҳ���֪ʶ������Ҫ�Գ����ж���Ϥ���������غ��ɸ��˲��͵�Դ���ϴ���������Ŀռ����������������վ����������Ҫ����ֻ�Ƕ���վ�ĸ��£�дһƪ���£������ϴ�һ��ͼƬ��������ľ�����������������վ�ϣ������ǽ�����վ��21����������������������վ��ʱ����������������HituxBlogԸ����һ��֮����Я�ֹ�����
<br>
          </span><br>

</td>
        <td width="7%">&nbsp;</td>
        <td width="40%" rowspan="3" valign="top"><table width="100%" height="144" border="0" cellpadding="0" cellspacing="0" class="line_table">
          <tr>
            <td width="7%" height="27" background="images/news-title-bg.gif"><img src="images/news-title-bg.gif" width="2" height="27" /></td>
            <td width="93%" background="images/news-title-bg.gif" class="left_bt2 left_ts">ʹ��֮ǰ����һ����������Ŷ��</td>
          </tr>
          <tr>
            <td height="102" colspan="2" valign="top"><span class="left_ts">1��</span> ��һ�δ���վ��̨���Ƚ�������վ������Ϣ����һ�°ɣ��ڡ��������á�-��<a href="web_settings.asp">��վ��Ϣ����</a>������ <br />
                <span class="left_ts">2��</span> �����������ڰ�ȫ�Կ��ǣ��޸�һ��Ĭ�ϵ�����ɣ��ڡ��������á�-��<a href="admin_list.asp">��̨�û�����</a>������<br />
                <span class="left_ts">3��</span> ��Ҫдһƪ�������ڡ����¹�����-��<a href="article_add.asp">��������</a>������ <br />
                <span class="left_ts">4��</span> ��Ҫ�ϴ�һ��ͼƬ���ڡ���������-��<a href="ad_add.asp">����ͼƬ</a>������ <br />
                <span class="left_ts">5��</span> Ŀǰ�����·��಻�ʺ��㣿�޸����ǰɣ��ڡ����������-��<a href="category_list.asp">�����б�</a>������<br />
                <span class="left_ts">6��</span> ��վ�������Ի�������Ŀǰ�������Ŀ����⹩��ѡ���Լ�����������ҳ��ģ����Զ��壬�������о��о���<a href="ThemeSetting.asp">��������</a> | <a href="web_models.asp">ģ�����</a>�� <br />
                <span class="left_ts">7��</span> ÿ�θ�����������վ�󣬱���������һ��������վ���ڡ����ɹ��������� <br />
                <span class="left_ts">8��</span> ϣ�������鱣����վ�ϳ��ֵĹ�棬���Ƕ����������Ͷ���СС֧�֣�<br />
                <span class="left_ts">9��</span>�����ʣ�������BUG����������ϵ�ķ�ʽ�Ƕ����ģ���ӭ���ٷ���վ(<a href="http://www.hitux.com/" target="_blank">www.hitux.com</a>)���Ի��Ƿ��ʼ�(411159226@qq.com)��</td>
          </tr>
          <tr>
            <td height="5" colspan="2">&nbsp;</td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td colspan="2">&nbsp;</td>
        <td>&nbsp;</td>
        </tr>
      <tr>
        <td colspan="2" valign="top"><!--JavaScript����-->
              <SCRIPT language=javascript>
function secBoard(n)
{
for(i=0;i<secTable.cells.length;i++)
secTable.cells[i].className="sec1";
secTable.cells[n].className="sec2";
for(i=0;i<mainTable.tBodies.length;i++)
mainTable.tBodies[i].style.display="none";
mainTable.tBodies[n].style.display="block";
}
          </SCRIPT>
              <!--HTML����-->
              <TABLE width=72% border=0 cellPadding=0 cellSpacing=0 id=secTable>
                <TBODY>
                  <TR align=middle height=20>
                    <TD align="center" class=sec2 onclick=secBoard(0)>��վ��Ϣ</TD>
                    <TD align="center" class=sec1 onclick=secBoard(1)>������Ϣ</TD>
                    <TD align="center" class=sec1 onclick=secBoard(2)>�����Ϣ</TD>
                    <TD align="center" class=sec1 onclick=secBoard(3)>�ÿ�����</TD>
                  </TR>
                </TBODY>
              </TABLE>
          <TABLE class=main_tab id=mainTable cellSpacing=0
cellPadding=0 width=100% border=0>
                <!--����TBODY���-->
                <TBODY style="DISPLAY: block">
                  <TR>
                    <TD vAlign=top align=middle>
<%
'��ҳ������Ϣ���ݶ�ȡ�滻
set rs=server.createobject("adodb.recordset")
sql="select web_name,web_slogan,web_url,web_title,web_person,web_birthdate,web_birthplace,web_shortintro from web_settings"
rs.open(sql),cn,1,1
if not rs.eof and not rs.bof then
web_name=rs("web_name")
web_url=rs("web_url")
web_slogan=rs("web_slogan")
web_title=rs("web_title")
web_person=rs("web_person")
web_birthdate=rs("web_birthdate")
web_birthplace=rs("web_birthplace")
web_shortintro=rs("web_shortintro")
end if
rs.close
set rs=nothing
%>
					<TABLE width=98% height="110" border=0 align="center" cellPadding=0 cellSpacing=0>
                        <TBODY>
                          <TR>
                            <TD height="5" colspan="3"></TD>
                          </TR>
                          <TR>
                            <TD width="4%" bgcolor="#FAFBFC">&nbsp;</TD>
                            <TD width="42%" height="25" bgcolor="#FAFBFC"><span class="left_txt">��վ���ƣ� </span>
                               
                                <span class="left_ts"><%=web_name%> </span></TD>
                            <TD width="54%" height="25" bgcolor="#FAFBFC"><span class="left_txt">��վ��ַ�� </span>
                               
                                <span class="left_ts"><%=web_url%> </span></TD>
                          </TR>
                          <TR>
                            <TD bgcolor="#FAFBFC">&nbsp;</TD>
                            <TD height="25" bgcolor="#FAFBFC"><span class="left_txt">��ҳ���⣺ </span>
                               
                                <span class="left_ts"> <%=web_title%></span></TD>
                            <TD height="25" bgcolor="#FAFBFC"><span class="left_txt">��վ��� </span>
                               
                                <span class="left_ts"><%=web_slogan%> </span></TD>
                          </TR>
                          <TR>
                            <TD bgcolor="#FAFBFC">&nbsp;</TD>
                            <TD height="25" bgcolor="#FAFBFC"><span class="left_txt">��վվ���� </span>
                                <span class="left_ts"> <%=web_person%></span></TD>
                            <TD height="25" bgcolor="#FAFBFC"><span class="left_txt">�������£� </span>
                               
                                <span class="left_ts"> <%=web_birthdate%></span></TD>
                          </TR>
                          <TR>
                            <TD bgcolor="#FAFBFC">&nbsp;</TD>
                            <TD height="25" bgcolor="#FAFBFC"><span class="left_txt">������ </span>
                               
                                <span class="left_ts"><%=web_birthplace%> </span></TD>
                            <TD height="25" bgcolor="#FAFBFC"><span class="left_txt">��飺 </span>
                                <span class="left_ts"> <%=left(web_shortintro,10)&"..."%></span></TD>
                          </TR>
                          <TR>
                            <TD height="5" colspan="3"></TD>
                          </TR>
                        </TBODY>
                    </TABLE></TD>
                  </TR>
                </TBODY>
                <!--����cells����-->
                <TBODY style="DISPLAY: none">
                  <TR>
                    <TD vAlign=top align=middle><TABLE width=98% height="110" border=0 align="center" cellPadding=0 cellSpacing=0>
                        <TBODY>
                          <TR>
                            <TD height="5" colspan="2"></TD>
                          </TR>
<%
'�����ļ��л�ȡ
set rs_1=server.createobject("adodb.recordset")
sql="select FolderName from web_Models_type where [id]=9"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
Article_FolderName=rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing

set rs=server.createobject("adodb.recordset")
sql="select top 5 [id],[title],[url],[file_path],[time] from [article] where view_yes=1 order by time desc"
rs.open(sql),cn,1,1
if not rs.eof then
do while not rs.eof  %>                          
                          <TR>
                            <TD bgcolor="#FAFBFC">&nbsp;</TD>
                            <TD height="25" bgcolor="#FAFBFC">��<span class="left_txt"><a href="<%="/"&Article_FolderName&"/"&rs("File_Path")%>" target="_blank"><%=left(rs("title"),25)%></a> (<%=rs("time")%>)</span></TD>
                          </TR>
<%
rs.movenext
loop
else
response.write "������Ϣ"
end if
rs.close
set rs=nothing
%>						  
                          <TR>
                            <TD height="5" colspan="2"></TD>
                          </TR>
                        </TBODY>
                    </TABLE></TD>
                  </TR>
                </TBODY>
                <!--����tBodies����-->
                <TBODY style="DISPLAY: none">
                  <TR>
                    <TD vAlign=top align=middle><TABLE width=98% border=0 align="center" cellPadding=0 cellSpacing=0>
                        <TBODY>
                          <TR>
                            <TD colspan="2"></TD>
                          </TR>
                          <TR>
                            <TD height="5" colspan="2"></TD>
                          </TR>
<%
'����ļ��л�ȡ
set rs_1=server.createobject("adodb.recordset")
sql="select FolderName from web_Models_type where [id]=5"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
Gallery_FolderName=rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing

set rs=server.createobject("adodb.recordset")
sql="select top 5 [id],[name],[time] from [web_ad_position] where view_yes=1 order by time desc"
rs.open(sql),cn,1,1
if not rs.eof then
do while not rs.eof  %>                          
                          <TR>
                            <TD bgcolor="#FAFBFC">&nbsp;</TD>
                            <TD height="25" bgcolor="#FAFBFC">��<span class="left_txt"><a href="<%="/"&Gallery_FolderName&"/"&rs("id")%>" target="_blank"><%=left(rs("name"),25)%></a> (<%=rs("time")%>)</span></TD>
                          </TR>
<%
rs.movenext
loop
else
response.write "������Ϣ"
end if
rs.close
set rs=nothing
%>	
                         
                          <TR>
                            <TD height="5" colspan="2"></TD>
                          </TR>
                        </TBODY>
                    </TABLE></TD>
                  </TR>
                </TBODY>
                <!--����display����-->
                <TBODY style="DISPLAY: none">
                  <TR>
                    <TD vAlign=top align=middle><TABLE width=98% border=0 align="center" cellPadding=0 cellSpacing=0>
                        <TBODY>
                          <TR>
                            <TD colspan="3"></TD>
                          </TR>
                          <TR>
                            <TD height="5" colspan="3"></TD>
                          </TR>
<%
'�����ļ��л�ȡ
set rs_1=server.createobject("adodb.recordset")
sql="select FolderName from web_Models_type where [id]=7"
rs_1.open(sql),cn,1,1
if not rs_1.eof then
Post_FolderName=rs_1("FolderName")
end if
rs_1.close
set rs_1=nothing

set rs=server.createobject("adodb.recordset")
sql="select top 5 [content],[time],[name] from web_article_comment where view_yes=1 and article_id=0  order by [time] desc"
rs.open(sql),cn,1,1
if not rs.eof then
do while not rs.eof  %>                          
                          <TR>
                            <TD bgcolor="#FAFBFC">&nbsp;</TD>
                            <TD height="25" bgcolor="#FAFBFC">��<span class="left_txt"><a href="<%="/"&Post_FolderName&"/"%>" target="_blank"><%=left(nohtml(rs("content")),25)&"..."%></a> (<%=rs("time")%>)</span></TD>
                          </TR>
<%
rs.movenext
loop
else
response.write "������Ϣ"
end if
rs.close
set rs=nothing
%>	
                          <TR>
                            <TD height="5" colspan="3"></TD>
                          </TR>
                        </TBODY>
                    </TABLE></TD>
                  </TR>
                </TBODY>
            </TABLE></td>
        <td>&nbsp;</td>
        </tr>
      <tr>
        <td width="2%">&nbsp;</td>
        <td width="51%" class="left_txt">&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
    <td background="images/mail_rightbg.gif">&nbsp;</td>
  </tr>
  <tr>
    <td valign="bottom" background="images/mail_leftbg.gif"><img src="images/buttom_left2.gif" width="17" height="17" /></td>
    <td background="images/buttom_bgs.gif"><img src="images/buttom_bgs.gif" width="17" height="17"></td>
    <td valign="bottom" background="images/mail_rightbg.gif"><img src="images/buttom_right2.gif" width="16" height="17" /></td>
  </tr>
</table>
</body>