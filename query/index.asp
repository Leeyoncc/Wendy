<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="X-UA-Compatible" content="IE=7">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!-- #include file="../inc/conn.asp" -->
<!-- #include file="../inc/web_config.asp" -->
<!-- #include file="../inc/html_clear.asp" -->
<%
search_q=request.querystring("q")
%>
<title>��Ѱ��<%=search_q%> - ����֪ʶѧϰ�����������֪ʶ����</title>
<link href="/css/ElegantBlue/style.css" rel="stylesheet" type="text/css" media="screen" />
<script  language="javascript" src="/js/slidealbum.js"></script>
</head>
<body>
<%
keywords=split(search_q," ")
c=ubound(keywords)
for i=0 to c
if i=0 then
search_sql1=search_sql1&"where  ( [title] like '%"&keywords(i)&"%'"
keywords_all=keywords(i)
else
search_sql1=search_sql1&" or   [title] like '%"&keywords(i)&"%'"
keywords_all=keywords_all&"+"&keywords(i)
end if
next

s_sql="select [title],[content],[file_path],[time] from [article] "&search_sql1&" )  and view_yes=1 order by [time] desc"
%>
<div id="wrapper">
	<div id="header-wrapper">
		<div id="header">
			<div id="logo">
		<h1><a href="/" title="����֪ʶѧϰ�����������֪ʶ����">����֪ʶѧϰ�����������֪ʶ����</a></h1>		
		<p class="slogan">�ý�վ��ø��Ӽ򵥣�����Ե�</p>	
			</div>
		</div>
	</div>
	<!-- end #header -->
	<div id="menu-wrapper">
		<div id="menu">
<ul><li><a href='/' >�� ҳ</a></li> <li><a href='/blog/' >�� ��</a></li> <li><a href='/album/' >�� ��</a></li> <li><a href='/Category/About' >�� ��</a></li> <li><a href='/FeedBack/' >�� ��</a></li> <li><a href='/Category/Contact' >�� ϵ</a></li> <li><a href='http://www.66diannao.com/' target='_blank'>�ٷ���վ</a></li> </ul>	
                </div>
	</div>
	<!-- end #menu -->
	<div id="page">
		<div id="page-bgtop">
			<div id="page-bgbtm">
				<div id="page-content">
					<div id="content">
			<div class="Web_Position">�����ڵ�λ��: <a href="/">��ҳ</a> > <a href='/query/'>��Ѱ</a></div>
		<div class="clearfix"></div>		
<!--search content start-->
<div id="search_content" class="clearfix">

<%
if search_q<>"" then 

set rs=server.createobject("adodb.recordset")
rs.open(s_sql),cn,1,1
%>

<%'=============��ҳ���忪ʼ��Ҫ�������ݿ��֮��
if err.number<>0 then '������
response.write "���ݿ����ʧ�ܣ�" & err.description
err.clear
else
if not (rs.eof and rs.bof) then '����¼���Ƿ�Ϊ��
r=cint(rs.RecordCount) '��¼����
rowcount = 10 '����ÿһҳ�����ݼ�¼�����ɸ���ʵ���Զ���
rs.pagesize = rowcount '��ҳ��¼��ÿҳ��ʾ��¼��
maxpagecount=rs.pagecount '��ҳҳ��
page=request.querystring("page")
  if page="" then
  page=1
  end if
rs.absolutepage = page 
rcount1=0
pagestart=page-5
pageend=page+5
if pagestart<1 then
pagestart=1
end if
if pageend>maxpagecount then
pageend=maxpagecount
end if
rcount=rs.RecordCount
'=============��ҳ�������%>

<!--position start-->
<div class="searchtip">��������Ѱ"<span class="FontRed"><%=search_q%></span>",�ҵ������Ϣ <span class="font_brown"><%=rcount%></span> ��</div>
<!--position end-->
<!--list start-->
<div class="result_list">
<div class="gray">��ʾ���ÿո���������Ѱ�ؼ��ʿɻ�ȡ������������'���� ����'��</div>
<dl>

<%'===========ѭ���忪ʼ
do while not rs.eof and rowcount%>
<%
title1=left(rs("title"),30)
for i=0 to c
title1=Replace(title1, keywords(i), "<span class='FontRed'>" & keywords(i)& "</span>")
next

content1=left(Clearhtml(rs("content")),110)
for i=0 to c
content1=Replace(content1,keywords(i), "<span class='FontRed'>" & keywords(i)& "</span>")
next
%>
<dt ><a href='<%="/"&Article_FolderName&"/"&rs("file_path")%>' target='_blank' title='<%=rs("title")%>'><%=title1%></a></dt>
<dd><%=content1%>...</dd>
<dd class="font12 arial font_green line"><a href='<%="/"&Article_FolderName&"/"&rs("file_path")%>' target='_blank'><span class="font_green"><%=web_url&Article_FolderName&"/"&rs("file_path")%></span></a><%=year(rs("time"))%>-<%=month(rs("time"))%>-<%=day(rs("time"))%></dd>
<%
rowcount=rowcount-1 
rs.movenext
loop
 '===========ѭ�������%>

</dl>
</div>
<!--list end-->

<!--page start-->
<div class="result_page clearfix">
<!--#include file="../inc/page_list.asp"-->
</div>
<!--page end-->

<%
else
response.write "<div class='search_welcome'>�ܱ�Ǹ,û���ҵ��� <span class='FontRed'>"&search_q&"</span> ��ص���Ϣ��<p >��ʾ���ÿո���������Ѱ�ؼ��ʿɻ�ȡ������������'���� ����'��</p></div>"
end if
end if
end if%>
</div>
<!--search content end-->	
		<div style="clear: both;">&nbsp;</div>
		</div>
					<!-- end #content -->
					<div id="sidebar">
			<ul>
	<li>		<div id="search">
			<form method="get" action="/query/index.asp">
				<fieldset>
				<input type="text" name="q" id="search-text" size="15" onBlur="if(this.value=='') this.value='Input Keywords ...';" 
onfocus="if(this.value=='Input Keywords ...') this.value='';" value="Input Keywords ..." /><input type="submit" id="search-submit" value="�� Ѱ" />
				</fieldset>
			</form>
		</div>
		<div class="clearfix"></div>
	</li>		
				<li>
					<h2>���˵���</h2>
					<p class="myphoto"><img src="/images/up_images/myphoto.jpg" width="80" height="89"></p>
					<p class="myintro">����֪ʶ<br>2013��05��04��<br>�й� �Ϻ�<br>121673232@qq.com</p>
					<p class="clearfix">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������Լ���һƬ��أ��ҿ���������졣����</a></p>
					</li>
				<li>
					<h2>�������</h2>
<!--slide album start-->
<div id="slidebox">
	<div id="slider">
<div class='slide'><A href='/Gallery/64/' target='_blank' title='����ͼƬ'><img class='diapo' src='/photos/20120529123831900.jpg' alt='����ͼƬ' width='210' ></a><div class='text'><span class='title'>����ͼƬ</span></div></div><div class='slide'><A href='/Gallery/64/' target='_blank' title='����ͼƬ'><img class='diapo' src='/photos/20120529123831633.jpg' alt='����ͼƬ' width='210' ></a><div class='text'><span class='title'>����ͼƬ</span></div></div><div class='slide'><A href='/Gallery/64/' target='_blank' title='����ͼƬ'><img class='diapo' src='/photos/20120529123831810.jpg' alt='����ͼƬ' width='210' ></a><div class='text'><span class='title'>����ͼƬ</span></div></div><div class='slide'><A href='/Gallery/64/' target='_blank' title='����ͼƬ'><img class='diapo' src='/photos/20120529123831996.jpg' alt='����ͼƬ' width='210' ></a><div class='text'><span class='title'>����ͼƬ</span></div></div><div class='slide'><A href='/Gallery/64/' target='_blank' title='����ͼƬ'><img class='diapo' src='/photos/20120529123831600.jpg' alt='����ͼƬ' width='210' ></a><div class='text'><span class='title'>����ͼƬ</span></div></div><div class='slide'><A href='/Gallery/64/' target='_blank' title='����ͼƬ'><img class='diapo' src='/photos/20120529123831821.jpg' alt='����ͼƬ' width='210' ></a><div class='text'><span class='title'>����ͼƬ</span></div></div>
	</div>
<script type="text/javascript">
/* ==== start script ==== */
slider.init();
</script>
</div>
<!--slide album end-->
				</li>
				<li>
					<h2>���ͷ���</h2>
<ul><li><a href='/Category/Enterntainment/'>������Ѷ</a> (3) <a href='/rss/Feed.asp?CatID=133' target='_blank'><img src='/images/rss_icon.gif'></a></li><li><a href='/Category/Love/'>������</a> (2) <a href='/rss/Feed.asp?CatID=135' target='_blank'><img src='/images/rss_icon.gif'></a></li><li><a href='/Category/Internet/'>������</a> (6) <a href='/rss/Feed.asp?CatID=136' target='_blank'><img src='/images/rss_icon.gif'></a></li><li><a href='/Category/Favorite/'>�����ղ�</a> (1) <a href='/rss/Feed.asp?CatID=138' target='_blank'><img src='/images/rss_icon.gif'></a></li><li><a href='/Category/Diary/'>�����ռ�</a> (1) <a href='/rss/Feed.asp?CatID=137' target='_blank'><img src='/images/rss_icon.gif'></a></li></ul>
				</li>
				<li>
					<h2>��������</h2>
<ul><li><a href='/html/0651825813.html' target='_blank'>�ձ��������������㳵 ֻ��һ������</a> (12)</li><li><a href='/html/2875165721.html' target='_blank'>��ͳ��ݺ�����ҵ��ֹ����������ҵ</a> (4)</li><li><a href='/html/943726544.html' target='_blank'>�ȸ�ƻ��΢���Ż������ƶ�������г�</a> (2)</li><li><a href='/html/624185537.html' target='_blank'>ŵ���Ǽƻ���δ���ֻ��м����ˮ����</a> (1)</li><li><a href='/html/7026945042.html' target='_blank'>����Ͱͺ�֧�����ڻ���¥ ½���콫</a> (1)</li><li><a href='/html/4670284933.html' target='_blank'>ǰ���Ѳ���ɧ�� ���ֽ��ֻ����ҵ���</a> (1)</li><li><a href='/html/1387244638.html' target='_blank'>�ɵ��˲��Ӵ��Է�Ҫ�۹��� ������</a> (1)</li><li><a href='/html/6854304456.html' target='_blank'>��Ůʱ��֣������͵���Χ�� �¸���</a> (1)</li><li><a href='/html/1820354344.html' target='_blank'>��ԫ��ɳ����д�漯 �ƽ��������ֹ</a> (1)</li><li><a href='/html/7043195512.html' target='_blank'>�������������ܹ�˾7000��Ա����</a> (0)</li></ul>
				</li>
			</ul>
		</div>
					<!-- end #sidebar -->
				</div>
				<div style="clear: both;">&nbsp;</div>
			</div>
		</div>
	</div>
	<!-- end #page -->
</div>
<div id="footer">
	<p>Copyright &copy; 2012 ���ɸ��˲���ϵͳBritar Yao All rights reserved<BR> Powered by ���ͳ��� <img src="/images/hituxblog-logo.png" width="80" alt='HituxBlog V1.4'> <a href="/rss" target="_blank"><img src="/images/rss_icon.gif"></a> <a href="/rss/feed.xml" target="_blank"><img src="/images/xml_icon.gif"></a> Themes By <A href="http://www.66diannao.com/" target=_blank>����֪ʶ</a>
</p>
</div>
<!-- end #footer -->
</body>
</html>

