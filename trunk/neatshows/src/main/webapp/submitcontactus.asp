<%
on error resume next

if (   (len(trim(request("email"))))>3  ) then

	strsql = "INSERT INTO [MailingLists].[dbo].[email] ([site],[address],[country]) VALUES (4,'"&request("email")&"','"&request("country")&"')"

	Dim conn
	set conn = Server.CreateObject("ADODB.Connection")
	conn.Open = "driver={SQL Server};server=mssql1.gearhost.com;uid=eanh;pwd=xxxxxxxx;database=mailinglists;"
	conn.execute(strSQL)
	conn.close
	set conn = nothing
	'response.write(strsql)
	'response.write("<br>" & len((trim(request("email")))))

End If

%>

<HTML>
<HEAD><TITLE></TITLE></HEAD>
<BODY BGCOLOR="FFFFFF">
<TABLE BORDER="0"  CELLPADDING="0" CELLSPACING="3" width="100%" style="border-collapse: collapse" bordercolor="#111111">


	<TR>

		<TD ALIGN="CENTER" COLSPAN="3">
		<IMG SRC="images/contactus.gif" ALT="I got contactus all over my hand once, it was terrible..." width="562" height="103">
		</TD>


	</TR>
	<TR>
		<TD ALIGN="CENTER" COLSPAN="3">
		<IMG SRC="images/blank.jpg"  HEIGHT="10" BORDER="0">
		</TD>
	</TR>

	<TR>
		<TD ALIGN="CENTER" COLSPAN="1" WIDTH="20%">
		</TD>
		<TD ALIGN="CENTER" COLSPAN="1">
			<FONT FACE="verdana, arial, helvetica" SIZE="3">
				Thanks, we will let you know when we're playing, getting nominated for the nobel prize, etc.
			</FONT>
		</TD>
		<TD ALIGN="CENTER" COLSPAN="1" WIDTH="20%">
		</TD>
	</TR>

	</TABLE>