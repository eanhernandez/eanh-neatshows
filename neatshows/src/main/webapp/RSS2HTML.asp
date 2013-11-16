<%
 Response.Expires = -1

 ' =========== RSS2HTML.ASP for ASP/ASP.NET ==========
 ' copyright 2005-2007 (c) www.Bytescout.com
 '  version 1.24, 1 October 2007
 ' ===============================================
 ' ##############################################################
 ' ####### CHECK OUR COMMERCIAL PRODUCTS FOR ASP/ASP.NET: #######
 ' SWF Scout [ http://bytescout.com/swfscout.html ]- create, read, modify flash movies (SWF)
 ' SWF SlideShow Scout [ http://bytescout.com/swfslideshowscout.html ]- convert JPEG,PNG,BMP into slideshow flash (SWF) with effects
 ' PDFDoc Scout [ http://bytescout.com/pdfdocscout.html ]- generate PDF documents with security options
 ' ##############################################################

 ' =========== configuration =====================
 ' ##### URL to RSS Feed to display #########
 URLToRSS = "http://blog.myspace.com/blog/rss.cfm?friendID=13026540" '"http://www.bytescout.com/blog/feed"

 s_Output = ""

 ' ##### max number of displayed items #####
 MaxNumberOfItems = 7

 ' ##### Main template constants
 'MainTemplateHeader = "<table border=1>"
 'MainTemplateFooter = "</table>"
 ' #####

' ######################################
 Keyword1 =  ""  ' Keyword1 =  "tech" - set non-empty keyword value to filter by this keyword
 Keyword2 = "" ' Keyword1 =  "win" - set non-empty keyword value to filter by this 2nd keyword too
' #################################

 ' ##### Item template.
 ' ##### {LINK} will be replaced with item link
 ' ##### {TITLE} will be replaced with item title
 ' ##### {DESCRIPTION} will be replaced with item description
 ' ##### {DATE} will be replaced with item date and time
 ' ##### {COMMENTSLINK} will be replaced with link to comments (if you use RSS feed from blog)
 ' ##### {CATEGORY} will be replaced with item category
 ItemTemplate = "<TD COLSPAN=""1"" ALIGN=""LEFT"" WIDTH=""70%""><FONT FACE=""verdana, arial, helvetica"" SIZE=""3""><strong>({DATE}) {TITLE}</strong> {DESCRIPTION} {LINK}</FONT></td>"

 ' ##### Error message that will be displayed if not items etc
 'ErrorMessage = "Error has occured while trying to process " &URLToRSS & "<BR>Please contact web-master"

 ' ================================================

 Set xmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
 xmlHttp.Open "Get", URLToRSS, false
 xmlHttp.Send()
 RSSXML = xmlHttp.ResponseText

 Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
 xmlDOM.async = False
 xmlDOM.validateOnParse = False
 xmlDom.resolveExternals = False

 If not xmlDOM.LoadXml(RSSXML) Then
     'ErrorMessage = "Can not load XML:" & vbCRLF & xmlDOM.parseError.reason & vbCRLF & ErrorMessage
 End If

 Set xmlHttp = Nothing ' clear HTTP object

 Set RSSItems = xmlDOM.getElementsByTagName("item") ' collect all "items" from downloaded RSS

 RSSItemsCount = RSSItems.Length-1

 ' if not <item>..</item> entries, then try to get <entry>..</entry>
 if RSSItemsCount = -1 Then
 Set RSSItems = xmlDOM.getElementsByTagName("entry") ' collect all "entry" (atom format) from downloaded RSS
 RSSItemsCount = RSSItems.Length-1

 End If

 Set xmlDOM = Nothing ' clear XML


 ' writing Header
 if RSSItemsCount > 0 then
  s_Output = s_Output & MainTemplateHeader
 End If

 j = -1



 For i = 0 To RSSItemsCount

 if i_ShowAllBlogEntries <> 0 Then s_Output = s_Output & "<TR> <TD ALIGN=""CENTER"" COLSPAN=""1"">"


 Set RSSItem = RSSItems.Item(i)

  for each child in RSSItem.childNodes

   Select case lcase(child.nodeName)
     case "title"
           RSStitle = child.text
     case "link"
	   If RSSLink = "" Then
		If child.Attributes.length>0 Then
		 RSSLink = child.GetAttribute("href")
		 	 if (RSSLink <> "") Then
 			  if child.GetAttribute("rel") <> "alternate" Then
	   			 RSSLink = ""
			  End If
 	   		 End If
		End If ' if has attributes
 		If RSSLink = "" Then
		 	RSSlink = child.text
	   	End If
	   End If
     case "description"
           RSSdescription = child.text
     case "content" ' atom format
           RSSdescription = child.text
     case "published"' atom format
           RSSDate = child.text
     case "pubdate"
           RSSDate = child.text
     case "comments"
           RSSCommentsLink = child.text
     case "category"
	  	Set CategoryItems = RSSItem.getElementsByTagName("category")
		RSSCategory = ""
	  		for each categoryitem in CategoryItems
           			if RSSCategory <> "" Then
					RSSCategory = RSSCategory & ", "
				End If

					RSSCategory = RSSCategory & categoryitem.text
			Next
   End Select



  next

' now check filter
 If (InStr(RSSTitle,Keyword1)>0) or (InStr(RSSTitle,Keyword2)>0) or (InStr(RSSDescription,Keyword1)>0) or (InStr(RSSDescription,Keyword2)>0) then

	  j = J+1

	  if J<MaxNumberOfItems then


		  if Len (RSSDescription) > 200 Then
		  	RSSlink = "(<A HREF=""" & RSSlink & """>more</A>)"
		  else
		  	RSSlink = ""
		  end if

		  ItemContent = Replace(ItemTemplate,"{LINK}",RSSlink)


		  ItemContent = Replace(ItemContent,"{TITLE}",RSSTitle)
		  ItemContent = Replace(ItemContent,"{DATE}",Mid(RSSDate,6,11))
		  ItemContent = Replace(ItemContent,"{COMMENTSLINK}",RSSCommentsLink)
		  ItemContent = Replace(ItemContent,"{CATEGORY}",RSSCategory)

		  s_Output = s_Output & Replace(ItemContent,"{DESCRIPTION}",RSSDescription)


		  if i_ShowAllBlogEntries <> 0 Then
		  		s_Output = s_Output & "<TD WIDTH=""15%""></TD></TR><TR><TD COLSPAN=""1""></TD><TD COLSPAN=""1""><HR></TD><TD COLSPAN=""1""></TD></TR>"
		  End If



		  ItemContent = ""
		  RSSLink = ""
	  End if
 End If


if i_ShowAllBlogEntries = 0 Then Exit For End If

 Next

 ' writing Footer
 if RSSItemsCount > 0 then
  s_Output = s_Output & MainTemplateFooter
 else
  s_Output = Application("s_blogentry") & MainTemplateFooter
 End If

Application("s_blogentry") = s_Output

response.write(s_Output)

' Response.End ' uncomment this for use in on-the-fly output
%>
