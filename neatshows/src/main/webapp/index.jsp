<%
    //String redirectURL = "default.htm";
    //response.sendRedirect(redirectURL);
    
    <%@ page language="java" contentType="text/html; charset=ISO-8859-1" pageEncoding="ISO-8859-1"%>
    <%
    response.setStatus(301);
    response.setHeader( "Location", "hhttp://www.talesfromthebirdbath.com);
    response.setHeader( "Connection", "close" );
%>
    
%>