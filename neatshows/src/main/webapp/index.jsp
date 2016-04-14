<%@ page language="java" contentType="text/html; charset=ISO-8859-1" pageEncoding="ISO-8859-1"%>
<%
//String redirectURL = "default.htm";
//response.sendRedirect(redirectURL);

response.setStatus(301);
response.setHeader( "Location", "hhttp://www.talesfromthebirdbath.com);
response.setHeader( "Connection", "close" );
%>

