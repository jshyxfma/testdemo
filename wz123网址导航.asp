<% 
id=request("id") 
liehuo_net_url=request("liehuo_net_url") 
If liehuo_net_url="" Then liehuo_net_url="wz123µ¼º½" 
Shortcut = "[InternetShortcut] " & vbCrLf 
Shortcut = Shortcut & "URL=http://www.wz123.net/" & vbCrLf 
Shortcut = Shortcut & "IDList=" & vbCrLf 
Shortcut = Shortcut & "IconFile=http://www.wz123.net/favicon.ico" & vbCrLf 
Shortcut = Shortcut & "IconIndex=1" & vbCrLf 
Shortcut = Shortcut & "[{000214A0-0000-0000-C000-000000000046}] " & vbCrLf 
Shortcut = Shortcut & "Prop3=19,2 " & vbCrLf 
Shortcut = Shortcut & " " & vbCrLf 
Response.AddHeader "Content-Disposition", "attachment;filename="&liehuo_net_url&".url;" 
Response.ContentType = "application/octet-stream" 
Response.Write Shortcut 
%>
