<%
Set objTypeLib = Server.CreateObject("ScriptLet.TypeLib")
guidNew = Left(objTypeLib.GUID, 38)
Set objTypeLib = Nothing
Response.Write Replace(Replace(guidNew, "}", ""), "{", "")
%>