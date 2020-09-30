<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Test.aspx.cs" Inherits="CardPerso.Administration.Test" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script type="text/javascript" src="Admin.js"></script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <OBJECT id="DemoActiveX" classid="clsid:BDEEC1F5-FE44-465c-A1A7-B3203C5C9B40" codebase="ReaderActiveX.cab"></OBJECT>  
 
        <script type="text/javascript">
            try {
                var obj = document.DemoActiveX;
                if (obj) {
                    alert(obj.SayHello());
                } else {
                    alert("Object is not created!");
                }
            } catch (ex) {
                alert("Some error happens, error message is: " + ex.Description);
            }
       
        </script>
</body>
</html>
