﻿<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="FilEdit.aspx.cs" Inherits="CardPerso.FilEdit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Изменение филиала отправки</title>
    <meta http-equiv="Pragma" content="no-cache" />
    <base target="_self" />
    <style type="text/css">
        span
        {
/*        	font-weight:bold; */
           	color:#000080;  
        }
    </style>
</head>
<body style="margin: 0px; background-color: #F7F7DE;">
    <form id="form1" runat="server">
    <table width="100%">
            <tr>
                <td colspan="2" 
                    style="border-width: thin; border-style: groove;"> 
                     <asp:ImageButton ID="bSave" runat="server" ImageUrl="~/Images/save.bmp" 
                     ToolTip="Сохранить" onclick="bSave_Click" />    
                </td>
            </tr>
            <tr>
                <td style="width: 15%">
                    <asp:Label ID="Label13" runat="server" Text="Филиал"></asp:Label>
                </td>
                <td style="width: 80%">
                    <asp:DropDownList ID="dListBranch" runat="server" Width="100%"></asp:DropDownList>
                </td>
            </tr>  
       </table>
    </form>
</body>
</html>
