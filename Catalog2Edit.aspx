<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Catalog2Edit.aspx.cs" Inherits="CardPerso.Catalog2Edit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Редактирование</title>
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
                <td style="width: 30%;"> 
                    <asp:Label ID="Label1" runat="server" Text="Название организации"></asp:Label>
                </td>
                <td style="width: 70%">
                     <asp:TextBox ID="tbName" runat="server" Width="98%" MaxLength=150></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td style="width: 30%;">
                    <asp:Label ID="Label2" runat="server" Text="Эмбоссируется на карте"></asp:Label>                    
                </td>
                <td style="width: 30%;">
                    <asp:TextBox ID="tbEmboss" runat="server" Width="98%"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="right">
                    <asp:Label ID="lbInform" runat="server" ForeColor="#400080"></asp:Label>
                </td>
            </tr>
       </table>
    </form>
</body>
</html>
