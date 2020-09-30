<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ListDeliverEdit.aspx.cs" Inherits="CardPerso.ListDeliverEdit" %>

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
                <td colspan="3" style="border-width: thin; border-style: groove;"> 
                     <asp:ImageButton ID="bSave" runat="server" ImageUrl="~/Images/save.bmp" 
                     ToolTip="Сохранить" onclick="bSave_Click" />    
                </td>
            </tr>
            <asp:Panel ID="pDeliver" runat="server">
            <tr>
                <td style="width: 15%"> 
                    <asp:Label ID="Label6" runat="server" Text="Наименование" Font-Size="Medium"></asp:Label>
                </td>
                <td style="width: 80%">
                     <asp:TextBox ID="tbName" runat="server" Width="99%"></asp:TextBox>
                </td>
                <td>
                </td>
            </tr>
            </asp:Panel> 
            <tr>
                <td style="width: 15%"> 
                    <asp:Label ID="Label1" runat="server" Text="Филиал"></asp:Label>
                </td>
                <td style="width: 80%">
                     <asp:DropDownList ID="dListFilial" runat="server" Width="100%">
                    </asp:DropDownList>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td colspan="3" align="right">
                    <asp:Label ID="lbInform" runat="server" ForeColor="#400080"></asp:Label>
                </td>
            </tr>
       </table>
    </form>
</body>
</html>
