<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="StorDocCardEdit.aspx.cs" Inherits="CardPerso.StorDocCardEdit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Добавление карточек</title>
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
                <td colspan="3" 
                    style="border-width: thin; border-style: groove;"> 
                     <asp:ImageButton ID="bSave" runat="server" ImageUrl="~/Images/save.bmp" 
                     ToolTip="Сохранить" onclick="bSave_Click" />    
                </td>
            </tr>
          <tr>
                <td rowspan="2" style="width: 20%">
                    <asp:RadioButtonList ID="rbType" runat="server" ForeColor="#000080">
                        <asp:ListItem Value="0">Файл</asp:ListItem>
                        <asp:ListItem Value="1">Карта</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
                <td style="width: 25%">
                    <asp:Label ID="Label1" runat="server" Text="Выберите файл"></asp:Label>
                    
                </td>
                <td>
                    <asp:DropDownList ID="dListFile" runat="server" Width="100%"></asp:DropDownList>
                </td>
            </tr> 
            <tr>
                <td style="width: 25%">
                    <asp:Label ID="Label6" runat="server" Text="Выберите карту"></asp:Label>
                    
                </td>
                <td>
                    <asp:DropDownList ID="dListCard" runat="server" Width="100%"></asp:DropDownList>
                </td>
            </tr> 
            <tr>
                <td colspan="3" align="right">
                    <asp:Label ID="lbCount" Visible="false" runat="server"></asp:Label>
                    <asp:Label ID="lbInform" runat="server" ForeColor="#400080"></asp:Label>
                </td>
            </tr>
       </table>
    </form>
</body>
</html>
