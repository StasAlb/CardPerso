<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="FltUsers.aspx.cs" Inherits="CardPerso.Administration.FltUsers" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Фильтр по пользователям</title>
    <meta http-equiv="Pragma" content="no-cache" />
    <base target="_self" />
    <style type="text/css">
        span
        {
           	color:#000080;  
        }
    </style>
</head>
<body style="margin: 0px; background-color: #F7F7DE;">
    <form id="form1" runat="server">
    <table width="100%">
        <tr>
            <td colspan="2" style="border-width: thin; border-style: groove;"> 
                 <asp:ImageButton ID="bSave" runat="server" ImageUrl="~/Images/save.bmp" 
                 ToolTip="Сохранить" onclick="bSave_Click" />    
            </td>
        </tr>
        <tr>
            <td style="width: 30%">
                <asp:Label ID="Label1" runat="server" Text="Логин"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:TextBox ID="tbLogin" runat="server" Width="98%"></asp:TextBox>
                    </td>
                 </tr>   
            </table>
            </td>
         </tr>
        <tr>
            <td style="width: 30%">
                <asp:Label ID="Label7" runat="server" Text="ФИО пользователя"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:TextBox ID="tbFio" runat="server" Width="98%"></asp:TextBox>
                    </td>
                 </tr>   
            </table>
            </td>
         </tr>
        <tr>
            <td style="width: 30%">
                <asp:Label ID="Label2" runat="server" Text="Должность"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:TextBox ID="tbProf" runat="server" Width="98%"></asp:TextBox>
                    </td>
                 </tr>   
            </table>
            </td>
         </tr>
         <tr>
            <td style="width: 30%">
                  <asp:Label ID="Label4" runat="server" Text="Отделение"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:DropDownList ID="dListBranch" runat="server" Width="99%"></asp:DropDownList>
                    </td>
                </tr>   
            </table>
            </td>
        </tr>
         <tr>
            <td style="width: 30%">
                  <asp:Label ID="Label3" runat="server" Text="Роль"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:DropDownList ID="dListRole" runat="server" Width="99%"></asp:DropDownList>
                    </td>
                </tr>   
            </table>
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

