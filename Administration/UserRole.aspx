<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="UserRole.aspx.cs" Inherits="CardPerso.Administration.UserRole" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Роли пользователя</title>
    <meta http-equiv="Pragma" content="no-cache" />
    <base target="_self" />
    <style type="text/css">
        span
        {
           	color:#000080;  
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table border="0" width="100%">
            <tr>
                <td align="left" style="width:100%; border-width: thin; border-style: groove;" colspan="3">
                    <asp:ImageButton ID="bSave" runat="server" ImageUrl="~/Images/save.bmp" 
                        ToolTip="Сохранить" onclick="bSave_Click" />
                </td>
            </tr>
            <tr>
                <td align="center" colspan="3">
                    <h2>
                        Роль пользователя</h2></td>
            </tr>
            <tr>
                <td style="width:20%"></td>
                <td>
                    <asp:CheckBoxList AutoPostBack="false" runat="server" ID="cblRoles" Width="100%"></asp:CheckBoxList>
                </td>
                <td style="width:20%"></td>
            </tr>
            <tr><td colspan="3"></td></tr>
        </table>
    </div>
    </form>
</body>
</html>
