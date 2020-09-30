<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="BranchEdit.aspx.cs" Inherits="CardPerso.BranchEdit" %>

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
                    <asp:Label ID="Label1" runat="server" Text="Код банка"></asp:Label>
                </td>
                <td style="width: 70%">
                     <asp:TextBox ID="tbKodBank" runat="server" Width="98%"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td style="width: 30%"> 
                    <asp:Label ID="Label2" runat="server" Text="Код подразделения"></asp:Label>
                </td>
                <td style="width: 70%">
                    <asp:TextBox ID="tbKodDep" runat="server" Width="98%"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td style="width: 30%"> 
                    <asp:Label ID="Label3" runat="server" Text="Наименование"></asp:Label>
                </td>
                <td style="width: 70%">
                    <asp:TextBox ID="tbDep" runat="server" Width="98%"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td style="width: 30%"> 
                    <asp:Label ID="Label4" runat="server" Text="Адрес"></asp:Label>
                </td>
                <td style="width: 70%">
                    <asp:TextBox ID="tbAdress" runat="server" Width="98%"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td style="width: 30%"> 
                    <asp:Label ID="Label5" runat="server" Text="Сотрудник"></asp:Label>
                </td>
                <td style="width: 70%">
                    <asp:TextBox ID="tbPeople" runat="server" Width="98%"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td style="width: 30%"> 
                    <asp:Label ID="Label13" runat="server" Text="E-mail"></asp:Label>
                </td>
                <td style="width: 70%">
                    <asp:TextBox ID="tbEmail" runat="server" Width="98%"></asp:TextBox>
                </td>
            </tr>
            <asp:Panel runat="server" ID="pAkBarsProperties">
            <tr>
                <td colspan="2" align="right">
                    <asp:CheckBox ID="chHead" runat="server" TextAlign="Left" Text="Головной офис" ForeColor="Navy" />              
                </td>
            </tr>
            <tr>
                <td colspan="2" align="right">
                    <asp:CheckBox ID="cbRKC" runat="server" TextAlign="Left" Text="РКЦ ДРБ" ForeColor="Navy" />              
                </td>
            </tr>
            <tr>
                <td colspan="2" align="right">
                    <asp:CheckBox ID="cbTrans" runat="server" TextAlign="Left" Text="Транспортная карта" ForeColor="Navy" />              
                </td>
            </tr>        
            <tr>
                <td colspan="2" align="right">
                    <asp:CheckBox ID="cbIsolated" runat="server" TextAlign="Left" Text="Обособленный офис" ForeColor="Navy" Visible="false"/>              
                </td>
            </tr>            
            <tr>
                <td colspan="2" align="right">
                    <asp:Label ID="lbInform" runat="server" ForeColor="#400080"></asp:Label>
                </td>
            </tr>
            </asp:Panel>
       </table>
    </form>
</body>
</html>
