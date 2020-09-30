<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="FltPurchase.aspx.cs" Inherits="CardPerso.FltPurchase" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Фильтр по закупочным договорам</title>
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
            <td colspan="2" style="border-width: thin; border-style: groove;"> 
                 <asp:ImageButton ID="bSave" runat="server" ImageUrl="~/Images/save.bmp" 
                 ToolTip="Сохранить" onclick="bSave_Click" />    
            </td>
        </tr>
        <tr>
            <td style="width: 25%">
                  <asp:Label ID="Label1" runat="server" Text="Номер договора"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td style="width: 40%">
                         <asp:TextBox ID="tbNumber" runat="server" Width="95%"></asp:TextBox>
                    </td>
                    <td style="width: 20%"></td>
                    <td style="width: 40%"></td>
                </tr>   
            </table>
            </td>
        </tr> 
         <tr>
            <td style="width: 25%">
                 <asp:Label ID="Label2" runat="server" Text="Дата договора с"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td style="width: 40%">
                         <asp:TextBox ID="tbDataSt" runat="server" Width="95%"></asp:TextBox>
                    </td>
                    <td align="center" style="width: 20%">
                         <asp:Label ID="Label3" runat="server" Text="по"></asp:Label>
                    </td>
                    <td style="width: 40%">
                        <asp:TextBox ID="tbDataEnd" runat="server" Width="95%"></asp:TextBox>
                    </td>
                </tr>   
            </table>
            </td>
        </tr>
        <tr>
            <td style="width: 25%">
                  <asp:Label ID="Label4" runat="server" Text="Поставщик"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:DropDownList ID="dListSup" runat="server" Width="99%"></asp:DropDownList>
                    </td>
                </tr>   
            </table>
            </td>
        </tr>   
        <tr>
            <td style="width: 25%">
                  <asp:Label ID="Label6" runat="server" Text="Изготовитель"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:DropDownList ID="dListManuf" runat="server" Width="99%"></asp:DropDownList>
                    </td>
                </tr>   
            </table>
            </td>
        </tr>  
        <tr>
            <td style="width: 25%">
                <asp:Label ID="Label5" runat="server" Text="Продукция"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:TextBox ID="tbProd" runat="server" Width="98%"></asp:TextBox>
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
