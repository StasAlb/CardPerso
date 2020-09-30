<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PurchaseProductEdit.aspx.cs" Inherits="CardPerso.PurchaseProductEdit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Добавление продукта</title>
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
                    <asp:Label ID="Label13" runat="server" Text="Продукт"></asp:Label>
                </td>
                <td style="width: 80%">
                    <asp:DropDownList ID="dListProd" runat="server" Width="100%"></asp:DropDownList>
                </td>
            </tr>  
            <tr>
                <td style="width: 15%">
                    <asp:Label ID="Label14" runat="server" Text="Количество"></asp:Label>
                </td>
                <td style="width: 80%">
                <table width="100%" cellpadding="0" cellspacing="0">
                        <tr>
                            <td style="width: 40%">
                                  <asp:TextBox ID="tbCount" runat="server" Width="95%"></asp:TextBox>
                            </td>
                            <td style="width: 20%" align="center">
                                <asp:Label ID="Label15" runat="server" Text="Цена"></asp:Label>
                            </td>
                            <td style="width: 40%">
                                <asp:TextBox ID="tbPrice" runat="server" Width="95%"></asp:TextBox>
                            </td>
                        </tr>   
                    </table>
                </td>
            </tr> 
            <tr>
                <td colspan="2" align="right">
                    <asp:Label ID="lbCount" Visible="false" runat="server"></asp:Label>
                    <asp:Label ID="lbInform" runat="server" ForeColor="#400080"></asp:Label>
                </td>
            </tr>
       </table>
    </form>
</body>
</html>
