<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ProductEdit.aspx.cs" Inherits="CardPerso.ProductEdit" %>

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
            <asp:Panel ID="pProduct" runat="server">
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
            <tr>
                <td style="width: 15%"> 
                    <asp:Label ID="Label7" runat="server" Text="Вид"></asp:Label>
                </td>
                <td style="width: 80%">
                     <table width="100%" cellpadding="0" cellspacing="0">
                        <tr>
                            <td style="width: 40%">
                                  <asp:DropDownList ID="dListKind" Width="98%" runat="server"></asp:DropDownList>
                            </td>
                            <td style="width: 20%" align="right">
                                <asp:Label ID="Label11" runat="server" Text="Тип"></asp:Label>
                            </td>
                            <td style="width: 40%" align="right">
                                <asp:TextBox ID="tbType" runat="server" Width="95%"></asp:TextBox>
                            </td>
                        </tr>   
                    </table>   
                </td>
                <td>
                </td>
            </tr>
            </asp:Panel> 
            <tr>
                <td style="width: 15%"> 
                    <asp:Label ID="Label1" runat="server" Text="Банк"></asp:Label>
                </td>
                <td style="width: 80%">
                     <asp:DropDownList ID="dListBank" runat="server" Width="100%">
                    </asp:DropDownList>
                </td>
                <td>
                </td>
            </tr>
             <tr>
                <td style="width: 15%"> 
                    <asp:Label ID="Label8" runat="server" Text="БИН"></asp:Label>
                </td>
                <td style="width: 80%">
                     <table width="100%" cellpadding="0" cellspacing="0">
                        <tr>
                            <td style="width: 40%">
                                  <asp:TextBox ID="tbBin" runat="server" Width="95%"></asp:TextBox>
                            </td>
                            <td style="width: 20%" align="right">
                                <asp:Label ID="Label9" runat="server" Text="Префикс"></asp:Label>
                            </td>
                            <td style="width: 40%" align="right">
                                <asp:TextBox ID="tbPrefix" runat="server" Width="95%"></asp:TextBox>
                            </td>
                        </tr>   
                    </table>   
                </td>
                <td>
                </td>
            </tr>
             <tr>
                <td style="width: 15%"> 
                    <asp:Label ID="Label2" runat="server" Text="Минимум"></asp:Label>
                </td>
                <td style="width: 80%">
                     <table width="100%" cellpadding="0" cellspacing="0">
                        <tr>
                            <td style="width: 40%">
                                  <asp:TextBox ID="tbMinimum" runat="server" Width="95%"></asp:TextBox>
                            </td>
                            <td style="width: 20%">
                            </td>
                            <td style="width: 40%" align="right">
                                <asp:CheckBox ID="cbWrapping" runat="server" Text="Требуется упаковка"></asp:CheckBox>
                                <br />
                                <asp:CheckBox ID="cbInformProduction" runat="server" Text="Информировать об изготовлении"></asp:CheckBox>
                            </td>
                        </tr>   
                    </table>   
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
