<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PurchaseDogEdit.aspx.cs" Inherits="CardPerso.PurchaseDogEdit" %>

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
                <td style="border-width: thin; border-style: groove;"> 
                     <asp:ImageButton ID="bSave" runat="server" ImageUrl="~/Images/save.bmp" 
                     ToolTip="Сохранить" onclick="bSave_Click" />    
                </td>
            </tr>
            <tr>
                <td>
                    <table width="100%" cellpadding="0" cellspacing="0">
                        <tr>
                            <td style="width: 15%">
                                  <asp:Label ID="Label6" runat="server" Text="Номер договора"></asp:Label>
                            </td>
                            <td style="width: 30%">
                                  <asp:TextBox ID="tbNumber" runat="server" Width="95%"></asp:TextBox>
                            </td>
                            <td style="width: 15%">
                            </td>
                            <td style="width: 30%">
                            </td>
                        </tr>   
                    </table> 
                </td>
            </tr>  
            <tr>
                <td>
                    <table width="100%" cellpadding="0" cellspacing="0">
                        <tr>
                            <td style="width: 15%">
                                 <asp:Label ID="Label8" runat="server" Text="Дата договора"></asp:Label>
                            </td>
                            <td style="width: 30%">
                                 <asp:TextBox ID="tbData" runat="server" Width="95%"></asp:TextBox>
                            </td>
                            <td style="width: 15%">
                                 <asp:Label ID="Label9" runat="server" Text="Дата поступления"></asp:Label>
                            </td>
                            <td style="width: 30%">
                                 <asp:TextBox ID="tbDataSt" runat="server" Width="95%"></asp:TextBox>
                            </td>
                        </tr>   
                    </table> 
                </td>
            </tr> 
             <tr>
                <td>
                     <table width="100%" cellpadding="0" cellspacing="0">
                        <tr>
                            <td style="width: 15%">
                                 <asp:Label ID="Label7" runat="server" Text="Поставщик"></asp:Label>
                            </td>
                            <td style="width: 30%">
                                 <asp:DropDownList ID="dListSup" runat="server" Width="97%"></asp:DropDownList>
                            </td>
                            <td style="width: 15%">
                                 <asp:Label ID="Label10" runat="server" Text="Изготовитель"></asp:Label>
                            </td>
                            <td style="width: 30%">
                                 <asp:DropDownList ID="dListManuf" runat="server" Width="97%"></asp:DropDownList>
                            </td>
                        </tr>   
                    </table> 
                </td>
            </tr> 
            <tr>
                <td>
                    <table width="100%" cellpadding="0" cellspacing="0">
                        <tr>
                            <td style="width: 15%">
                                 <asp:Label ID="Label11" runat="server" Text="Дата выписки"></asp:Label>
                            </td>
                            <td style="width: 30%">
                                 <asp:TextBox ID="tbDataR" runat="server" Width="95%"></asp:TextBox>
                            </td>
                            <td style="width: 15%">
                                 <asp:Label ID="Label12" runat="server" Text="Отметка выписки"></asp:Label>
                            </td>
                            <td style="width: 30%">
                                 <asp:TextBox ID="tbComment" runat="server" Width="95%"></asp:TextBox>
                            </td>
                        </tr>   
                    </table> 
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="lbInform" runat="server" ForeColor="#400080"></asp:Label>
                </td>
            </tr>
        </table> 
    </form>
</body>
</html>
