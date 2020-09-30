<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="CardProperties.aspx.cs" Inherits="CardPerso.CardProperties" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Свойство карты</title>
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
        <asp:ImageButton runat="server" ID="bSave" OnClick="bSave_Click" ImageUrl="~/Images/save.bmp"
            ToolTip="Сохранить" />
    </div>
    <asp:Panel runat="server" ID="pProperty">
        <asp:Label runat="server" ID="lTitle"></asp:Label><br />
        <asp:DropDownList runat="server" id="ddlProperty"></asp:DropDownList>
        <br /><asp:CheckBox ID="selectallcard" runat="server" Checked="false" Text="применить для всех карт документа" TextAlign="Right" />
    </asp:Panel>
    <asp:Panel runat="server" ID="pRename">
        <table>
            <tr><td>
                Держатель:
            </td>
            <td>
                <asp:TextBox runat="server" ID="tbFio" Width="200"></asp:TextBox>
            </td>
            </tr>
            <tr>
            <td>
            Паспорт:
            </td>
            <td>
            <asp:TextBox runat="server" ID="tbPass" Width="200"></asp:TextBox>
            </td>
            </tr>
        </table>
    </asp:Panel>
    </form>
</body>
</html>
