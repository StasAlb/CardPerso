<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ConfField.aspx.cs" Inherits="CardPerso.ConfField" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Настройка полей</title>
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
        <table width="100%" cellpadding="2" cellspacing="2">
            <tr>
                <td style="border-width: thin; border-style: groove;"> 
                     <asp:ImageButton ID="bSave" runat="server" ImageUrl="~/Images/save.bmp" 
                     ToolTip="Сохранить" onclick="bSave_Click" />    
                </td>
            </tr>
            <tr>
                <td> 
                    <asp:CheckBoxList ID="chFields" runat="server" 
                        BorderWidth="2px" Width="98%" TextAlign="Right" BorderStyle="Groove" 
                        ForeColor="Navy" CellPadding="2" CellSpacing="2" >
                    </asp:CheckBoxList>
                </td>
            </tr>
       </table>
    </form>
</body>
</html>
