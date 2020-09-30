<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ChangeDate.aspx.cs" Inherits="CardPerso.ChangeDate" %>
<%@ Register Assembly="DatePicker" Namespace="OstCard.WebControls" TagPrefix="ost" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Изменение даты</title>
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
        <td style="text-align:center;">
    
    <asp:Label runat="server" ID="lMessage"></asp:Label>    
    </td>
    </tr>
    <tr>
    <td align="center">
                          <ost:DatePicker ID="DatePickerD" runat="server" AutoPostBack="true" PaneWidth="150px">
        <PaneTableStyle BorderColor="#707070" BorderWidth="1px" BorderStyle="Solid" />
        <PaneHeaderStyle BackColor="#0099FF" />
        <TitleStyle ForeColor="White" Font-Bold="true" />
        <NextPrevMonthStyle ForeColor="White" Font-Bold="true" />
        <NextPrevYearStyle ForeColor="#E0E0E0" Font-Bold="true" />
        <DayHeaderStyle BackColor="#E8E8E8" />
        <TodayStyle BackColor="#FFFFCC" ForeColor="#000000" Font-Underline="false" BorderColor="#FFCC99"/>
        <AlternateMonthStyle BackColor="#F0F0F0" ForeColor="#707070" Font-Underline="false"/>
        <MonthStyle BackColor="" ForeColor="#000000" Font-Underline="false"/>            
            </ost:DatePicker>
    </td>
    </tr>
    </table>
    </div>
    </form>
</body>
</html>
