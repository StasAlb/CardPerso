<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="OrganizationEdit.aspx.cs" Inherits="CardPerso.OrganizationEdit" %>
<%@ Register Assembly="DatePicker" Namespace="OstCard.WebControls" TagPrefix="ost" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Справочник организаций</title>
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
                    <asp:Label ID="Label1" runat="server" Text="Организация"></asp:Label>
                </td>
                <td style="width: 70%">
                     <asp:TextBox ID="tbTitle" runat="server" Width="98%" MaxLength="150" Enabled="false"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td style="width: 30%"> 
                    <asp:Label ID="Label2" runat="server" Text="Эмбоссировано на карте"></asp:Label>
                </td>
                <td style="width: 70%">
                    <asp:TextBox ID="tbEmboss" runat="server" Width="98%" MaxLength="50" Enabled="false"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td style="width: 30%"> 
                    <asp:Label ID="Label3" runat="server" Text="Доверенное лицо"></asp:Label>
                </td>
                <td style="width: 70%">
                    <asp:TextBox ID="tbPerson" runat="server" Width="98%" MaxLength="150"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td style="width: 30%"> 
                    <asp:Label ID="Label4" runat="server" Text="Должность"></asp:Label>
                </td>
                <td style="width: 70%">
                    <asp:TextBox ID="tbPosition" runat="server" Width="98%" MaxLength="50"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td style="width: 30%"> 
                    <asp:Label ID="Label5" runat="server" Text="Серия номер паспорта" ></asp:Label>
                </td>
                <td style="width: 70%">
                    <asp:TextBox ID="tbPassport" runat="server" Width="98%" MaxLength="15"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td style="width: 30%"> 
                    <asp:Label ID="Label6" runat="server" Text="Кем выдан"></asp:Label>
                </td>
                <td style="width: 70%">
                    <asp:TextBox ID="tbPDivision" runat="server" Width="98%" MaxLength="150"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td style="width: 30%"> 
                    <asp:Label ID="Label7" runat="server" Text="Дата выдачи"></asp:Label>
                </td>
                <td style="width: 70%">
                          <ost:DatePicker ID="DatePickerPassport" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
            <tr>
                <td style="width: 30%"> 
                    <asp:Label ID="Label8" runat="server" Text="Номер доверенности"></asp:Label>
                </td>
                <td style="width: 70%">
                    <asp:TextBox ID="tbDoveren" runat="server" Width="98%" MaxLength="30"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td style="width: 30%"> 
                    <asp:Label ID="Label9" runat="server" Text="Срок действия доверенности"></asp:Label>
                </td>
                <td style="width: 70%">
                <table><tr><td>
                          <ost:DatePicker ID="DatePickerStart" runat="server" AutoPostBack="true" PaneWidth="150px" >
        <PaneTableStyle BorderColor="#707070" BorderWidth="1px" BorderStyle="Solid" />
        <PaneHeaderStyle BackColor="#0099FF" />
        <TitleStyle ForeColor="White" Font-Bold="true" />
        <NextPrevMonthStyle ForeColor="White" Font-Bold="true" />
        <NextPrevYearStyle ForeColor="#E0E0E0" Font-Bold="true" />
        <DayHeaderStyle BackColor="#E8E8E8" />
        <TodayStyle BackColor="#FFFFCC" ForeColor="#000000" Font-Underline="false" BorderColor="#FFCC99"/>
        <AlternateMonthStyle BackColor="#F0F0F0" ForeColor="#707070" Font-Underline="false"/>
        <MonthStyle BackColor="" ForeColor="#000000" Font-Underline="false"/>            
            </ost:DatePicker></td><td>-</td><td><ost:DatePicker ID="DatePickerEnd" runat="server" AutoPostBack="true" PaneWidth="150px" >
        <PaneTableStyle BorderColor="#707070" BorderWidth="1px" BorderStyle="Solid" />
        <PaneHeaderStyle BackColor="#0099FF" />
        <TitleStyle ForeColor="White" Font-Bold="true" />
        <NextPrevMonthStyle ForeColor="White" Font-Bold="true" />
        <NextPrevYearStyle ForeColor="#E0E0E0" Font-Bold="true" />
        <DayHeaderStyle BackColor="#E8E8E8" />
        <TodayStyle BackColor="#FFFFCC" ForeColor="#000000" Font-Underline="false" BorderColor="#FFCC99"/>
        <AlternateMonthStyle BackColor="#F0F0F0" ForeColor="#707070" Font-Underline="false"/>
        <MonthStyle BackColor="" ForeColor="#000000" Font-Underline="false"/>            
            </ost:DatePicker></td></tr></table> 
                </td>
            </tr>
            <tr>
                <td colspan="2" align="right">
                    <asp:Label ID="lInform" runat="server" ForeColor="#400080"></asp:Label>
                </td>
                
            </tr>            
       </table>
    </form>
</body>
</html>
