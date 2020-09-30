<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="FltStorDoc.aspx.cs" Inherits="CardPerso.FltStorDoc" %>
<%@ Register Assembly="DatePicker" Namespace="OstCard.WebControls" TagPrefix="ost" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Фильтр по операциям</title>
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
                 <asp:Label ID="Label2" runat="server" Text="Дата документа с"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td style="width: 40%">
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
            </ost:DatePicker>
                    </td>
                    <td align="center" style="width: 20%">
                         <asp:Label ID="Label3" runat="server" Text="по"></asp:Label>
                    </td>
                    <td style="width: 40%">
                          <ost:DatePicker ID="DatePickerEnd" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
            </td>
        </tr>
        <tr>
            <td style="width: 25%">
                  <asp:Label ID="Label4" runat="server" Text="Тип документа"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:DropDownList ID="dListType" runat="server" Width="99%"></asp:DropDownList>
                    </td>
                </tr>   
            </table>
            </td>
        </tr>   
        <tr>
            <td style="width: 25%">
                  <asp:Label ID="Label6" runat="server" Text="Филиал"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:DropDownList ID="dListBranch" runat="server" Width="99%"></asp:DropDownList>
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
            <td style="width: 25%">
                <asp:Label ID="Label1" runat="server" Text="Номер карты"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:TextBox ID="tbCard" runat="server" Width="98%"></asp:TextBox>
                    </td>
                 </tr>   
            </table>
            </td>
        </tr>
        <tr>
            <td style="width: 25%">
                <asp:Label ID="Label7" runat="server" Text="Создал"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:TextBox ID="tbCreate" runat="server" Width="98%"></asp:TextBox>
                    </td>
                 </tr>   
            </table>
            </td>
        </tr>        
        <tr>
            <td colspan="2" align="right">
                <asp:CheckBox ID="chAll" runat="server" TextAlign="Left" Text="Показать все документы" ForeColor="Navy" Checked="true" Visible="false" />              
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
