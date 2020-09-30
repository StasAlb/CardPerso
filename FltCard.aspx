<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="FltCard.aspx.cs" Inherits="CardPerso.FltCard" %>
<%@ Register Assembly="DatePicker" Namespace="OstCard.WebControls" TagPrefix="ost" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Фильтр по картам</title>
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
                <asp:Label ID="Label7" runat="server" Text="ФИО держателя"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:TextBox ID="tbFio" runat="server" Width="98%"></asp:TextBox>
                    </td>
                 </tr>   
            </table>
            </td>
         </tr>
        <tr>
            <td style="width: 25%">
                <asp:Label ID="Label8" runat="server" Text="Паспорт"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:TextBox ID="tbPassport" runat="server" Width="98%"></asp:TextBox>
                    </td>
                 </tr>   
            </table>
            </td>
         </tr>
        <tr>
            <td style="width: 25%">
                <asp:Label ID="Label9" runat="server" Text="Номер счета"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:TextBox ID="tbAccount" runat="server" Width="98%"></asp:TextBox>
                    </td>
                 </tr>   
            </table>
            </td>
         </tr>
        <tr>
            <td style="width: 25%">
                <asp:Label ID="Label1" runat="server" Text="Номер карточки"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:TextBox ID="tbNumber" runat="server" Width="98%"></asp:TextBox>
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
                          <asp:DropDownList ID="dListProd" runat="server" Width="99%"></asp:DropDownList>
                    </td>
                 </tr>   
            </table>
            </td>
         </tr>
         <tr>
            <td style="width: 25%">
                  <asp:Label ID="Label4" runat="server" Text="Статус"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:DropDownList ID="dListStatus"  OnSelectedIndexChanged="dListStatus_SelectedIndexChanged" runat="server" Width="99%"></asp:DropDownList>
                    </td>
                </tr>   
            </table>
            </td>
        </tr>
         <tr>
            <td style="width: 25%">
                  <asp:Label ID="Label34" runat="server" Text="Состояние"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:DropDownList ID="ddlProp" runat="server" Width="99%"></asp:DropDownList>
                    </td>
                </tr>   
            </table>
            </td>
        </tr>        
        <tr>
            <td style="width: 25%">
                  <asp:Label ID="Label10" runat="server" Text="Банк"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:DropDownList ID="dListBank" runat="server" Width="99%"></asp:DropDownList>
                    </td>
                </tr>   
            </table>
            </td>
        </tr>   
         <tr>
            <td style="width: 25%">
                  <asp:Label ID="Label11" runat="server" Text="Организация"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                        <asp:DropDownList ID="dListOrgan" runat="server" Width="99%" Visible="false"></asp:DropDownList>
                        <asp:TextBox runat="server" id="tbOrganization" Width="99%"></asp:TextBox>
                    </td>
                </tr>   
            </table>
            </td>
        </tr>
        <tr>
            <td style="width: 25%">
                  <asp:Label ID="Label36" runat="server" Text="Школьные продукты"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                        <asp:DropDownList ID="DListSchoolProduct" runat="server" Width="99%"></asp:DropDownList>
                    </td>
                </tr>   
            </table>
            </td>
        </tr>
        <tr>
            <td style="width: 25%">
                 <asp:Label ID="Label12" runat="server" Text="Начало действия с"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td style="width: 40%">
                          <ost:DatePicker ID="DatePickerBegin1" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
                         <asp:Label ID="Label13" runat="server" Text="по"></asp:Label>
                    </td>
                    <td style="width: 40%">
                        <ost:DatePicker ID="DatePickerBegin2" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
                 <asp:Label ID="Label14" runat="server" Text="Окончание действия с"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td style="width: 40%">
                         <ost:DatePicker ID="DatePickerEnd1" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
                         <asp:Label ID="Label15" runat="server" Text="по"></asp:Label>
                    </td>
                    <td style="width: 40%">
                        <ost:DatePicker ID="DatePickerEnd2" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
                 <asp:Label ID="Label2" runat="server" Text="Дата изготовления с"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td style="width: 40%">
                         <ost:DatePicker ID="DatePickerProd1" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
                        <ost:DatePicker ID="DatePickerProd2" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
                  <asp:Label ID="Label6" runat="server" Text="Филиал отправки"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:DropDownList ID="dListBranchCard" runat="server" Width="99%"></asp:DropDownList>
                    </td>
                </tr>
            </table>
            </td>
        </tr> 
        <tr>
            <td style="width: 25%">
                  <asp:Label ID="Label16" runat="server" Text="Филиал выпуска"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:DropDownList ID="dListBranchProd" runat="server" Width="99%"></asp:DropDownList>
                    </td>
                </tr>   
            </table>
            </td>
        </tr>
        <%if (getViewBranchCurrent() == true)
          {%>
        <tr>
            <td style="width: 25%">
                  <asp:Label ID="Label35" runat="server" Text="Филиал нахождения"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:DropDownList ID="dListBranchCurrent"  runat="server" Width="99%"></asp:DropDownList>
                    </td>
                </tr>   
            </table>
            </td>
        </tr>
        <%}%>
        <tr>
            <td style="width: 25%">
                 <asp:Label ID="Label17" runat="server" Text="Дата передачи с"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td style="width: 40%">
                         <ost:DatePicker ID="DatePickerPered1" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
                         <asp:Label ID="Label18" runat="server" Text="по"></asp:Label>
                    </td>
                    <td style="width: 40%">
                        <ost:DatePicker ID="DatePickerPered2" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
                  <asp:Label ID="Label19" runat="server" Text="Сотрудник"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                        <asp:DropDownList ID="ddlWorker" runat="server" Width="99%"></asp:DropDownList>
                    </td>
                </tr>   
            </table>
            </td>
        </tr>
         <tr>
            <td style="width: 25%">
                  <asp:Label ID="Label20" runat="server" Text="Курьерская служба"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                        <asp:DropDownList ID="ddlCourier" runat="server" Width="99%"></asp:DropDownList>
                    </td>
                </tr>   
            </table>
            </td>
        </tr>
         <tr>
            <td style="width: 25%">
                  <asp:Label ID="Label21" runat="server" Text="Номер накладной"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                        <asp:TextBox ID="tbInvoice" runat="server" Width="98%"></asp:TextBox>
                    </td>
                </tr>   
            </table>
            </td>
        </tr>
        <tr>
            <td style="width: 25%">
                 <asp:Label ID="Label22" runat="server" Text="Дата получения с"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td style="width: 40%">
                         <ost:DatePicker ID="DatePickerReceive1" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
                         <asp:Label ID="Label23" runat="server" Text="по"></asp:Label>
                    </td>
                    <td style="width: 40%">
                        <ost:DatePicker ID="DatePickerReceive2" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
                 <asp:Label ID="Label30" runat="server" Text="Дата выдачи с"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td style="width: 40%">
                         <ost:DatePicker ID="DatePickerClient1" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
                         <asp:Label ID="Label31" runat="server" Text="по"></asp:Label>
                    </td>
                    <td style="width: 40%">
                        <ost:DatePicker ID="DatePickerClient2" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
                  <asp:Label ID="Label33" runat="server" Text="Выдавший сотрудник"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:DropDownList ID="ddlClientWorker" runat="server" Width="99%"></asp:DropDownList>
                    </td>
                </tr>   
            </table>
            </td>
        
        </tr>
        <tr>
            <td style="width: 25%">
                 <asp:Label ID="Label24" runat="server" Text="Дата об уничтожении с"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td style="width: 40%">
                         <ost:DatePicker ID="DatePickerDestroy11" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
                         <asp:Label ID="Label25" runat="server" Text="по"></asp:Label>
                    </td>
                    <td style="width: 40%">
                        <ost:DatePicker ID="DatePickerDestroy12" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
                 <asp:Label ID="Label26" runat="server" Text="Дата получения (филиал) с"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td style="width: 40%">
                         <ost:DatePicker ID="DatePickerFilial1" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
                         <asp:Label ID="Label27" runat="server" Text="по"></asp:Label>
                    </td>
                    <td style="width: 40%">
                        <ost:DatePicker ID="DatePickerFilial2" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
                 <asp:Label ID="Label28" runat="server" Text="Дата уничтожения с"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td style="width: 40%">
                         <ost:DatePicker ID="DatePickerDestroy21" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
                         <asp:Label ID="Label29" runat="server" Text="по"></asp:Label>
                    </td>
                    <td style="width: 40%">
                        <ost:DatePicker ID="DatePickerDestroy22" runat="server" AutoPostBack="true" PaneWidth="150px" >
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
                  <asp:Label ID="Label37" runat="server" Text="Наличие пин-конверта"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                         <asp:DropDownList ID="ddlPins" runat="server" Width="99%">
                             <asp:ListItem Value="1" Text="Не важно"></asp:ListItem>
                             <asp:ListItem Value="2" Text="Присутствует"></asp:ListItem>
                             <asp:ListItem Value="3" Text="Отсутствует"></asp:ListItem>
                         </asp:DropDownList>
                    </td>
                </tr>   
            </table>
            </td>
        </tr>
        
        <tr>
            <td style="width: 25%">
                  <asp:Label ID="Label32" runat="server" Text="Комментарий"></asp:Label>
            </td>
            <td style="width: 75%">
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td>
                          <asp:TextBox ID="tbComment" runat="server" Width="98%"></asp:TextBox>
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
