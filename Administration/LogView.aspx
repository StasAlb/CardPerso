<%@ Page Title="События" Language="C#" MasterPageFile="~/Administration/Administration.Master" AutoEventWireup="true" CodeBehind="LogView.aspx.cs" Inherits="CardPerso.Administration.LogView" %>
<%@ Register Assembly="DatePicker" Namespace="OstCard.WebControls" TagPrefix="ost" %>
<asp:Content ID="LogContent" runat="server" ContentPlaceHolderID="AdministrationContent">
    <asp:Panel runat="server" ID="PanelTwoPeriod" Visible="true">
    <table>
        <tr>
            <td align="left" colspan="2"><table><tr><td>За период с</td><td><ost:DatePicker  ID="DatePickerStart" runat="server" AutoPostBack="true" PaneWidth="150px">
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
            <td>по</td>
            <td><ost:DatePicker ID="DatePickerEnd" runat="server" AutoPostBack="true" PaneWidth="150px">
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
            <td align="left"><asp:Label Font-Size="X-Large" ID="lDatePickerTwoError" runat="server" Visible="false" Text="*" ForeColor="Red"></asp:Label></td>
            </tr>
            </table>
            </td>
        </tr>
        <tr><td style="width:75px;">
         Логин:</td><td> <asp:TextBox ID="tbLogin" runat="server"></asp:TextBox>
        </td></tr>
        <tr><td>
         Действие:</td><td> <asp:TextBox ID="tbEvent" runat="server"></asp:TextBox>
                <asp:ImageButton ID="bRefresh" runat="server" ImageUrl="~/Images/in.bmp" 
                     ToolTip="Обновить" onclick="bRefresh_Click" /> 
                <asp:ImageButton ID="bExcel" runat="server" ImageUrl="~/Images/excel.bmp" 
                     ToolTip="Вывод в Excel" onclick="bExcel_Click" />     
         
        </td></tr>        
        </table>
    </asp:Panel>
    <asp:GridView ID="gvLog" runat="server" AutoGenerateColumns="False" width="97%"  
        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" 
        CellPadding="3"
        DataKeyNames="UserName,ActionDate,Description">
        <FooterStyle BackColor="White" ForeColor="#000066" />
        <RowStyle ForeColor="#000066" />    
        <Columns>
            <asp:BoundField DataField="UserName" HeaderText="Логин" SortExpression="UserName">
                <ItemStyle HorizontalAlign="Left"  Width="20%"/>
            </asp:BoundField>
            <asp:BoundField DataField="ActionDate" HeaderText="Дата" SortExpression="ActionDate">
                <ItemStyle HorizontalAlign="Left"  Width="20%"/>
            </asp:BoundField>
            <asp:BoundField DataField="Description" HeaderText="Действие" SortExpression="Description">
                <ItemStyle HorizontalAlign="Left" />
            </asp:BoundField>
        </Columns>
        <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
        <SelectedRowStyle BackColor="#669999" ForeColor="White" />
        <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
        <EmptyDataTemplate>
        <div style=" font-weight:bold; color:White; background-color: #669999;">
            Нет записей
        </div>   
        </EmptyDataTemplate>
    </asp:GridView>    
</asp:Content>