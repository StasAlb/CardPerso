<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="StorDocCardRec.aspx.cs"
    Inherits="CardPerso.StorDocCardRec" %>

<%@ Register Assembly="DatePicker" Namespace="OstCard.WebControls" TagPrefix="ost" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">

    <title>Выбор карточек</title>
    <meta http-equiv="Pragma" content="no-cache" />
    <base target="_self" />
    <style type="text/css">
        span
        {
            /*        	font-weight:bold; */
            color: #000080;
        }
    </style>
    
    <script type="text/javascript">
        var cntsel = "<%=cntSelectedCard%>";
        var selectcard="Вы действительно хотите выбрать " + cntsel + " карт?";
    </script>
        
</head>
<body style="margin: 0px; background-color: #F7F7DE;">

    <form id="form1" runat="server">
    <table width="100%" cellpadding="1" cellspacing="1">
        <tr>
            <td>
                <asp:ImageButton ID="bSave" runat="server" ImageUrl="~/Images/save.bmp" ToolTip="Сохранить"
                    OnClick="bSave_Click" Visible="false" 
                    OnClientClick="return confirm(selectcard);"/>
                <asp:ImageButton ID="bFilter" runat="server" ImageUrl="~/Images/select2.bmp" ToolTip="Поиск"
                    OnClick="bFilter_Click" />
            </td>
            <td align="right">
                <asp:Label ID="lbInform" runat="server" ForeColor="Red"></asp:Label>
                <asp:CheckBox ID="cbExclude" runat="server" Text="Исключая карты по фильтру"></asp:CheckBox>
                <asp:Label ID="lbCount" runat="server" Font-Bold="True" ForeColor="#400080"></asp:Label>
            </td>
        </tr>
        <tr>
            <td colspan="2">
                <table width="100%" border="0">
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
                            <asp:Label ID="Label3" runat="server" Text="Организация"></asp:Label>
                        </td>
                        <td style="width: 75%">
                            <table width="100%" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td>
                                        <asp:TextBox ID="tbCompany" runat="server" Width="98%"></asp:TextBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 25%">
                            <asp:Label ID="Label4" runat="server" Text="Организация"></asp:Label>
                        </td>
                        <td style="width: 75%">
                            <table width="100%" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td>
                                        <asp:DropDownList ID="dListCompany" runat="server" Width="99%">
                                        </asp:DropDownList>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 25%">
                            <asp:Label ID="Label6" runat="server" Text="Продукт"></asp:Label>
                        </td>
                        <td style="width: 75%">
                            <table width="100%" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td>
                                        <asp:DropDownList ID="dListProduct" runat="server" Width="99%">
                                        </asp:DropDownList>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>                    
                    <tr>
                        <td style="width: 25%">
                            <asp:Label ID="Label5" runat="server" Text="Дата изготовления с"></asp:Label>
                        </td>
                        <td style="width: 75%">
                            <table width="100%" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td>
                                        <asp:CheckBox AutoPostBack="false" runat="server" ID="cbDateProd" Text="" OnCheckedChanged="cbDateProd_CheckedChanged" />
                                    </td>
                                    <td>
                                        <ost:DatePicker ID="DatePickerStart" Enabled="true" runat="server" AutoPostBack="false"
                                            PaneWidth="150px">
                                            <PaneTableStyle BorderColor="#707070" BorderWidth="1px" BorderStyle="Solid" />
                                            <PaneHeaderStyle BackColor="#0099FF" />
                                            <TitleStyle ForeColor="White" Font-Bold="true" />
                                            <NextPrevMonthStyle ForeColor="White" Font-Bold="true" />
                                            <NextPrevYearStyle ForeColor="#E0E0E0" Font-Bold="true" />
                                            <DayHeaderStyle BackColor="#E8E8E8" />
                                            <TodayStyle BackColor="#FFFFCC" ForeColor="#000000" Font-Underline="false" BorderColor="#FFCC99" />
                                            <AlternateMonthStyle BackColor="#F0F0F0" ForeColor="#707070" Font-Underline="false" />
                                            <MonthStyle BackColor="" ForeColor="#000000" Font-Underline="false" />
                                        </ost:DatePicker>
                                    </td>
                                    <td>
                                        по
                                    </td>
                                    <td>
                                        <ost:DatePicker ID="DatePickerEnd" Enabled="true" runat="server" AutoPostBack="false"
                                            PaneWidth="150px">
                                            <PaneTableStyle BorderColor="#707070" BorderWidth="1px" BorderStyle="Solid" />
                                            <PaneHeaderStyle BackColor="#0099FF" />
                                            <TitleStyle ForeColor="White" Font-Bold="true" />
                                            <NextPrevMonthStyle ForeColor="White" Font-Bold="true" />
                                            <NextPrevYearStyle ForeColor="#E0E0E0" Font-Bold="true" />
                                            <DayHeaderStyle BackColor="#E8E8E8" />
                                            <TodayStyle BackColor="#FFFFCC" ForeColor="#000000" Font-Underline="false" BorderColor="#FFCC99" />
                                            <AlternateMonthStyle BackColor="#F0F0F0" ForeColor="#707070" Font-Underline="false" />
                                            <MonthStyle BackColor="" ForeColor="#000000" Font-Underline="false" />
                                        </ost:DatePicker>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 25%">
                            <asp:Label ID="Label8" runat="server" Text="Невостребованные"></asp:Label>
                        </td>
                        <td style="width: 75%" align="left">
                            <table width="100%" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td>
                                        <asp:CheckBox AutoPostBack="false"  TextAlign="Right" Checked="false" runat="server" ID="NotSend" Text=" (более 30 дней)"/>                                        
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        
        <tr>
            <td colspan="2">
                <div id="pCards" style="overflow: auto; height: 320px;"> <!-- height: 260px; -->
                                
                    
                    <%
                      /*
                        <asp:GridView ID="gvCards" runat="server" AutoGenerateColumns="False" BackColor="#F7F7DE"
                        BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="3" DataKeyNames="id,pan,fio,company"
                        Width="98%" Style="margin-right: 0px" AllowSorting="True">
                      */
                    %>  
                        
                        
                       <asp:ObjectDataSource ID="ods" TypeName="CardPerso.StorDocCardRec"  SortParameterName="SortExpression"   
                       SelectCountMethod="GetCardsCount" EnablePaging="true" SelectMethod="GetCards"  OnObjectCreating="CardCreating" 
                       MaximumRowsParameterName="MaximumRows" StartRowIndexParameterName="StartRowIndex"  runat="server"/>
                    
                        
                        <asp:GridView ID="gvCards" runat="server" AutoGenerateColumns="False" BackColor="#F7F7DE"
                        BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="3" DataKeyNames="id,pan,fio,company"
                        Width="98%" Style="margin-right: 0px" AllowSorting="True" 
                        AllowPaging="True" DataSourceID="ods">
                        <PagerSettings Mode="NumericFirstLast" Position="Bottom" />
                        
                        
                        <FooterStyle BackColor="White" ForeColor="#000066" />
                        <RowStyle ForeColor="#000066" />
                        <Columns>
                            <asp:BoundField DataField="pan" HeaderText="Номер">
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="fio" HeaderText="Клиент">
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="company" HeaderText="Организация">
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                        </Columns>
                        <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
                        <SelectedRowStyle BackColor="#669999" ForeColor="White" />
                        <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                        <EmptyDataTemplate>
                            <div style="font-weight: bold; color: White; background-color: #669999;">
                                Нет записей
                            </div>
                        </EmptyDataTemplate>
                    </asp:GridView>
                </div>
            </td>
        </tr>
        <tr>
            <td style="width: 25%">
                <asp:Label ID="Label2" runat="server" Text="Комментарий"></asp:Label>
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
    </table>
    <asp:Label ID="lbSearch" runat="server" Visible="False"></asp:Label>
    </form>        
    </body>    
    

</html> 
