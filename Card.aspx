<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/CardPerso.Master" CodeBehind="Card.aspx.cs" Inherits="CardPerso.Card" EnableEventValidation="false" %>
<asp:Content ID="CardsHeader" runat="server" ContentPlaceHolderID="HeadContent">
    <title>Карты</title>
</asp:Content>
  

<asp:Content ID="CardsContent" runat="server" ContentPlaceHolderID="MainContent">
       <table style="width: 3200px" cellpadding="0" cellspacing="0">
         <tr>
            <td>
                <asp:ImageButton ID="bSetFilter" runat="server" ImageUrl="~/Images/sfil.bmp" 
                    ToolTip="Фильтр"  OnClientClick="return show_flt_card();" 
                    onclick="bSetFilter_Click" />
                <asp:ImageButton ID="bResetFilter" runat="server" ImageUrl="~/Images/rfil.bmp"
                     ToolTip="Снять фильтр" onclick="bResetFilter_Click" />
                <asp:ImageButton ID="bConfFieldD" runat="server" ImageUrl="~/Images/field.bmp"
                     ToolTip="Настройка полей" 
                    OnClientClick="return show_field('type=card','650px');"
                    onclick="bConfFieldD_Click" />
                <asp:ImageButton ID="bDeleteCards" runat="server" ImageUrl="~/Images/del.bmp" 
                    ToolTip="Удаление карт" onclick="bDeleteCard_Click" />
                <asp:ImageButton ID="bFilCards" runat="server" ImageUrl="~/Images/reestr.bmp" 
                    ToolTip="Изменение филиала отправки" OnClientClick="return show_filedit();" onclick="bFilCards_Click" />
                <asp:ImageButton runat="server" ID="bHistory" ImageUrl="~/Images/info.bmp"
                    ToolTip="Информация по карте" OnClick="bHistory_OnClick"/>
                <asp:ImageButton ID="bExcel" runat="server" ImageUrl="~/Images/excel.bmp" 
                         ToolTip="Вывод в Excel" onclick="bExcel_Click" />
                <asp:ImageButton ID="bExcelXML" runat="server" ImageUrl="~/Images/excel2.bmp"
                         ToolTip="Вывод в Excel через XML" OnClick="bExcelXML_Click" />
                <asp:Label ID="lbCount" runat="server" Font-Bold="True" ForeColor="#400080"></asp:Label>
            </td>
        </tr>
        <tr>
        <td>
        <asp:Panel ID="pSearch" runat="server" DefaultButton="bPanSearch">
        <asp:Label runat="server" ID="lMisc">Быстрый поиск по номеру карты:</asp:Label>
        <asp:TextBox runat="server" ID="tbPanSearch"></asp:TextBox>
        <asp:Label runat="server" ID="Label1">Держатель:</asp:Label>
        <asp:TextBox runat="server" ID="tbFioSearch"></asp:TextBox>
        <asp:Button runat="server" ID="bPanSearch" OnClick="bPanSearch_Click" Text="Поиск" />
        </asp:Panel>
        </td>
        </tr>
        <tr>
            <td>
                <%/*="Branch_main_filial=" + getBF().ToString() + " Branch_current=" + getBC().ToString()*/%>
                <div id="pCard" style="overflow: auto;">
                <%/*
                <asp:GridView ID="gvCard" runat="server" BackColor="White" 
                BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="3" 
                AutoGenerateColumns="False" DataKeyNames="id,id_stat,pan" AllowSorting="True" 
                        onsorting="gvCard_Sorting" Width="100%" AllowPaging="True" 
                        onpageindexchanging="gvCard_PageIndexChanging">
                */ 
                %>
                <asp:ObjectDataSource ID="ods" TypeName="CardPerso.Card"  SortParameterName="SortExpression"   
                       SelectCountMethod="GetCardsCount" EnablePaging="true" SelectMethod="GetCards"  OnObjectCreating="CardCreating" 
                       MaximumRowsParameterName="MaximumRows" StartRowIndexParameterName="StartRowIndex"  runat="server"/> 
                <asp:GridView ID="gvCard" runat="server" BackColor="White" 
                BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="3" 
                AutoGenerateColumns="False" DataKeyNames="id,id_stat,pan,id_branchcurrent" AllowSorting="True" 
                        onsorting="gvCard_Sorting" Width="100%" AllowPaging="True" 
                        onpageindexchanging="gvCard_PageIndexChanging" DataSourceID="ods" OnRowDataBound="gvCard_OnRowDataBound" OnSelectedIndexChanged="gvCard_OnSelectedIndexChanged">
                    <PagerSettings Mode="NumericFirstLast" Position="TopAndBottom" />
                <FooterStyle BackColor="White" ForeColor="#000066" />
                <RowStyle ForeColor="#000066" />
                <Columns>                    
                    <asp:TemplateField>
                        <HeaderTemplate>
                            <asp:CheckBox runat="server" ID="cbSelectAll" AutoPostBack="true" OnCheckedChanged="cbSelectAll_CheckedChanged"/>
                        </HeaderTemplate>
                        <ItemTemplate>
                            <asp:CheckBox runat="server" ID="cbSelect" AutoPostBack="true" OnCheckedChanged="cbSelect_CheckedChanged"
                            Visible='<%#((int)Eval("id_stat")==1 || (int)Eval("id_stat")==2 || ((int)Eval("id_stat")==4) && getBF()>0 && getBF()==getBC() && getBC()==(int)Eval("id_branchcurrent"))?true:false%>'/>
                        </ItemTemplate>
                    </asp:TemplateField>
<%--                    <asp:BoundField DataField="fio" HeaderText="ФИО держателя" 
                    SortExpression="fio">
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>--%>
                    <asp:TemplateField HeaderText="ФИО держателя" SortExpression="fio">
                        <ItemTemplate>
                            <table style="align-content:stretch; width:100%;margin:0;">
                                <tr>
                                    <td style="align-content:flex-start"><asp:Label runat="server" Text='<%#(string)Eval("fio")%>'></asp:Label></td>
                                    <td style="vertical-align:top;align-content:flex-end; width:20px;"><asp:Image runat="server" ImageUrl="~/Images/icons8-n-16.jpg" Visible='<%#((bool)Eval("isPin"))%>' /></td>
                                </tr>
                            </table>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:BoundField DataField="pan" HeaderText="Номер" 
                    SortExpression="pan" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="prod_name" HeaderText="Продукт" 
                    SortExpression="prod_name" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="bank_name" HeaderText="Банк" 
                    SortExpression="bank_name" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="company" HeaderText="Организация" 
                    SortExpression="company" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="dateStart" DataFormatString="{0:d}" HeaderText="Дата начала" 
                    SortExpression="dateStart" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="dateEnd" DataFormatString="{0:d}" HeaderText="Дата окончания" 
                    SortExpression="dateEnd" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="passport" HeaderText="Паспорт" 
                    SortExpression="passport" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="account" HeaderText="Номер счета" 
                    SortExpression="account" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="dateProd" DataFormatString="{0:d}" HeaderText="Дата изготовления" 
                    SortExpression="dateProd" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="status" HeaderText="Статус" 
                    SortExpression="id_stat" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="prop" HeaderText="Состояние" SortExpression="id_prop">
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="date_courier" DataFormatString="{0:d}" HeaderText="Дата передачи" 
                    SortExpression="date_courier" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="courier_name" HeaderText="Сотрудник" 
                    SortExpression="courier_name" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="courier" HeaderText="Курьерская служба" 
                    SortExpression="id_courier" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="invoice" HeaderText="Номер накладной" 
                    SortExpression="invoice" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="DepBranchInit" HeaderText="Филиал выпуска" 
                    SortExpression="DepBranchInit" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="DepBranchCard" HeaderText="Филиал отправки" 
                    SortExpression="DepBranchCard" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="BranchCardTransport" HeaderText="Филиал отправки (полный)"
                    SortExpression="BranchCardTransport">
                    <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="dateReceipt" DataFormatString="{0:d}" HeaderText="Дата получения" 
                    SortExpression="dateReceipt" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="dateClient" DataFormatString="{0:d}" HeaderText="Дата выдачи клиенту" 
                    SortExpression="dateClient" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="clientWorker" HeaderText="Выдавший сотрудник" SortExpression="clientWorker">
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>                    
                    <asp:BoundField DataField="dateSendTerminate" DataFormatString="{0:d}" HeaderText="Дата об уничтожение" 
                    SortExpression="dateSendTerminate" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="Comment" HeaderText="Причина отметки" 
                    SortExpression="Comment" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="dateGetTerminate" DataFormatString="{0:d}" HeaderText="Дата получения (филиал)" 
                    SortExpression="dateGetTerminate" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="dateTerminated" DataFormatString="{0:d}" HeaderText="Дата уничтожения" 
                    SortExpression="dateTerminated" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:TemplateField>
                        <ItemTemplate>
                            <asp:HiddenField runat="server" ID="cardid" Value='<%#Eval("id")%>' />
                        </ItemTemplate>
                    </asp:TemplateField>
                  </Columns>
                <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
                <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
                <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                <EmptyDataTemplate>
                    <div style=" font-weight:bold; color:White; background-color: #669999;">
                        Нет записей
                    </div>
                </EmptyDataTemplate>
                </asp:GridView>
                </div>
            </td>
        </tr>
      </table>  
    <asp:Label ID="lbSearch" runat="server" Visible="False"></asp:Label>
    <asp:Label ID="lbSort" runat="server" Visible="False"></asp:Label>
    <asp:Label ID="lbSortIndex" runat="server" Visible="False"></asp:Label>
    <div id="card_history" title="" class="ui-widget" style="display: none">
        <table width="100%">
            <tr><td align="center">
                <asp:Label runat="server" id="lCardHistory" Font-Size="Large"></asp:Label>
            </td></tr>
            <tr><td>
                <asp:GridView runat="server" ID="dgCardHistory" AutoGenerateColumns="false" Width="90%" BorderStyle="None">
                    <Columns>
                        <asp:BoundField DataField="number_doc" HeaderText="Номер" ItemStyle-Width="10%" HeaderStyle-HorizontalAlign="Left">
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="name" HeaderText="Операция" ItemStyle-Width="30%"  HeaderStyle-HorizontalAlign="Left">
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="date_doc" HeaderText="Дата" ItemStyle-Width="20%"  HeaderStyle-HorizontalAlign="Left"  DataFormatString="{0:dd.MM.yyyy}">
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="department" HeaderText="Отделение" ItemStyle-Width="40%"  HeaderStyle-HorizontalAlign="Left">
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </td></tr>
        </table>
    </div>
<script type="text/javascript">
    function clickHistory() {
        HideWait();
        $("#card_history").dialog(
            {
                title: "История карты",
                resizable: false,
                modal: true,
                width: 700,
                close: function(event, ui) {
                    $(this).dialog("destroy");
                    return true;
                },
                buttons: {
                    Закрыть: function() {
                        $(this).dialog("close");
                        return;
                    }
                }
            }
        );
        $("#card_history").keydown(function(event) {
            if (event.keyCode == 13) {
                $(this).parent().find("button:eq(0)").trigger("click");
                return false;
            }
        });
        $("#card_history").focus();
        return false;
    }
</script>
</asp:Content>    
