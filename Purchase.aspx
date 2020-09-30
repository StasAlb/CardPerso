<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/CardPerso.Master" CodeBehind="Purchase.aspx.cs" Inherits="CardPerso.Purchase" %>
<asp:Content ID="PurchaseHeader" runat="server" ContentPlaceHolderID="HeadContent">
    <title>Закупочные договора</title>
</asp:Content>
<asp:Content id="PurchaseContent" runat="server" ContentPlaceHolderID="MainContent">
    <table width="100%" cellpadding="0" cellspacing="0">
         <tr>
            <td>
                <asp:ImageButton ID="bNewD" runat="server" ImageUrl="~/Images/new.gif" 
                    ToolTip="Новый договор" 
                    OnClientClick="return show_purchase_dog('mode=1');" onclick="bNewD_Click" />
                <asp:ImageButton ID="bEditD" runat="server" ImageUrl="~/Images/edit.bmp" 
                     ToolTip="Редактировать договор" onclick="bEditD_Click" />
                <asp:ImageButton ID="bDeleteD" runat="server" ImageUrl="~/Images/del.bmp" 
                    ToolTip="Удалить договор" onclick="bDeleteD_Click" />
                <asp:ImageButton ID="bSetFilter" runat="server" ImageUrl="~/Images/sfil.bmp" 
                     ToolTip="Фильтр"  OnClientClick="return show_flt_purchase();" 
                    onclick="bSetFilter_Click" />
                <asp:ImageButton ID="bResetFilter" runat="server" ImageUrl="~/Images/rfil.bmp" 
                     ToolTip="Снять фильтр" onclick="bResetFilter_Click" />
                <asp:ImageButton ID="bConfFieldD" runat="server" ImageUrl="~/Images/field.bmp" 
                     ToolTip="Настройка полей" 
                     OnClientClick="return show_field('type=purchase_dog','280px');" 
                    onclick="bConfFieldD_Click" />
                <asp:ImageButton ID="bExcel" runat="server" ImageUrl="~/Images/excel.bmp" 
                     ToolTip="Вывод в Excel" onclick="bExcel_Click" />
            </td>
            <td align="right">
                <asp:Label ID="lbInform" runat="server" ForeColor="Red"></asp:Label>
                <asp:Label ID="lbCountD" runat="server" Font-Bold="True" ForeColor="#400080"></asp:Label>
            </td>
        </tr>
        <tr>
            <td colspan="2">
               <div id="pDogs" style="overflow: auto; height: 200px;">
               <asp:GridView ID="gvPurchDogs" runat="server" AutoGenerateColumns="False" 
                BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" 
                CellPadding="3" DataKeyNames="id,number_dog,date_dog,date_stor,id_sup,id_manuf,date_record,comment" 
                Width="98%" onselectedindexchanged="gvPurchDogs_SelectedIndexChanged" 
                AllowSorting="True" onsorting="gvPurchDogs_Sorting" AllowPaging="True" 
                       onpageindexchanging="gvPurchDogs_PageIndexChanging" PageSize="5">
                   <PagerSettings Mode="NumericFirstLast" />
                <FooterStyle BackColor="White" ForeColor="#000066" />
                <RowStyle ForeColor="#000066" />
                <Columns>
                    <asp:BoundField DataField="number_dog" HeaderText="Номер" 
                        SortExpression="number_dog" >
                    </asp:BoundField>
                    <asp:BoundField DataField="date_dog" DataFormatString="{0:d}" 
                        HeaderText="Дата" SortExpression="date_dog" >
                    </asp:BoundField>
                    <asp:BoundField DataField="date_stor" DataFormatString="{0:d}" 
                        HeaderText="Дата хранилище" SortExpression="date_stor" >    
                    </asp:BoundField>
                    <asp:BoundField DataField="supplier" HeaderText="Поставщик" 
                        SortExpression="id_sup" >
                    </asp:BoundField>
                    <asp:BoundField DataField="manuf" HeaderText="Изготовитель" 
                        SortExpression="id_manuf" >
                    </asp:BoundField>
                    <asp:BoundField DataField="date_record" DataFormatString="{0:d}" 
                        HeaderText="Дата выписки" SortExpression="date_record" >
                    </asp:BoundField>
                    <asp:BoundField DataField="comment" HeaderText="Отметка выписки" 
                        SortExpression="comment" >
                    </asp:BoundField>
                    <asp:CommandField ButtonType="Image" SelectImageUrl="~/Images/select.gif" 
                        SelectText="Выбрать" ShowSelectButton="True" />
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
             </div>
            </td>
        </tr>
        <tr>
            <td style="height: 30px" valign="bottom">
                    <asp:ImageButton ID="bNewP" runat="server" ImageUrl="~/Images/new.gif" 
                        ToolTip="Новый продукт" onclick="bNewP_Click" />
                    <asp:ImageButton ID="bEditP" runat="server" ImageUrl="~/Images/edit.bmp" 
                     ToolTip="Редактировать продукт" onclick="bEditP_Click" />
                    <asp:ImageButton ID="bDeleteP" runat="server" ImageUrl="~/Images/del.bmp" 
                    ToolTip="Удалить продукт" onclick="bDeleteP_Click" />
                    <asp:ImageButton ID="bConfFieldP" runat="server" ImageUrl="~/Images/field.bmp"
                     ToolTip="Настройка полей" 
                        OnClientClick="return show_field('type=purchase_product','220px');" 
                        onclick="bConfFieldP_Click" />
            </td>
            <td align="right">
                <asp:Label ID="lbCountP" runat="server" Font-Bold="True" ForeColor="#400080"></asp:Label>
            </td>
        </tr>
        <tr>
            <td colspan="2">
                <div id="pProducts" style="overflow: auto; height: 200px;">
                <asp:GridView ID="gvProducts" runat="server" BackColor="White" 
                BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="3" 
                AutoGenerateColumns="False" Width="98%" DataKeyNames="id,id_prb,prod_name,bank_name,cnt" 
                     onselectedindexchanged="gvProducts_SelectedIndexChanged">
                <FooterStyle BackColor="White" ForeColor="#000066" />
                <RowStyle ForeColor="#000066" />
                <Columns>
                    <asp:BoundField DataField="prod_name" HeaderText="Наименование" >
                    </asp:BoundField>
                    <asp:BoundField DataField="bank_name" HeaderText="Банк" >
                    </asp:BoundField>
                    <asp:BoundField DataField="cnt" HeaderText="Кол-во" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="price" HeaderText="Цена" DataFormatString="{0:f2}" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="summa" HeaderText="Сумма" DataFormatString="{0:f2}" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:CommandField ButtonType="Image" SelectImageUrl="~/Images/select.gif" 
                        SelectText="Выбрать" ShowSelectButton="True" ItemStyle-HorizontalAlign="Center" />
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
             </div>
            </td>
        </tr>
    </table>
    <asp:Label ID="lbSearch" runat="server" Visible="False"></asp:Label>
    <asp:Label ID="lbSort" runat="server" Visible="False"></asp:Label>
    <asp:Label ID="lbSortIndex" runat="server" Visible="False"></asp:Label>    
    <script type="text/javascript">
    window.onload = load_purchase;
    </script>
</asp:Content>
