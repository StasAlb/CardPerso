<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/Catalogs.Master" CodeBehind="Supplier.aspx.cs" Inherits="CardPerso.Supplier" %>
<asp:Content ID="SuppliersHeader" runat="server" ContentPlaceHolderID="CatalogHeaderContent">
    <title>Поставщики</title>
</asp:Content>
<asp:Content ID="SuppliersContent" runat="server" ContentPlaceHolderID="CatalogContent">
            <table width="100%">
            <tr>
            <td>
                <asp:ImageButton ID="bNew" runat="server" ImageUrl="~/Images/new.gif" 
                    ToolTip="Новый" OnClientClick="return show_catalog('type=supplier&mode=1');" onclick="bNew_Click" />
                <asp:ImageButton ID="bEdit" runat="server" ImageUrl="~/Images/edit.bmp" 
                     ToolTip="Редактировать" onclick="bEdit_Click"   />
                <asp:ImageButton ID="bDelete" runat="server" ImageUrl="~/Images/del.bmp" 
                    ToolTip="Удалить" onclick="bDelete_Click" />
                <asp:ImageButton ID="bExcel" runat="server" ImageUrl="~/Images/excel.bmp" 
                     ToolTip="Вывод в Excel" onclick="bExcel_Click" />     
            </td>
            <td align="right">
                <asp:Label ID="lbInform" runat="server" ForeColor="Red"></asp:Label>
                <asp:Label ID="lbCount" runat="server" Font-Bold="True" ForeColor="#400080"></asp:Label>
            </td>
            </tr>
            
             <tr>
             <td colspan="2">
                <div id="pSupplier" 
                     style="overflow: auto; height: 300px;">
                <asp:GridView ID="gvSuppliers" runat="server" AutoGenerateColumns="False" width="97%"  
                    BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" 
                    CellPadding="3" 
                    DataKeyNames="id,name" 
                    onselectedindexchanged="gvSuppliers_SelectedIndexChanged">
                    <FooterStyle BackColor="White" ForeColor="#000066" />
                    <RowStyle ForeColor="#000066" />
                    <Columns>
                        <asp:BoundField DataField="name" HeaderText="Наименование" >
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
    <script type="text/javascript">
    window.onload = load_supplier;
    </script>
</asp:Content>
