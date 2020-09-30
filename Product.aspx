<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/Catalogs.Master" CodeBehind="Product.aspx.cs" Inherits="CardPerso.Product" %>
<asp:Content ID="ProductHeader" runat="server" ContentPlaceHolderID="CatalogHeaderContent">
    <title>Продукция</title>
</asp:Content>
<asp:Content ID="ProductContent" runat="server" ContentPlaceHolderID="CatalogContent">
    <table width="100%">
            <tr>
            <td>
                <asp:ImageButton ID="bNew" runat="server" ImageUrl="~/Images/new.gif" 
                    ToolTip="Новый" OnClientClick="return show_product('mode=1');" onclick="bNew_Click" />
                <asp:ImageButton ID="bEdit" runat="server" ImageUrl="~/Images/edit.bmp" 
                     ToolTip="Редактировать" onclick="bEdit_Click"   />
                <asp:ImageButton ID="bDelete" runat="server" ImageUrl="~/Images/del.bmp" 
                    ToolTip="Удалить" onclick="bDelete_Click" />
                <asp:ImageButton ID="bAddBank" runat="server" ImageUrl="~/Images/in.bmp" 
                     ToolTip="Привязать банк" onclick="bAddBank_Click" /> 
                <asp:ImageButton ID="bDelBank" runat="server" ImageUrl="~/Images/out.bmp" 
                     ToolTip="Отвязать банк" onclick="bDelBank_Click" /> 
                <asp:ImageButton ID="bLinkProd" runat="server" ImageUrl="~/Images/in.bmp"
                     ToolTip="Привязать к продукту" onclick="bLinkProd_Click" />
                <asp:ImageButton ID="bUnlinkProd" runat="server" ImageUrl="~/Images/out.bmp"
                     ToolTip="Отвязать продукт" onclick="bUnlinkProd_Click" />                     
                <asp:ImageButton ID="bExcel" runat="server" ImageUrl="~/Images/excel.bmp"                 
                     ToolTip="Вывод в Excel" onclick="bExcel_Click" /> 
                <asp:ImageButton ID="bSortUp" runat="server" ImageUrl="~/Images/up.bmp" 
                     ToolTip="Переместить вверх" onclick="bSortUp_Click" /> 
                <asp:ImageButton ID="bSortDown" runat="server" ImageUrl="~/Images/down.bmp" 
                     ToolTip="Переместить вниз" onclick="bSortDown_Click" /> 
            </td>
            <td align="right">
                <asp:Label ID="lbInform" runat="server" ForeColor="Red"></asp:Label>
                <asp:Label ID="lbCount" runat="server" Font-Bold="True" ForeColor="#400080"></asp:Label>
            </td>
            </tr>
            
             <tr>
             <td colspan="2">
                <div id="pProduct" 
                     style="overflow: auto; height: 300px;">
                <asp:GridView ID="gvProducts" runat="server" AutoGenerateColumns="False" width="97%"  
                    BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" 
                    CellPadding="3" 
                        DataKeyNames="id_prb,id_bank,id_prod,id_sort,prod_name,prod_name_h,type_prod,bank_name,bin,prefix_ow,prefix_file,parent" 
                        onselectedindexchanged="gvProducts_SelectedIndexChanged">
                    <FooterStyle BackColor="White" ForeColor="#000066" />
                    <RowStyle ForeColor="#000066" />
                    <Columns>
                        <asp:TemplateField HeaderText="Наименование">
                            <ItemStyle HorizontalAlign="Left"/>
                            <ItemTemplate>
                                <%#(string)Eval("prod_name") %>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="type_prod" 
                            HeaderText="Вид" >
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="prefix_ow" 
                            HeaderText="Тип" >
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="bank_name" 
                            HeaderText="Банк" >
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="bin" 
                            HeaderText="БИН" >
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="prefix_file" 
                            HeaderText="Префикс" >
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="min_cnt" 
                            HeaderText="Минимум" >
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:CommandField ButtonType="Image" SelectImageUrl="~/Images/select.gif" 
                            SelectText="Выбрать" ShowSelectButton="True" FooterStyle-HorizontalAlign="NotSet" ItemStyle-HorizontalAlign="Center" />
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
    window.onload = load_product;
    </script>
</asp:Content>
