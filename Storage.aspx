<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/CardPerso.Master" CodeBehind="Storage.aspx.cs" Inherits="CardPerso.Storage" %>
<asp:Content ID="StorageHeader" runat="server" ContentPlaceHolderID="HeadContent">
    <title>Хранилище</title>
    <script type="text/javascript" src="hint.js"></script>
</asp:Content>
<asp:Content ID="StorageContent" runat="server" ContentPlaceHolderID="MainContent">
     <table width="100%" cellpadding="0" cellspacing="0">
         <tr>
            <td>
                <asp:ImageButton ID="bSetFilter" runat="server" ImageUrl="~/Images/sfil.bmp" 
                     ToolTip="Фильтр"  OnClientClick="return show_flt_storage();" onclick="bSetFilter_Click" />
                <asp:ImageButton ID="bResetFilter" runat="server" ImageUrl="~/Images/rfil.bmp" 
                     ToolTip="Снять фильтр" onclick="bResetFilter_Click" />
                <asp:ImageButton ID="bConfField" runat="server" ImageUrl="~/Images/field.bmp" 
                     ToolTip="Настройка полей"  
                    OnClientClick="return show_field('type=storage','320px');" onclick="bConfField_Click" />
                <asp:ImageButton ID="bExcel" runat="server" ImageUrl="~/Images/excel.bmp" 
                     ToolTip="Вывод в Excel" onclick="bExcel_Click" />
                <asp:CheckBox ID="showChild" text="подчиненные филиалы" runat="server" Visible="True" 
                AutoPostBack="True" OnCheckedChanged="showChild_Changed" ToolTip="Показать подчиненные филиалы"/>     
                
            </td>
            <td align="right">
                <asp:Label ID="lbCount" runat="server" Font-Bold="True" ForeColor="#400080"></asp:Label>
            </td>
        </tr>
        <tr>
            <td colspan="2">
                <div id="pStorage" style="overflow: auto; height: 300px;">
                <asp:GridView ID="gvStorage" Width="98%" runat="server" BackColor="White" 
                BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="3" 
                AutoGenerateColumns="False"  DataKeyNames="id" onsorting="gvStorage_Sorting" 
                        AllowSorting="True">
                <FooterStyle BackColor="White" ForeColor="#000066" />
                <RowStyle ForeColor="#000066" />
                <Columns>
                    <asp:TemplateField HeaderText="Наименование" SortExpression="name">
                        <ItemStyle HorizontalAlign="Left" />
                        <ItemTemplate>
                            <%#(string)Eval("name") %>
                            <DIV style="position:absolute;" id='<%#("hint" + (string)Eval("id").ToString()) %>'></DIV>
                            <!--<img style='visibility:<%# (((string)Eval("image")).Length==0)?"hidden":"visible"%>;' src="images/info.png" alt="" onmouseover="SetToolTip('<%#Eval("image") %>', '<%#("hint" + (string)Eval("id").ToString()) %>')" onmouseout="hideInfo(this, '<%#("hint" + (string)Eval("id").ToString()) %>')"/>-->                            
                            <label style='color:Red;visibility:<%# (((string)Eval("image")).Length==0)?"hidden":"visible"%>;' onclick="SetToolTip('<%#Eval("image") %>', '<%#("hint" + (string)Eval("id").ToString()) %>')" onmouseout="hideInfo(this, '<%#("hint" + (string)Eval("id").ToString()) %>')"> (фото)</label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:BoundField DataField="bank_name" HeaderText="Банк" 
                    SortExpression="bank_name">
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="bin" HeaderText="Бин" 
                    SortExpression="bin">
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="min_cnt" HeaderText="Минимум" 
                    SortExpression="min_cnt">
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="cnt_new" HeaderText="Новые" 
                    SortExpression="cnt_new">
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
<%--                    <asp:BoundField DataField="cnt_wrk" HeaderText="На руках" 
                    SortExpression="cnt_wrk">
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>--%>
                    <asp:BoundField DataField="cnt_perso" HeaderText="Персо" 
                    SortExpression="cnt_perso">
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="cnt_brak" HeaderText="Брак" 
                    SortExpression="cnt_brak">
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="cnt_expire" HeaderText="Истекшие"
                    SortExpression="cnt_expire">
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="cnt_notaskedfor" HeaderText="Невостребованные"
                    SortExpression="cnt_notaskedfor">
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
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
    
    <script type="text/javascript">
    window.onload = load_storage;
    </script>
</asp:Content>    
