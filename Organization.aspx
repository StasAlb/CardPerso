<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/Catalogs.Master" CodeBehind="Organization.aspx.cs" Inherits="CardPerso.Organization" %>
<asp:Content runat="server" ID="OrgTitle" ContentPlaceHolderID="CatalogHeaderContent">
    <title>Организации</title>
</asp:Content>
<asp:Content runat="server" ID="Org" ContentPlaceHolderID="CatalogContent">
    <table width="100%">
            <tr>
            <td>
                <asp:ImageButton ID="bNew" runat="server" ImageUrl="~/Images/new.gif" 
                    ToolTip="Новая организация" OnClientClick="return show_org('mode=1');" onclick="bNew_Click" />
                <asp:ImageButton ID="bEdit" runat="server" ImageUrl="~/Images/edit.bmp" 
                     ToolTip="Редактировать организацию" onclick="bEdit_Click"/>
                <asp:ImageButton ID="bDelete" runat="server" ImageUrl="~/Images/del.bmp" 
                    ToolTip="Удалить организацию" onclick="bDelete_Click" />
                <asp:ImageButton ID="bNewP" runat="server" ImageUrl="~/Images/in.bmp"
                    ToolTip="Добавить сотрудника" OnClick="bNewP_Click" />
                <asp:ImageButton ID="bEditP" runat="server" ImageUrl="~/Images/sfil.bmp"
                    ToolTip="Редактировать сотрудника" OnClick="bEditP_Click" />
                <asp:ImageButton ID="bDelP" runat="server" ImageUrl="~/Images/out.bmp"
                    ToolTip="Удалить сотрудника" OnClick="bDelP_Click" />
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
                <div id="pOrg" 
                     style="overflow: auto; height: 300px;">
                <asp:GridView ID="gvOrg" runat="server" AutoGenerateColumns="False" width="97%"  
                    BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" 
                    CellPadding="3" 
                        DataKeyNames="idP,idO,Title,EmbossTitle,Person,Position,Warrent,WStart,WEnd" 
                        onselectedindexchanged="gvOrg_SelectedIndexChanged">
                    <FooterStyle BackColor="White" ForeColor="#000066" />
                    <RowStyle ForeColor="#000066" />
                    <Columns>
                        <asp:BoundField DataField="title" HeaderText="Организация" >
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="EmbossTitle" HeaderText="Эмбоссировано" >
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Person" HeaderText="Доверенное лицо" >
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Position" HeaderText="Должность" >
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:BoundField DataField="WarrentS" HeaderText="Доверенность" >
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
    window.onload = load_organization;
    </script>
</asp:Content>
