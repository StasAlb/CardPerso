<%@ Page Title="Управление пользователями" Language="C#" AutoEventWireup="true" MasterPageFile="~/Administration/Administration.Master"CodeBehind="ManageUsers.aspx.cs" Inherits="CardPerso.Administration.ManageUser" %>
<asp:Content ID="ManageUserContent" runat="server" ContentPlaceHolderID="AdministrationContent">
    <table width="100%">
    <tr>
    <td>
        <asp:ImageButton ID="bNew" runat="server" ImageUrl="~/Images/new.gif" 
            ToolTip="Новый" OnClientClick="return show_user('mode=1');" onclick="bNew_Click" />
        <asp:ImageButton ID="bEdit" runat="server" ImageUrl="~/Images/edit.bmp" 
             ToolTip="Редактировать" onclick="bEdit_Click"   />
        <asp:ImageButton ID="bRoles" runat="server" ImageUrl="~/Images/field.bmp" 
             ToolTip="Роли" onclick="bRoles_Click" />
        <asp:ImageButton ID="bSetFilter" runat="server" ImageUrl="~/Images/sfil.bmp" 
             ToolTip="Фильтр"  OnClientClick="return show_userfilter();" 
             onclick="bSetFilter_Click" />
        <asp:ImageButton ID="bResetFilter" runat="server" ImageUrl="~/Images/rfil.bmp" 
             ToolTip="Снять фильтр" onclick="bResetFilter_Click" />                          
        <asp:ImageButton ID="bDelete" runat="server" ImageUrl="~/Images/del.bmp" 
            ToolTip="Не активный" onclick="bDelete_Click" />
        <asp:ImageButton ID="bActivate" runat="server" ImageUrl="~/Images/rfil.gif"
            ToolTip="Активировать" onclick="bActivate_Click" />
    </td>
    <td align="right">
        <asp:Label ID="lbInform" runat="server" ForeColor="Red"></asp:Label>
        <asp:Label ID="lbCount" runat="server" Font-Bold="True" ForeColor="#400080"></asp:Label>
    </td>
    </tr>
     <tr>
     <td colspan="2">
        <div id="pUsers"> 
             <!--style="overflow: auto; height: 300px;">-->
        
    <asp:GridView ID="gvUsers" runat="server" AutoGenerateColumns="False" width="97%"  
        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" 
        CellPadding="3" 
        DataKeyNames="UserId,UserName,FIO,Position,Branch" 
                onselectedindexchanged="gvUsers_SelectedIndexChanged" 
                onrowdatabound="gvUsers_RowDataBound">
        <FooterStyle BackColor="White" ForeColor="#000066" />
        <RowStyle ForeColor="#000066" />    
        <Columns>
            <asp:BoundField DataField="UserName" HeaderText="Логин" SortExpression="UserName">
                <ItemStyle HorizontalAlign="Left"/>
            </asp:BoundField>
            <asp:BoundField DataField="FIO" HeaderText="Фамилия, Имя Отчество" SortExpression="FIO">
                <ItemStyle HorizontalAlign="Left" />
            </asp:BoundField>
            <asp:BoundField DataField="Position" HeaderText="Должность" SortExpression="Position">
                <ItemStyle HorizontalAlign="Left" />
            </asp:BoundField>
            <asp:BoundField DataField="Branch" HeaderText="Отделение" SortExpression="Branch">
                <ItemStyle HorizontalAlign="Left" />
            </asp:BoundField>
            <asp:BoundField DataField="Roles" HeaderText="Роли" SortExpression="Roles">
                <ItemStyle HorizontalAlign="Left" />
            </asp:BoundField>
            <asp:CommandField ButtonType="Image" SelectImageUrl="~/Images/select.gif" 
                SelectText="Выбрать" ShowSelectButton="True" 
                FooterStyle-HorizontalAlign="NotSet" ItemStyle-HorizontalAlign="Center" >                        
<ItemStyle HorizontalAlign="Center"></ItemStyle>
            </asp:CommandField>
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
    <asp:Label ID="lbFilter" runat="server" Visible="false"></asp:Label>
</asp:Content>
