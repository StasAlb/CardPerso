<%@ Page Title="Управление ролями" Language="C#" MasterPageFile="~/Administration/Administration.Master" AutoEventWireup="true" CodeBehind="ManageRoles.aspx.cs" Inherits="CardPerso.Administration.ManageRoles" %>
<asp:Content ID="ManageRolesContent" runat="server" ContentPlaceHolderID="AdministrationContent">
    <table width="100%">
    <tr>
    <td  style="width:400px">
        <asp:ImageButton ID="bNew" runat="server" ImageUrl="~/Images/new.gif" 
            ToolTip="Новая" OnClientClick="return show_role('mode=1');" onclick="bNew_Click" />
        <asp:ImageButton ID="bEdit" runat="server" ImageUrl="~/Images/edit.bmp" 
             ToolTip="Редактировать" onclick="bEdit_Click" />
        <asp:ImageButton ID="bDelete" runat="server" ImageUrl="~/Images/del.bmp" 
            ToolTip="Удалить" onclick="bDelete_Click" />
    </td>
    <td>
    <asp:Label ID="lAction" runat="server"></asp:Label>
    </td>
    </tr>    
     <tr>
     <td>
        <div id="pRoles" 
             style="overflow: auto; height: 450px;">
        
    <asp:GridView ID="gvRoles" runat="server" AutoGenerateColumns="False" width="97%"  
        BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" 
        CellPadding="3" 
        DataKeyNames="RoleId,RoleName,UserCnt"
                onselectedindexchanged="gvRoles_SelectedIndexChanged">
        <FooterStyle BackColor="White" ForeColor="#000066" />
        <RowStyle ForeColor="#000066" />    
        <Columns>
            <asp:BoundField DataField="RoleName" HeaderText="Название роли">
                <ItemStyle HorizontalAlign="Left" />
            </asp:BoundField>
            <asp:BoundField DataField="UserCnt" HeaderText="Пользователи">
                <ItemStyle HorizontalAlign="Left" />
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
    <td valign="top">
        <div id="pActions"
            style="overflow:auto; height: 650px;">
        
        <asp:CheckBoxList ID="clbAction" runat="server" DataTextField="ActionName" 
                DataValueField="id"
                onselectedindexchanged="clbAction_SelectedIndexChanged" 
                AutoPostBack="True"></asp:CheckBoxList> 
        </div>
    </td>
    </tr>
    </table>
    <script type="text/javascript">
        window.onload = load_managerole;
    </script>
</asp:Content>
