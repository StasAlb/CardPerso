<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="UserEdit.aspx.cs" Inherits="CardPerso.Administration.UserEdit1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Редактирование пользователя</title>
    <meta http-equiv="Pragma" content="no-cache" />
    <base target="_self" />
    <style type="text/css">
        span
        {
           	color:#000080;  
        }
    </style>

</head>
<body style="margin: 0px; background-color: #F7F7DE;">
    <form id="form1" runat="server">
            <table border="0">
                <tr>
                    <td align="left" style="width:100%; border-width: thin; border-style: groove;" colspan="2">
                        <asp:ImageButton ID="bSave" runat="server" ImageUrl="~/Images/save.bmp" 
                            ToolTip="Сохранить" onclick="bSave_Click" />
                    </td>
                </tr>
                <tr>
                    <td align="left">
                            <asp:Label ID="UserNameLabel" runat="server" AssociatedControlID="UserName">Логин 
                            пользователя:</asp:Label>
                        </td>
                        <td>
                            <asp:TextBox ID="UserName" runat="server" Width="130px" Enabled="false"></asp:TextBox>
                            <asp:RequiredFieldValidator ID="UserNameRequired" runat="server" 
                                ControlToValidate="UserName" ErrorMessage="Требуется ввести логин пользователя" 
                                ToolTip="Логин пользователя" ValidationGroup="NewUserWizard">*</asp:RequiredFieldValidator>
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <asp:Label ID="UserLastNameLabel" runat="server" AssociatedControlID="UserLastName">Фамилия:</asp:Label>
                        </td>
                        <td>
                            <asp:TextBox ID="UserLastName" runat="server" Width="130px"></asp:TextBox>
                            <asp:RequiredFieldValidator ID="UserLastNameRequired" runat="server" 
                                ControlToValidate="UserLastName" ErrorMessage="Требуется ввести фамилию пользователя" 
                                ToolTip="Фамилия пользователя" ValidationGroup="NewUserWizard">*</asp:RequiredFieldValidator>
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <asp:Label ID="UserFirstNameLabel" runat="server" AssociatedControlID="UserFirstName">Имя:</asp:Label>
                        </td>
                        <td>
                            <asp:TextBox ID="UserFirstName" runat="server" Width="130px"></asp:TextBox>
                            <asp:RequiredFieldValidator ID="UserFirstNameRequired" runat="server" 
                                ControlToValidate="UserFirstName" ErrorMessage="Требуется ввести имя пользователя" 
                                ToolTip="Имя пользователя" ValidationGroup="NewUserWizard">*</asp:RequiredFieldValidator>
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <asp:Label ID="UserSecondNameLabel" runat="server" AssociatedControlID="UserSecondName">Отчество:</asp:Label>
                        </td>
                        <td>
                            <asp:TextBox ID="UserSecondName" runat="server" Width="130px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <asp:Label ID="UserPositionLabel" runat="server" AssociatedControlID="UserPosition">Должность:</asp:Label>
                        </td>
                        <td>
                            <asp:TextBox ID="UserPosition" runat="server" Width="130px"></asp:TextBox>
                            <asp:RequiredFieldValidator ID="UserPositionRequired" runat="server" 
                                ControlToValidate="UserPosition" ErrorMessage="Требуется ввести должность пользователя" 
                                ToolTip="Должность пользователя" ValidationGroup="NewUserWizard">*</asp:RequiredFieldValidator>
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <asp:Label ID="EmailLabel" runat="server" AssociatedControlID="Email">E-mail:</asp:Label>
                        </td>
                        <td>
                            <asp:TextBox ID="Email" runat="server" Width="130px"></asp:TextBox>
                            <%/*
                            <asp:RequiredFieldValidator ID="EmailRequired" runat="server" 
                                ControlToValidate="Email" ErrorMessage="Введите E-mail" 
                                ToolTip="E-mail пользователя" ValidationGroup="NewUserWizard">*</asp:RequiredFieldValidator>
                            */
                             %>    
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <asp:Label ID="BranchLabel" runat="server" AssociatedControlID="BranchDDL">Подразделение</asp:Label>
                        </td>
                        <td>
                            <asp:DropDownList ID="BranchDDL" runat="server"></asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td  align="left">
                            <asp:Label ID="SeriesLabel" runat="server" AssociatedControlID="tbPassportSeries">Серия паспорта:</asp:Label>
                        </td>
                        <td>
                            <asp:TextBox ID="tbPassportSeries" runat="server"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td  align="left"> 
                            <asp:Label ID="NumberLabel" runat="server" AssociatedControlID="tbPassportNumber">Номер паспорта:</asp:Label>
                        </td>
                        <td>
                            <asp:TextBox ID="tbPassportNumber" runat="server"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="center" colspan="2" style="color:Red;">
                            <asp:Literal ID="ErrorMessage" runat="server" EnableViewState="False"></asp:Literal>
                        </td>
                    </tr>
                </table>
        </form>
    </body>
</html>
