<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="UserAdd.aspx.cs" Inherits="CardPerso.Administration.UserEdit" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Создание нового пользователя</title>
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

    <asp:CreateUserWizard ID="NewUserWizard" runat="server" 
        CompleteSuccessText="Пользователь успешно создан" 
        oncreateduser="NewUserWizard_CreatedUser" LoginCreatedUser="False" 
        CreateUserButtonText="Создать" 
        DuplicateUserNameErrorMessage="Данный логин существует. Пожалуйста, измените." 
        ContinueButtonText="Продолжить" FinishCompleteButtonText="Сохранить" 
        FinishPreviousButtonText="" oncreatinguser="NewUserWizard_CreatingUser" 
        Width="480px" InvalidPasswordErrorMessage="Минимальная длина пароля {0}" 
        RequireEmail="False">
        <StepPreviousButtonStyle BorderStyle="None" />
    <WizardSteps>
        <asp:CreateUserWizardStep ID="NewUserWizardStep1" runat="server" 
            Title="Создание нового пользователя">
        <ContentTemplate>
            <table border="0">
                <tr>
                    <td align="center" colspan="2">
                        <h2>
                            Создание нового пользователя</h2></td>
                </tr>
                <tr>
                    <td align="left">
                            <asp:Label ID="UserNameLabel" runat="server" AssociatedControlID="UserName">Логин 
                            пользователя:</asp:Label>
                        </td>
                        <td>
                            <asp:TextBox ID="UserName" runat="server" Width="130px"></asp:TextBox>
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
                            <asp:TextBox ID="Email" runat="server" Width="130px" CausesValidation="false" ToolTip="E-mail пользователя"></asp:TextBox>
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
                        <td align="center" colspan="2">
                            <asp:CompareValidator ID="PasswordCompare" runat="server" 
                                ControlToCompare="Password" ControlToValidate="ConfirmPassword" 
                                Display="Dynamic" 
                                ErrorMessage="Пароль не подтвержден" 
                                ValidationGroup="NewUserWizard"></asp:CompareValidator>
                        </td>
                    </tr>
                    <tr>
                        <td align="center" colspan="2" style="color:Red;">
                            <asp:Literal ID="ErrorMessage" runat="server" EnableViewState="False"></asp:Literal>
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <asp:Label ID="PasswordLabel" runat="server" AssociatedControlID="Password" Visible="false">Пароль:</asp:Label>
                        </td>
                        <td>
                            <asp:TextBox ID="Password" runat="server" Text="123" TextMode="Password" Width="130px" Visible="false"></asp:TextBox>
                            <asp:RequiredFieldValidator ID="PasswordRequired" runat="server" 
                                ControlToValidate="Password" ErrorMessage="Введите пароль" 
                                ToolTip="Введите пароль" ValidationGroup="NewUserWizard">*</asp:RequiredFieldValidator>
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <asp:Label ID="ConfirmPasswordLabel" runat="server" Visible="false" 
                                AssociatedControlID="ConfirmPassword">Подтверждение пароля:</asp:Label>
                        </td>
                        <td>
                            <asp:TextBox ID="ConfirmPassword" runat="server" Text="123" TextMode="Password" Visible="false"
                                Width="130px"></asp:TextBox>
                            <asp:RequiredFieldValidator ID="ConfirmPasswordRequired" runat="server" 
                                ControlToValidate="ConfirmPassword" 
                                ErrorMessage="Требуется подтверждение пароля" 
                                ToolTip="Требуется подтверждение пароля" ValidationGroup="NewUserWizard">*</asp:RequiredFieldValidator>
                        </td>
                    </tr>                    
                </table>
                </ContentTemplate>
        </asp:CreateUserWizardStep>
        
    <asp:CompleteWizardStep runat="server">
    <ContentTemplate>
        <table border="0" width="100%">
            <tr>
                <td align="center" colspan="3">
                    <h2>
                        Роль пользователя</h2></td>
            </tr>
            <tr>
                <td style="width:20%"></td>
                <td>
                    <asp:CheckBoxList AutoPostBack="false" runat="server" ID="cblRoles" Width="100%"></asp:CheckBoxList>
                </td>
                <td style="width:20%"></td>
            </tr>
            <tr><td colspan="3"></td></tr>
            <tr>
                <td></td>
                <td align="center">
                <asp:Button runat="server" ID="bRoles" Text="Сохранить" onclick="bRoles_Click" />
                </td>
                <td></td>
            </tr>
        </table>
    </ContentTemplate>
    </asp:CompleteWizardStep>
    </WizardSteps>
</asp:CreateUserWizard>
</form>
</body>
</html>
