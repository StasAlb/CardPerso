<%@ Page Title="Выполнить вход" Language="C#" MasterPageFile="~/CardPerso.master" AutoEventWireup="true"
    CodeBehind="Login.aspx.cs" Inherits="CardPerso.Account.Login" %>
<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">

<script type="text/javascript">
    function testFunction()
    {
        /*
        var req=[];
        req.push({ name: "user", value: "231231232132"});
        ajaxJson("Account/Login.aspx/getBranchesForUser", req, alert, alert, true);
        */
    }        
</script>

</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <h2 style="text-align:center">
        Выполнить вход
    </h2>
    <table width="100%"><tr><td align="center">
    <asp:Login ID="LoginUser" runat="server" EnableViewState="true"  
        FailureText="Данный логин/пароль не найден" 
            onauthenticate="LoginUser_Authenticate" onloggingin="LoginUser_LoggingIn" 
            onloggedin="LoginUser_LoggedIn">
        <LayoutTemplate>
            <span class="failureNotification" style="text-align:center">
                <asp:Literal ID="FailureText" runat="server"></asp:Literal>
            </span>
            <asp:ValidationSummary ID="LoginUserValidationSummary" runat="server" CssClass="failureNotification" 
                 ValidationGroup="LoginUserValidationGroup"/>
            <div>
                <fieldset class="login">
                    <legend>Сведения учетной записи</legend>
                    <p>
                        <asp:Label ID="UserNameLabel" runat="server" AssociatedControlID="UserName">Имя 
                        пользователя:</asp:Label>
                        <asp:TextBox ID="UserName" runat="server" CssClass="textEntry"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="UserNameRequired" runat="server" ControlToValidate="UserName" 
                             CssClass="failureNotification" ErrorMessage="Поле ''Имя пользователя'' является обязательным." ToolTip="Поле ''Имя пользователя'' является обязательным." 
                             ValidationGroup="LoginUserValidationGroup">*</asp:RequiredFieldValidator>
                    </p>
                    <p>
                        <asp:Label ID="PasswordLabel" runat="server" AssociatedControlID="Password">Пароль:</asp:Label>
                        <asp:TextBox ID="Password" runat="server" CssClass="passwordEntry" TextMode="Password"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="PasswordRequired" runat="server" ControlToValidate="Password" 
                             CssClass="failureNotification" ErrorMessage="Поле ''Пароль'' является обязательным." ToolTip="Поле ''Пароль'' является обязательным." 
                             ValidationGroup="LoginUserValidationGroup">*</asp:RequiredFieldValidator>
                    </p>
                    <p>
                        <asp:Label ID="BranchLabel" runat="server" AssociatedControlID="Branch">Подразделение:</asp:Label>
                        <asp:DropDownList ID="Branch" runat="server" style="width:320px;border: 1px solid #ccc;"></asp:DropDownList>
                    </p>
                </fieldset>
                <p class="submitButton">
                    <asp:Button ID="LoginButton" runat="server" CommandName="Login" Text="Выполнить вход" ValidationGroup="LoginUserValidationGroup"/>
                </p>
            </div>
        </LayoutTemplate>
    </asp:Login>
    </td></tr></table>
    <!--<label onclick="testFunction();">testFunction</label>-->
</asp:Content>
