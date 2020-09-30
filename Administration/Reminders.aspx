<%@ Page Language="C#" MasterPageFile="~/Administration/Administration.Master" AutoEventWireup="true" CodeBehind="Reminders.aspx.cs" Inherits="CardPerso.Administration.Reminders" Title="Напоминания" %>
<asp:Content ID="RemindersContent" ContentPlaceHolderID="AdministrationContent" runat="server">
    <div id="pReminder" style="overflow: auto; height: 300px;">
    <asp:Repeater ID="rReceiveInFilial" runat="server">
        <HeaderTemplate>
            <h2>Прием ценностей в филиалах</h1>
            <table style="width:100%" border="1" cellpadding="1" cellspacing="1" >
                <tr>
                    <td colspan="2">
                        просроченное напоминание</td>
                    <td style="width:40%">список рассылки</td>
                    <td style="width:20%"></td>
                </tr>
        </HeaderTemplate>
        <FooterTemplate>
            </table>
        </FooterTemplate>
        <ItemTemplate>
            <tr>
            <td style="width:40px;border-right:0"></td>
            <td style="border-left:0"><%#Eval("Message") %></td>
            <td style="width:20%"><%#Eval("FIO")%></td>
            <td style="text-align:center"><asp:Button runat="server" OnClick="bSendMessage_Click" ID="bSendMessage" CommandArgument=<%#Eval("ToButton")%> Text="Отправить" Enabled=<%#Eval("Enbl")%> /></td>
            </tr>
        </ItemTemplate>
    </asp:Repeater>
    </div>
<script type="text/javascript">
    //window.onload = load_bank;
 </script>
</asp:Content>