<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="StorDocEdit.aspx.cs" Inherits="CardPerso.StorDocEdit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Редактирование</title>
    <meta http-equiv="Pragma" content="no-cache" />
    <base target="_self" />
    <style type="text/css">
        span
        {
/*        	font-weight:bold; */
           	color:#000080;  
        }
    </style>
    <meta http-equiv="X-UA-Compatible" content="IE=edge;" />
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <meta http-equiv="Pragma" content="no-cache" />
    <link href="~/Styles/Site.css" rel="stylesheet" type="text/css" />
    
    
    <script type="text/javascript" src="Dialog.js"></script>
    
    <script type="text/javascript" src="<%=Page.ResolveClientUrl("~/javascript/JSscriptTemp.js")%>"></script>  
    
    <script type="text/javascript" src="<%=Page.ResolveClientUrl("~/javascript/jquery-2.1.1.js")%>"></script>  
    <script type="text/javascript" src="<%=Page.ResolveClientUrl("~/javascript/jquery-migrate-1.2.1.js")%>"></script>  
  
    <script type="text/javascript" src="<%=Page.ResolveClientUrl("~/javascript/jquery-ui.js")%>"></script>  
    <script type="text/javascript" src="<%=Page.ResolveClientUrl("~/javascript/jquery-ui-i18n-rus.js")%>"></script>  
    <script type="text/javascript" src="<%=Page.ResolveClientUrl("~/javascript/servertable.js")%>"></script>  
    
    <script type="text/javascript" src="<%=Page.ResolveClientUrl("~/javascript/dialog.js")%>"></script>  
    
    <script type="text/javascript" src="<%=Page.ResolveClientUrl("~/javascript/globalize.js")%>"></script>  
    <script type="text/javascript" src="<%=Page.ResolveClientUrl("~/javascript/globalize.culture.de-DE.js")%>"></script>  
    
    <script type = "text/javascript">
        function DisableButton() {
            document.getElementById("<%=bSave.ClientID %>").disabled = true;
        }
        window.onbeforeunload = DisableButton;
    </script>
    <link href="<%=Page.ResolveClientUrl("~/css/jquery-ui.css")%>" rel="stylesheet" type="text/css"/>
</head>
<body style="margin: 0px; background-color: #F7F7DE;">
    <form id="form1" runat="server">
    
    <table width="100%">
            <tr>
                <td colspan="2" 
                    style="border-width: thin; border-style: groove;"> 
                     <asp:ImageButton ID="bSave" runat="server" ImageUrl="~/Images/save.bmp" 
                     ToolTip="Сохранить" OnClick="bSave_Click" />
                </td>
            </tr>
            <tr>
                <td style="width: 30%">
                    <asp:Label ID="Label8" runat="server" Text="Дата документа"></asp:Label>
                </td>
                <td style="width: 70%">
                    <asp:TextBox ID="tbData" runat="server" Width="40%" Enabled="False"></asp:TextBox>
                </td>
            </tr> 
            <tr>
                <td style="width: 30%">
                    <asp:Label ID="Label7" runat="server" Text="Тип документа"></asp:Label>
                </td>
                <td style="width: 70%">
                    <div class="ui-widget">
                     <asp:DropDownList ID="dListType" runat="server" Width="90%" 
                         AutoPostBack="True" 
                         onselectedindexchanged="dListType_SelectedIndexChanged"></asp:DropDownList>
                    </div>
                </td>
            </tr> 
            <asp:Panel ID="pReceiveType" runat="server" Enabled="false">
                <tr>
                    <td>
                        <asp:Label runat="server" Text="Тип приема"></asp:Label>
                    </td>
                    <td>
                        <asp:RadioButtonList ID="rbReceiveType" runat="server" ForeColor="#000080" RepeatDirection="Horizontal" AutoPostBack="True">
                            <asp:ListItem Value="0" Selected="True">Поштучно</asp:ListItem>
                            <asp:ListItem Value="1">Пакетно</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
            </asp:Panel>
            <asp:Panel ID="pTypeAct" runat="server">
                <tr>
                    <td>
                        <asp:Label ID="Label2" runat="server" Text="Действие"></asp:Label>
                    </td>
                    <td>
                        <asp:RadioButtonList onselectedindexchanged="rbType_SelectedIndexChanged" ID="rbType" runat="server" ForeColor="#000080" RepeatDirection="Horizontal" AutoPostBack="True">
                            <asp:ListItem Value="0">Филиал</asp:ListItem>
                            <asp:ListItem Value="1">Рассылка</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
            </asp:Panel>
            <asp:Panel ID="pBranch" runat="server">
            <tr>
                <td>
                    <asp:Label ID="Label9" runat="server" Text="Филиал"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="dListBranch" runat="server" Width="90%" AutoPostBack="True"
                    onselectedindexchanged="dListBranch_SelectedIndexChanged"></asp:DropDownList>
                </td>
            </tr>
            </asp:Panel>
            <asp:Panel ID="pAct" runat="server">
            <tr>
                <td>
                    <asp:Label ID="Label4" runat="server" Text="Документ"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="dListAct" OnSelectedIndexChanged="dListAct_SelectedIndexChanged" AutoPostBack="true"  runat="server" Width="90%"></asp:DropDownList>
                </td>
            </tr>
            </asp:Panel>
            <asp:Panel ID="pDeliver" runat="server">
            <tr>
                <td>
                    <asp:Label ID="Label1" runat="server" Text="Рассылка"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="dListDeliver" runat="server" Width="90%"></asp:DropDownList>
                </td>
            </tr>
            </asp:Panel>
            
            <asp:Panel ID="pCourier" runat="server">
            <tr>
                <td>
                    <asp:Label ID="Label10" runat="server" Text="Курьерская служба"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="dListCr" runat="server" Width="90%"></asp:DropDownList>
                </td>
            </tr>
            <tr>    
                <td>
                    <asp:Label ID="Label12" runat="server" Text="Номер накладной"></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbInvoiceCr" runat="server" Width="40%"></asp:TextBox>
                </td>
            </tr>
            <tr>    
                <td>
                    <asp:Label ID="Label3" runat="server" Text="Сотрудник"></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbCourier" runat="server" Width="89%"></asp:TextBox>
                </td>
            </tr>
            </asp:Panel>
            <asp:Panel ID="pFilFil" runat="server">
            <tr>
                <td>
                    <asp:Label ID="Label5" runat="server" Text="Филиал откуда"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="dFilFrom" runat="server" Width="90%" AutoPostBack="True"
                    onselectedindexchanged="dFilFrom_SelectedIndexChanged"></asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label6" runat="server" Text="Филиал куда"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="dFilTo" runat="server" Width="90%" AutoPostBack="True"
                    onselectedindexchanged="dFilTo_SelectedIndexChanged"></asp:DropDownList>
                </td>
            </tr>
            </asp:Panel>
        
        <asp:Panel runat="server" ID="pPassport">
        <tr>
            <td style="width: 30%">
                <asp:Label ID="Label13" runat="server" Text="Серия паспорта"></asp:Label>
            </td>
            <td style="width: 70%">
                <asp:TextBox ID="tbPassportSeries" runat="server" Width="89%"></asp:TextBox>		  
            </td>
        </tr>
        <tr>
            <td style="width: 30%">
                <asp:Label ID="Label14" runat="server" Text="Номер паспорта"></asp:Label>
            </td>
            <td style="width: 70%">
                <asp:TextBox ID="tbPassportNumber" runat="server" Width="89%"></asp:TextBox>
            </td>
        </tr>
        </asp:Panel>
        <asp:Panel ID="pAccountablePerson" runat="server">
            <tr>
                <td>
                    <asp:Label ID="lAccountablePerson" runat="server" Text="Подотчетное лицо"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="dListPerson" AutoPostBack="True" onselectedindexchanged="dListPerson_OnSelectedIndexChanged" runat="server" Width="90%"></asp:DropDownList>
                </td>
            </tr>
        </asp:Panel>
        <asp:Panel ID="pBook124" runat="server">
            <tr>
                <td>
                    <asp:Label ID="Label15" runat="server" Text="Ответственный работник"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="dListBook124Person" AutoPostBack="True" onselectedindexchanged="dListBook124Person_OnSelectedIndexChanged" runat="server" Width="90%"></asp:DropDownList>
                </td>
            </tr>
        </asp:Panel>

            <tr>
                <td style="width: 30%">
                    <asp:Label ID="Label11" runat="server" Text="Комментарий"></asp:Label>
                </td>
                <td style="width: 70%">
                    <asp:TextBox ID="tbComment" runat="server" Width="89%"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="right">
                    <asp:Label ID="lbInform" runat="server" ForeColor="#400080"></asp:Label>
                </td>
            </tr>
            <asp:Panel ID="pForKilling" runat="server">
            <tr>
                <td colspan="2" align="right">
                    <asp:CheckBox ID="cbForKilling" runat="server" Text="На уничтожение" />
                </td>
            </tr>
            </asp:Panel>

       </table>
    </form>
</body>
</html>
