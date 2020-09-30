<%@ Page Language="C#" MasterPageFile="~/Catalogs.Master" AutoEventWireup="true" CodeBehind="AccountablePerson.aspx.cs" Inherits="CardPerso.AccountablePerson" Title="Untitled Page" %>
<asp:Content ID="AccountablePersonHeader" ContentPlaceHolderID="CatalogHeaderContent" runat="server">
    <title>Подотчетные лица</title>
</asp:Content>
<asp:Content ID="AccountablePersonContent" ContentPlaceHolderID="CatalogContent" runat="server">
    <table width="100%">
            <tr>
            <td>
                <asp:ImageButton ID="bNew" runat="server" ImageUrl="~/Images/new.gif" 
                    ToolTip="Новый" onclick="bNew_Click" />
                <asp:ImageButton ID="bEdit" runat="server" ImageUrl="~/Images/edit.bmp" 
                     ToolTip="Редактировать" onclientclick="return clickEdit();"/>
                <asp:ImageButton ID="bDelete" runat="server" ImageUrl="~/Images/del.bmp" 
                    ToolTip="Удалить" onclientclick="return clickDel();"/>
                <asp:ImageButton ID="bExcel" runat="server" ImageUrl="~/Images/excel.bmp" 
                     ToolTip="Вывод в Excel" onclick="bExcel_Click"/>  
                
            </td>
            <td align="right">
                <asp:Label ID="lbInform" runat="server" ForeColor="Red"></asp:Label>
                <asp:Label ID="lbCount" runat="server" Font-Bold="True" ForeColor="#400080"></asp:Label>
            </td>
            </tr>
            <!--
            <tr>
                <td>
                <table cellpadding="2px" cellspacing="5px" width="100%" border="0">
                <tr>
                <td>
                    <asp:Label ID="fi_secondnameLabel" runat="server" Text="Фамилия:"></asp:Label>
                </td>
                <td>
			   	    <asp:TextBox ID="fi_secondname" runat="server" class="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			    </td>
    			
			    <td>
                    <asp:Label ID="fi_firstnameLabel" runat="server" Text="Имя:"></asp:Label>
                </td>
                <td>
			   	    <asp:TextBox ID="fi_firstname" runat="server" class="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			    </td>
			    <td>

                <td>
                    <asp:Label ID="fi_positionLabel" runat="server" Text="Должность:"></asp:Label>
                </td>
                <td>
			   	    <asp:TextBox ID="fi_position" runat="server" class="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			    </td>
			    <td>
                    <asp:Label ID="fi_personnelnumberLabel" runat="server" Text="Табельный номер:"></asp:Label>
                </td>
                <td>
			   	    <asp:TextBox ID="fi_personnelnumber" runat="server" class="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			    </td>
			    </tr>
			    </table>
			    </td>
            </tr>
            -->
            <tr>
             <td colspan="2">
                <div id="pAccountablePerson" 
                     style="overflow: auto; height: 300px;">
                <asp:GridView ID="gvAccountablePersons" runat="server" AutoGenerateColumns="False" width="97%"  
                    BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" 
                    CellPadding="3" 
                    DataKeyNames="id,fio" 
                        onselectedindexchanged="gvAccountablePersons_SelectedIndexChanged"> 
                    <FooterStyle BackColor="White" ForeColor="#000066" />
                    <RowStyle ForeColor="#000066" />
                    <Columns>
                        <asp:BoundField DataField="fio" HeaderText="ФИО" >
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="UserName" HeaderText="Логин">
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:CommandField ButtonType="Image" SelectImageUrl="~/Images/select.gif" 
                            SelectText="Выбрать" ShowSelectButton="True" ItemStyle-HorizontalAlign="Center" />
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
    
    <div id="popup_accountableperson" title=""  class="ui-widget" style="display:none">
		
		<div style="display:none">
            <asp:Button ID="bNewAccountablePerson" runat="server" Text="1" onclick="bNewAccountablePerson_Click"/>
		    <asp:Button ID="bSaveAccountablePerson" runat="server" Text="2" onclick="bSaveAccountablePerson_Click" />       
            <asp:Button ID="bCancelAccountablePerson"  runat="server" Text="3" onclick="bCancelAccountablePerson_Click" />
            <asp:Button ID="bDelAccountablePerson"  runat="server" Text="4"  onclick="bDelAccountablePerson_Click"/>
		</div>
		
		<table cellpadding="2px" cellspacing="5px" width="100%" border="0">
		<tr>
			<td width="130px">
                <asp:Label ID="secondnameLabel" runat="server" Text="Фамилия:"></asp:Label>
                <font color="red">*</font>
            </td>
            <td>
			   	<asp:TextBox ID="secondname" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td width="50px">
                <asp:Label ID="firstnameLabel" runat="server" Text="Имя:"></asp:Label>
                <font color="red">*</font>
            </td>
            <td>
			   	<asp:TextBox ID="firstname" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="patronymicLabel" runat="server" Text="Отчество:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="patronymic" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
        <tr>
            <td>
                <asp:Label ID="passportLabel" runat="server" Text="Паспорт:<br>(формат 1234 123456)"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="passport" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
            <td>
                <asp:Label ID="loginLabel" runat="server" Text="Логин:"></asp:Label>
                <font color="red">*</font>
            </td>
            <td>
                <asp:TextBox ID="userLogin" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%" ></asp:TextBox>
            </td>
            <td />
            <td />
        </tr>

      	</table> 
      	
      	<div id="accordion_accountableperson">
        <h3>Счета по картам MASTERCARD</h3>
        <div>
        <table cellpadding="2px" cellspacing="5px" width="100%">
		<tr>	
			<td>
                <b><asp:Label ID="MasterCard_In_False_Label" runat="server" Text="Выдача из хранилища"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="MasterCard_In_False_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MasterCard_In_False_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="MasterCard_In_False_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MasterCard_In_False_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="MasterCard_In_True_Label" runat="server" Text="Выдача из сейфа"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="MasterCard_In_True_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MasterCard_In_True_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="MasterCard_In_True_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MasterCard_In_True_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="MasterCard_Out_False_Label" runat="server" Text="Выдача клиенту"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="MasterCard_Out_False_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MasterCard_Out_False_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="MasterCard_Out_False_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MasterCard_Out_False_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="MasterCard_Return_False_Label" runat="server" Text="Возврат в хранилище"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="MasterCard_Return_False_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MasterCard_Return_False_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="MasterCard_Return_False_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MasterCard_Return_False_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="MasterCard_Return_True_Label" runat="server" Text="Возврат в сейф"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="MasterCard_Return_True_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MasterCard_Return_True_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="MasterCard_Return_True_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MasterCard_Return_True_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		</table>
        </div>
        <h3>Счета по картам VISA</h3>
        <div>
        <table cellpadding="2px" cellspacing="5px" width="100%">
		<tr>	
			<td>
                <b><asp:Label ID="VisaCard_In_False_Label" runat="server" Text="Выдача из хранилища"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="VisaCard_In_False_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="VisaCard_In_False_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="VisaCard_In_False_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="VisaCard_In_False_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="VisaCard_In_True_Label" runat="server" Text="Выдача из сейфа"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="VisaCard_In_True_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="VisaCard_In_True_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="VisaCard_In_True_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="VisaCard_In_True_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="VisaCard_Out_False_Label" runat="server" Text="Выдача клиенту"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="VisaCard_Out_False_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="VisaCard_Out_False_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="VisaCard_Out_False_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="VisaCard_Out_False_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="VisaCard_Return_False_Label" runat="server" Text="Возврат в хранилище"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="VisaCard_Return_False_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="VisaCard_Return_False_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="VisaCard_Return_False_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="VisaCard_Return_False_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="VisaCard_Return_True_Label" runat="server" Text="Возврат в сейф"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="VisaCard_Return_True_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="VisaCard_Return_True_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="VisaCard_Return_True_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="VisaCard_Return_True_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		</table>
        </div>
        <h3>Счета по картам МИР</h3>
        <div>
        <table cellpadding="2px" cellspacing="5px" width="100%">
		<tr>	
			<td>
                <b><asp:Label ID="MirCard_In_False_Label" runat="server" Text="Выдача из хранилища"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="MirCard_In_False_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MirCard_In_False_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="MirCard_In_False_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MirCard_In_False_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="MirCard_In_True_Label" runat="server" Text="Выдача из сейфа"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="MirCard_In_True_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MirCard_In_True_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="MirCard_In_True_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MirCard_In_True_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="MirCard_Out_False_Label" runat="server" Text="Выдача клиенту"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="MirCard_Out_False_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MirCard_Out_False_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="MirCard_Out_False_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MirCard_Out_False_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="MirCard_Return_False_Label" runat="server" Text="Возврат в хранилище"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="MirCard_Return_False_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MirCard_Return_False_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="MirCard_Return_False_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MirCard_Return_False_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="MirCard_Return_True_Label" runat="server" Text="Возврат в сейф"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="MirCard_Return_True_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MirCard_Return_True_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="MirCard_Return_True_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="MirCard_Return_True_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		</table>
        </div>
     	<h3>Счета по картам NFC</h3>
        <div>
        <table cellpadding="2px" cellspacing="5px" width="100%">
		<tr>	
			<td>
                <b><asp:Label ID="NFCCard_In_False_Label" runat="server" Text="Выдача из хранилища"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="NFCCard_In_False_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="NFCCard_In_False_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="NFCCard_In_False_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="NFCCard_In_False_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="NFCCard_In_True_Label" runat="server" Text="Выдача из сейфа"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="NFCCard_In_True_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="NFCCard_In_True_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="NFCCard_In_True_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="NFCCard_In_True_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="NFCCard_Out_False_Label" runat="server" Text="Выдача клиенту"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="NFCCard_Out_False_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="NFCCard_Out_False_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="NFCCard_Out_False_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="NFCCard_Out_False_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="NFCCard_Return_False_Label" runat="server" Text="Возврат в хранилище"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="NFCCard_Return_False_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="NFCCard_Return_False_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="NFCCard_Return_False_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="NFCCard_Return_False_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="NFCCard_Return_True_Label" runat="server" Text="Возврат в сейф"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="NFCCard_Return_True_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="NFCCard_Return_True_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="NFCCard_Return_True_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="NFCCard_Return_True_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		</table>
        </div>
        <h3>Счета по сервисным картам</h3>
        <div>
        <table cellpadding="2px" cellspacing="5px" width="100%">
		<tr>	
			<td>
                <b><asp:Label ID="ServiceCard_In_False_Label" runat="server" Text="Выдача из хранилища"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="ServiceCard_In_False_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="ServiceCard_In_False_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="ServiceCard_In_False_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="ServiceCard_In_False_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="ServiceCard_In_True_Label" runat="server" Text="Выдача из сейфа"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="ServiceCard_In_True_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="ServiceCard_In_True_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="ServiceCard_In_True_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="ServiceCard_In_True_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="ServiceCard_Out_False_Label" runat="server" Text="Выдача клиенту"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="ServiceCard_Out_False_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="ServiceCard_Out_False_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="ServiceCard_Out_False_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="ServiceCard_Out_False_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="ServiceCard_Return_False_Label" runat="server" Text="Возврат в хранилище"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="ServiceCard_Return_False_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="ServiceCard_Return_False_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="ServiceCard_Return_False_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="ServiceCard_Return_False_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="ServiceCard_Return_True_Label" runat="server" Text="Возврат в сейф"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="ServiceCard_Return_True_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="ServiceCard_Return_True_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="ServiceCard_Return_True_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="ServiceCard_Return_True_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		</table>
        </div>
        <h3>Счета по ПИН-конвертам</h3>
        <div>
        <table cellpadding="2px" cellspacing="5px" width="100%">
		<tr>	
			<td>
                <b><asp:Label ID="PinConvert_In_False_Label" runat="server" Text="Выдача из хранилища"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="PinConvert_In_False_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="PinConvert_In_False_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="PinConvert_In_False_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="PinConvert_In_False_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="PinConvert_In_True_Label" runat="server" Text="Выдача из сейфа"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="PinConvert_In_True_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="PinConvert_In_True_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="PinConvert_In_True_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="PinConvert_In_True_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="PinConvert_Out_False_Label" runat="server" Text="Выдача клиенту"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="PinConvert_Out_False_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="PinConvert_Out_False_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="PinConvert_Out_False_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="PinConvert_Out_False_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="PinConvert_Return_False_Label" runat="server" Text="Возврат в хранилище"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="PinConvert_Return_False_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="PinConvert_Return_False_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="PinConvert_Return_False_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="PinConvert_Return_False_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		<tr>	
			<td>
                <b><asp:Label ID="PinConvert_Return_True_Label" runat="server" Text="Возврат в сейф"></asp:Label></b>
            </td>
			<td>
                <asp:Label ID="PinConvert_Return_True_Debet_Label" runat="server" Text="Дебет:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="PinConvert_Return_True_Debet" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
			<td>
                <asp:Label ID="PinConvert_Return_True_Credit_Label" runat="server" Text="Кредит:"></asp:Label>
            </td>
            <td>
			   	<asp:TextBox ID="PinConvert_Return_True_Credit" runat="server" CssClass="ui-widget-content ui-corner-all" style="width:100%"></asp:TextBox>
			</td>
		</tr>
		</table>
        </div>
    </div>

 		
    </div>
    
    <script type="text/javascript">
   
   var bNewAccountablePerson="#<%=bNewAccountablePerson.ClientID%>";
   var bSaveAccountablePerson="#<%=bSaveAccountablePerson.ClientID%>";
   var bCancelAccountablePerson="#<%=bCancelAccountablePerson.ClientID%>";
   var bDelAccountablePerson="#<%=bDelAccountablePerson.ClientID%>";
   var fio ='<%=((gvAccountablePersons.Rows.Count > 0) ? gvAccountablePersons.DataKeys[gvAccountablePersons.SelectedIndex].Values["fio"]:"")%>';
   
   function clickNew()
   {
        ShowEnterAccountablePerson("Новый",function f1() { ShowWait(); $(bNewAccountablePerson).click();},function f2() {$(bCancelAccountablePerson).click();});
        return false;
   }
   
   function clickEdit()
   {
        ShowEnterAccountablePerson("Редактирование",function f1() { ShowWait(); $(bSaveAccountablePerson).click();},function f2() {$(bCancelAccountablePerson).click();});
        return false;
   }
   
   function clickDel()
   {
        ShowQuestion("Удаление", "Вы действительно хотите удалить подотчетное лицо " + fio + "?", function f() { ShowWait(); $(bDelAccountablePerson).click();}, null);
        return false;
   }
   
   function setAccountActive(num)
   {
        $("#accordion_accountableperson").accordion( "option", "active", num );
   }
   
   
   function ShowEnterAccountablePerson(title,onOK,onClose)
    {
	    HideWait();
	    //$("#accordion_accountableperson").accordion({heightStyle: "content", animate: false});
        $("#accordion_accountableperson").css("display","none");

	    $("#popup_accountableperson").dialog(
	    {
	    title:title,
	    resizable:false,
	    modal:true,
	    width:700,
	    close: function(event, ui) {  $(this).dialog("destroy"); if(onClose!=null) onClose(); return true; },
	    buttons:{
      			    Сохранить: function()
      			    {
      				    onClose=null;
  					    $(this).dialog("close");
  					    onOK();
  					    return;
  				    },
				    Отмена: function()
				    {
					    $(this).dialog("close");
					    return;
				    }
			    }
	    });
	    $("#popup_accountableperson").keydown(function (event)     {
            if (event.keyCode == 13) {
                $(this).parent().find("button:eq(0)").trigger("click");
                return false;
            }
        });
	    
	   $("#popup_accountableperson").focus();
	
    }
    </script>
    
</asp:Content>
