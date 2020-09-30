<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/Catalogs.Master" CodeBehind="Branch.aspx.cs" Inherits="CardPerso.Branch" %>
<asp:Content ID="BranchHeader" runat="server" ContentPlaceHolderID="CatalogHeaderContent">
    <title>Подразделения</title>
</asp:Content>
<asp:Content ID="BranchContent" runat="server" ContentPlaceHolderID="CatalogContent">
<%if (isMainFilial() == true || bUseMainBranch.Checked == true || bUseShift.Checked == true) {%> 
<script type="text/javascript">

  var oprdaystart="#<%=oprdaystart%>";
  var oprdayend="#<%=oprdayend%>";
  var shift1s="#<%=shift1s.ClientID%>";
  var shift1e="#<%=shift1e.ClientID%>";
  var shift2s="#<%=shift2s.ClientID%>";
  var shift2e="#<%=shift2e.ClientID%>";

  $.widget( "ui.timespinner", $.ui.spinner, 
  {
    options: 
    {
      // seconds
      step: 60 * 1000,
      // hours
      page: 60
    },
    _parse: function( value ) 
    {
      if ( typeof value === "string" ) 
      {
        // already a timestamp
        if ( Number( value ) == value )
        {
          return Number( value );
        }
        var g=+Globalize.parseDate( value );
        return g;
      }
      return value;
    },
    _format: function( value ) 
    {
        return Globalize.format( new Date(value), "t" ); 
    }
  });
  
  $(function() 
  {
    Globalize.culture("de-DE");
  });
</script>  
<%}%>  

<%if(isMainFilial()==true || bUseMainBranch.Checked==true) {%>  
<script type="text/javascript">
  $(function() 
  {
    $(oprdaystart).timespinner();
    $(oprdayend).timespinner();
  });
</script>
<%}%>    

<%if(bUseShift.Checked==true) {%>  
<script type="text/javascript">
  $(function() 
  {
    $(shift1s).timespinner();
    $(shift1e).timespinner();
    $(shift2s).timespinner();
    $(shift2e).timespinner();
  });
</script>
<%}%>    

    <table width="100%">
            <tr>
            <td>
                <%if(isAccountSave()==true) {%>    
                <asp:ImageButton ID="acountNumberSave" runat="server" ImageUrl="~/Images/save.bmp" ToolTip="Сохранить счета" onclick="bAcountNumberSave_Click"/>
                <%}%>    
                <asp:ImageButton ID="bNew" runat="server" ImageUrl="~/Images/new.gif" 
                    ToolTip="Новый" OnClientClick="return show_branch('mode=1');" onclick="bNew_Click"/>
                <asp:ImageButton ID="bEdit" runat="server" ImageUrl="~/Images/edit.bmp" 
                     ToolTip="Редактировать" onclick="bEdit_Click"/>
                <asp:ImageButton ID="bDelete" runat="server" ImageUrl="~/Images/del.bmp" 
                    ToolTip="Удалить" onclick="bDelete_Click"/>
                <asp:ImageButton ID="bAddOffice" runat="server" ImageUrl="~/Images/in.bmp" 
                     ToolTip="Привязать офис" onclick="bAddOffice_Click" /> 
                <asp:ImageButton ID="bDelOffice" runat="server" ImageUrl="~/Images/out.bmp" 
                     ToolTip="Отвязать офис" onclick="bDelOffice_Click"/> 
                <asp:ImageButton ID="bExcel" runat="server" ImageUrl="~/Images/excel.bmp" 
                     ToolTip="Вывод в Excel" onclick="bExcel_Click"/>     
                
            </td>
            <td align="right">
                <asp:CheckBox ID="bShowAccount" runat="server" TextAlign="Right" AutoPostBack="true" 
                    Text="счета подразделений и опердень" oncheckedchanged="bShowAccount_CheckedChanged" ForeColor="#400080"/>  
                <asp:Label ID="lbInform" runat="server" ForeColor="Red"></asp:Label>
                <asp:Label ID="lbCount" runat="server" Font-Bold="True" ForeColor="#400080"></asp:Label>
            </td>
            </tr>
            
             <tr>
             <td colspan="2">
                <div id="pBranch" 
                     style="overflow: auto; height: 300px;">
                <asp:GridView ID="gvBranchs" runat="server" AutoGenerateColumns="False" width="97%"  
                    BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" 
                    CellPadding="3" 
                    DataKeyNames="id,id_parent,ident_bank,ident_dep,department,office" 
                    onselectedindexchanged="gvBranchs_SelectedIndexChanged">
                    <FooterStyle BackColor="White" ForeColor="#000066" />
                    <RowStyle ForeColor="#000066" />
                    <Columns>
                        <asp:BoundField DataField="ident_bank" HeaderText="Код банка" >
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ident_dep" 
                            HeaderText="Код подразделения" >
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:BoundField DataField="department" 
                            HeaderText="Подразделение" >
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="office" 
                            HeaderText="Офис" >
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
              </tr>
            
        </table>    
        <div id="acountTable" style="display:none">            
        <div style="height:30px"/>   
        <table border="0" width="100%" style="background-color: #F7F7DE;">
        <tr>
                
                <td style="background-color:#006699;color:white;font-weight:bold;">Счета подразделения (<%=getDepartment()%>)</td>
        </tr>
        <!--<tr>        
                <td style="border-width: thin; border-style: groove; width:100%">
                    
                </td>
         </tr>-->
         <tr>
                <td>
                    <div id="acountFields" style="overflow: auto;border:1px solid black;">
                        <table cellspacing="0" style="color:#000080;font-weight:normal" border="0" width="100%">
                            
                            <tr>
                                <td style="border-bottom: 1px solid black;"></td>
                                <td align="left" colspan="2" style="font-weight:bold;border-bottom: 1px solid black;">Поступление</td>
                                <td align="left" colspan="2" style="font-weight:bold;border-bottom: 1px solid black;">Выдача</td>
                                <td align="left" colspan="2" style="font-weight:bold;border-bottom: 1px solid black;">Возврат</td>
                            </tr>
                            <tr>
                                <td rowspan="2" style="font-weight:bold;border-bottom: 1px solid black;">Мастер Кард</td>
                                <td colspan="1">Дебет:</td>
                                <td><asp:TextBox ID="mastercardInDebet"  Width="135pt" runat="server"/></td>
                                <td colspan="1">Дебет:</td>
                                <td><asp:TextBox ID="mastercardOutDebet" Width="135pt" runat="server"/></td>
                                <td colspan="1">Дебет:</td>
                                <td><asp:TextBox ID="mastercardRetDebet" Width="135pt" runat="server"/></td>
                            </tr>
                            <tr>
                                <td colspan="1" style="border-bottom:1px solid black;">Кредит:</td>
                                <td style="border-bottom: 1px solid black;"><asp:TextBox ID="mastercardInCredit" Width="135pt" runat="server"/></td>
                                <td colspan="1" style="border-bottom: 1px solid black;">Кредит:</td>
                                <td style="border-bottom: 1px solid black;"><asp:TextBox ID="mastercardOutCredit" Width="135pt" runat="server"/></td>
                                <td colspan="1" style="border-bottom: 1px solid black;">Кредит:</td>
                                <td style="border-bottom: 1px solid black;"><asp:TextBox ID="mastercardRetCredit" Width="135pt" runat="server"/></td>
                            </tr>
                            
                            <tr>
                                <td rowspan="2" style="font-weight:bold;border-bottom:1px solid black;">Visa</td>
                                <td colspan="1">Дебет:</td>
                                <td><asp:TextBox ID="visaInDebet" Width="135pt" runat="server"/></td>
                                <td colspan="1">Дебет:</td>
                                <td><asp:TextBox ID="visaOutDebet" Width="135pt" runat="server"/></td>
                                <td colspan="1">Дебет:</td>
                                <td><asp:TextBox ID="visaRetDebet" Width="135pt" runat="server"/></td>
                            </tr>
                            <tr>
                                <td colspan="1" style="border-bottom:1px solid black;">Кредит:</td>
                                <td style="border-bottom:1px solid black;"><asp:TextBox ID="visaInCredit" Width="135pt" runat="server"/></td>
                                <td colspan="1" style="border-bottom:1px solid black;">Кредит:</td>
                                <td style="border-bottom:1px solid black;"><asp:TextBox ID="visaOutCredit" Width="135pt" runat="server"/></td>
                                <td colspan="1" style="border-bottom:1px solid black;">Кредит:</td>
                                <td style="border-bottom:1px solid black;"><asp:TextBox ID="visaRetCredit" Width="135pt" runat="server"/></td>
                            </tr>
                            
                            <tr>
                                <td rowspan="2" style="font-weight:bold;border-bottom: 1px solid black;">NFC Карты</td>
                                <td colspan="1">Дебет:</td>
                                <td><asp:TextBox ID="nfcInDebet" Width="135pt" runat="server"/></td>
                                <td colspan="1">Дебет:</td>
                                <td><asp:TextBox ID="nfcOutDebet" Width="135pt" runat="server"/></td>
                                <td colspan="1">Дебет:</td>
                                <td colspan="2"><asp:TextBox ID="nfcRetDebet" Width="135pt" runat="server"/></td>
                            </tr>
                            <tr>
                                <td colspan="1" style="border-bottom:1px solid black;">Кредит:</td>
                                <td style="border-bottom:1px solid black;"><asp:TextBox ID="nfcInCredit" Width="135pt" runat="server"/></td>
                                <td colspan="1" style="border-bottom:1px solid black;">Кредит:</td>
                                <td style="border-bottom:1px solid black;"><asp:TextBox ID="nfcOutCredit" Width="135pt" runat="server"/></td>
                                <td colspan="1" style="border-bottom:1px solid black;">Кредит:</td>
                                <td style="border-bottom:1px solid black;"><asp:TextBox ID="nfcRetCredit" Width="135pt" runat="server"/></td>
                            </tr>
                                                        
                            <tr>
                                <td rowspan="2" style="font-weight:bold;border-bottom: 1px solid black;">Сервисные карты</td>
                                <td colspan="1">Дебет:</td>
                                <td><asp:TextBox ID="srvInDebet" Width="135pt" runat="server"/></td>
                                <td colspan="1">Дебет:</td>
                                <td><asp:TextBox ID="srvOutDebet" Width="135pt" runat="server"/></td>
                                <td colspan="1">Дебет:</td>
                                <td><asp:TextBox ID="srvRetDebet" Width="135pt" runat="server"/></td>
                            </tr>
                            <tr>
                                <td colspan="1" style="border-bottom:1px solid black;">Кредит:</td>
                                <td style="border-bottom:1px solid black;"><asp:TextBox ID="srvInCredit" Width="135pt" runat="server"/></td>
                                <td colspan="1" style="border-bottom:1px solid black;">Кредит:</td>
                                <td style="border-bottom:1px solid black;"><asp:TextBox ID="srvOutCredit" Width="135pt" runat="server"/></td>
                                <td colspan="1" style="border-bottom:1px solid black;">Кредит:</td>
                                <td style="border-bottom:1px solid black;"><asp:TextBox ID="srvRetCredit" Width="135pt" runat="server"/></td>
                            </tr>
                            
                            <tr>
                                <td rowspan="2" style="font-weight:bold;border-bottom: 1px solid black;">Карты МИР</td>
                                <td colspan="1">Дебет:</td>
                                <td><asp:TextBox ID="mirInDebet" Width="135pt" runat="server"/></td>
                                <td colspan="1">Дебет:</td>
                                <td><asp:TextBox ID="mirOutDebet" Width="135pt" runat="server"/></td>
                                <td colspan="1">Дебет:</td>
                                <td><asp:TextBox ID="mirRetDebet" Width="135pt" runat="server"/></td>
                            </tr>
                            <tr>
                                <td colspan="1" style="border-bottom:1px solid black;">Кредит:</td>
                                <td style="border-bottom:1px solid black;"><asp:TextBox ID="mirInCredit" Width="135pt" runat="server"/></td>
                                <td colspan="1" style="border-bottom:1px solid black;">Кредит:</td>
                                <td style="border-bottom:1px solid black;"><asp:TextBox ID="mirOutCredit" Width="135pt" runat="server"/></td>
                                <td colspan="1" style="border-bottom:1px solid black;">Кредит:</td>
                                <td style="border-bottom:1px solid black;"><asp:TextBox ID="mirRetCredit" Width="135pt" runat="server"/></td>
                            </tr>
                            
                            <tr>
                                <td rowspan="2" style="font-weight:bold">Пин конверты</td>
                                <td colspan="1">Дебет:</td>
                                <td><asp:TextBox ID="pinInDebet" Width="135pt" runat="server"/></td>
                                <td colspan="1">Дебет:</td>
                                <td><asp:TextBox ID="pinOutDebet" Width="135pt" runat="server"/></td>
                                <td colspan="1">Дебет:</td>
                                <td><asp:TextBox ID="pinRetDebet" Width="135pt" runat="server"/></td>
                            </tr>
                            <tr>
                                <td colspan="1">Кредит:</td>
                                <td><asp:TextBox ID="pinInCredit" Width="135pt" runat="server"/></td>
                                <td colspan="1">Кредит:</td>
                                <td><asp:TextBox ID="pinOutCredit" Width="135pt" runat="server"/></td>
                                <td colspan="1">Кредит:</td>
                                <td><asp:TextBox ID="pinRetCredit" Width="135pt" runat="server"/></td>
                                
                                
                            </tr>
                            
                        </table>
                    </div>
                </td>
         </tr>
          <tr>
                <td style="background-color:#006699;color:white;font-weight:bold;">Операционный день</td>
             
           </tr>
           <tr>
                <td>
                    <div id="Div1" style="overflow: auto;border:1px solid black;">
                        <table cellspacing="0" style="color:#000080;font-weight:normal" border="0" width="100%">
                        <%if (isMainFilial() == false){%>  
                          <tr>
                               <td colspan="5" style="border-bottom: 1px solid black;">
                               <asp:CheckBox ID="bUseMainBranch" runat="server" TextAlign="Right" AutoPostBack="true" 
                                    Text="не использовать время головного офиса" OnCheckedChanged="bUseMainBranch_CheckedChanged"/>  
                               </td>
                          </tr>
                        <%}%>  
                           <tr>
                               <td>Время начала с:</td>
                               <td><asp:TextBox  ID="operDayStart" runat="server"/></td>
                               <td>Время окончания до:</td>
                               <td><asp:TextBox ID="operDayEnd" runat="server"/></td>
                               <td>
                                  <div id="today_tommorrow">
                                    <asp:RadioButton Text="Окончание сегодня" ID="today" runat="server" GroupName="today_tommorrow"/> <!--<label for="today">Сегодня</label>-->
                                    <asp:RadioButton Text="Окончание завтра" ID="tomorrow" runat="server" GroupName="today_tommorrow"/> <!--<label for="today">Завтра</label>-->
                                  </div>
                               </td>
                               
                           </tr>
                           <tr>
                               <td colspan="5" align="center" style="border-top: 1px solid black;border-bottom: 1px solid black;">
                                    <asp:CheckBox ID="bUseShift" runat="server" TextAlign="Right" 
                                        AutoPostBack="true" Text="использовать смены" 
                                        oncheckedchanged="bUseShift_CheckedChanged" />
                               </td>
                           </tr>    
                           <tr>
                                <td colspan="5" align="center">Первая смена с:&nbsp;
                                <asp:TextBox ID="shift1s" runat="server"/>&nbsp;до:
                                <asp:TextBox ID="shift1e" runat="server"/>&nbsp;Вторая смена с:&nbsp;<asp:TextBox ID="shift2s" runat="server"/>
                                    &nbsp;до:&nbsp;<asp:TextBox ID="shift2e" runat="server"/>
                                </td>
                           </tr>
                           <tr>
                            <td colspan="5" align="center"><label><%=raschet%></label></td>
                           </tr>
                        </table>
                    </div>
                </td>
           </tr>             
                            

        </table>
        </div>
       
    <%if(isShowAccount()==true) {%>    
    <script type="text/javascript">
    //window.onload = load_branch;
    window.onload=function() 
    { 
        document.getElementById("acountTable").style.display="block";
        document.getElementById("pBranch").style.height=screen.height-840+'px'; 
        
    }
    </script>
    <%}%> 
    <%if(isShowAccount()==false) {%>    
    <script type="text/javascript">
      window.onload=function() 
      { 
        document.getElementById("acountTable").style.display="none";
        load_branch();
      }
    </script>
    <%}%> 
    
    
</asp:Content>    