<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/CardPerso.Master" CodeBehind="StorDoc.aspx.cs" Inherits="CardPerso.StorDoc" %>
<asp:Content ID="StorDocHeader" runat="server" ContentPlaceHolderID="HeadContent">
    <meta http-equiv="Pragma" content="no-cache" />
    <title>Движение</title>
    <!--script type="text/javascript" src="jquery.js"></script>
    <script type="text/javascript" src="splitter.js"></script>
    <script type="text/javascript" src="jquery.cookie.js"></script>
    <script type="text/javascript">
    $().ready(function() {
	    $("#MySplitter").splitter({
	        type: "h", 
	        minTop: 100,
	        splitHorizontal: true,
	        outline: true,
	        anchorToWindow: true,
	        cookie: "splittercookie"
	        });	
    });
    function PaneResize()
    {
        //alert(document.getElementById("BottomPane").clientHeight-20);
        document.getElementById("pDocs").style.height = document.getElementById("TopPane").clientHeight-25+'px';
        document.getElementById("pBottom").style.height = document.getElementById("BottomPane").clientHeight-40+'px';
    }
</script-->
<style type="text/css" media="all">
#MySplitter {
	border: none;
	min-width: 500px;	/* Splitter can't be too thin ... */
	min-height: 300px;	/* ... or too flat */
	height: 100%;
}
#MySplitter .Pane {
	overflow: auto;
	background: #def;
}
#MySplitter .TopPane 
{
	min-height: 100px;
}
#MySplitter .BottomPane
{
	min-height: 100px;
}

/* Splitbar styles; these are the default class names */
.hsplitbar {
	height: 6px;
	background: #669 url(img/hgrabber.gif) no-repeat center;
}
.hsplitbar.active, .hsplitbar:hover {
	background: #c66 url(img/hgrabber.gif) no-repeat center;	
}
    
</style>


</asp:Content>
<asp:Content ID="StorDocContent" runat="server" ContentPlaceHolderID="MainContent">

    <div id="MySplitter" style="width:98%">
    <div id="TopPane">
    <table width="100%" cellpadding="0" cellspacing="0">
         <tr>
            <td>
                <asp:ImageButton ID="bNewD" runat="server" ImageUrl="~/Images/new.gif" 
                ToolTip="Новый документ" OnClientClick="return show_stordoc('mode=1');" onclick="bNewD_Click" />
                <asp:ImageButton ID="bEditD" runat="server" ImageUrl="~/Images/edit.bmp" 
                 ToolTip="Редактировать документ" onclick="bEditD_Click" />
                <asp:ImageButton ID="bDeleteD" runat="server" ImageUrl="~/Images/del.bmp" 
                ToolTip="Удалить документ" onclick="bDeleteD_Click"  />
                <asp:ImageButton ID="bSetFilter" runat="server" ImageUrl="~/Images/sfil.bmp" 
                ToolTip="Фильтр"  OnClientClick="return show_flt_stordoc();" onclick="bSetFilter_Click" />
                <asp:ImageButton ID="bResetFilter" runat="server" ImageUrl="~/Images/rfil.bmp" 
                 ToolTip="Снять фильтр" onclick="bResetFilter_Click" />
                 <asp:ImageButton ID="bConfFieldD" runat="server" ImageUrl="~/Images/field.bmp" 
                 ToolTip="Настройка полей" OnClientClick="return show_field('type=stordoc','260px');" 
                    onclick="bConfFieldD_Click" />
                 <asp:ImageButton ID="bSostD" runat="server" ImageUrl="~/Images/select.bmp" 
                     ToolTip="Изменить состояние" onclick="bSostD_Click" /> 
                 <asp:ImageButton ID="bAutoD" runat="server" ImageUrl="~/Images/auto.bmp" 
                     ToolTip="Сформировать автоматически" OnClientClick="return confirm('Сформировать документ автоматически?');" 
                     onclick="bAutoD_Click" />
                <asp:ImageButton ID="bDeleteD2" runat="server" ImageUrl="~/Images/del2.bmp" 
                 ToolTip="Удалить продукцию" 
                    OnClientClick="return confirm('Удалить продукцию из документа?');" 
                    onclick="bDeleteD2_Click" />
                <asp:ImageButton ID="bChangeDate" runat="server" ImageUrl="~/Images/edit.gif" ToolTip="Изменить дату документа" OnClick="bChangeDate_Click" />    
                 <asp:DropDownList ID="dListDoc" runat="server" Width="150px"></asp:DropDownList>
                 <asp:ImageButton ID="bExcelD" runat="server" ImageUrl="~/Images/excel.bmp" 
                     ToolTip="Сформировать документ" onclick="bExcelD_Click" />
            </td>
            <td align="right">
                <asp:Label ID="lbInform" runat="server" ForeColor="Red"></asp:Label>
                <asp:Label ID="lbCountD" runat="server" Font-Bold="True" ForeColor="#400080"></asp:Label>
            </td>
        </tr>
        <tr>
            <td colspan="2">
            <div id="pDocs">
               <asp:GridView ID="gvDocs" runat="server" AutoGenerateColumns="False" 
                BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" 
                CellPadding="3" DataKeyNames="id,type,number_doc,priz_gen,date_doc,type_name,comment,id_branch,branch,id_courier,id_deliv,id_act,courier,invoice_courier,courier_name, fulldate, LoweredUserName" 
                Width="98%" onselectedindexchanged="gvDocs_SelectedIndexChanged" 
                style="margin-right: 0px" AllowSorting="True" onsorting="gvDocs_Sorting" AllowPaging="true" OnPageIndexChanging="gvDocs_PageIndexChanging">
                <PagerSettings Mode="NumericFirstLast" Position="TopAndBottom" />
                <FooterStyle BackColor="White" ForeColor="#000066" />
                <RowStyle ForeColor="#000066" />
                <Columns>
                    <asp:BoundField DataField="number_doc" HeaderText="Номер" 
                        SortExpression="number_doc" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="fulldate"  
                        HeaderText="Дата" SortExpression="date_doc, time_doc" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="type_name"
                        HeaderText="Тип документа" SortExpression="type" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="branch"
                        HeaderText="Филиал" SortExpression="id_branch" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="gen" HeaderText="Состояние" 
                        SortExpression="priz_gen" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="LoweredUserName" HeaderText="Создал"
                        SortExpression="LoweredUserName">
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="comment" HeaderText="Комментарий" 
                        SortExpression="comment" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:CommandField ButtonType="Image" SelectImageUrl="~/Images/select.gif" 
                        SelectText="Выбрать" ShowSelectButton="True" ItemStyle-HorizontalAlign="Center" >
                        <ItemStyle HorizontalAlign="Center"/>
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
        </div>
        <div id="BottomPane">
        <table width="100%">
        <tr>
            <td>
            <table style="width: 100%">
                <tr>
                <td style="height: 30px" valign="bottom">
                    <asp:Panel  ID="pActionProd" runat="server">
                    <asp:Label ID="lbCountProd" runat="server" Font-Bold="True" ForeColor="#400080"></asp:Label>
                    <asp:ImageButton ID="bNewProd" runat="server" ImageUrl="~/Images/new.gif" 
                        ToolTip="Новый продукт" onclick="bNewProd_Click"  />
                    <asp:ImageButton ID="bEditProd" runat="server" ImageUrl="~/Images/edit.bmp" 
                        ToolTip="Редактировать продукт" onclick="bEditProd_Click"  />
                    <asp:ImageButton ID="bDelProd" runat="server" ImageUrl="~/Images/del.bmp" 
                        ToolTip="Удалить продукт" onclick="bDelProd_Click" />
                    <asp:ImageButton ID="bAutoProd" runat="server" ImageUrl="~/Images/auto.bmp" 
                     ToolTip="Сформировать продукцию автоматически" 
                            OnClientClick="return confirm('Сформировать продукцию автоматически?');" 
                            onclick="bAutoProd_Click"/>
                    </asp:Panel>
                    
                    <asp:Panel  ID="pActionCard" runat="server"  DefaultButton="bPanSearch">
                    <asp:Label ID="lbCountCard" runat="server" Font-Bold="True" ForeColor="#400080"></asp:Label>
                    <asp:ImageButton ID="bNewCard" runat="server" ImageUrl="~/Images/new.gif" 
                        ToolTip="Новая карта" onclick="bNewCard_Click" />
                    <asp:ImageButton ID="bEditCard" runat="server" ImageUrl="~/Images/edit.bmp" 
                        ToolTip="Изменить комментарий" onclick="bEditCard_Click" />
                    <asp:ImageButton ID="bCardProperty" runat="server" ImageUrl="~/Images/select.bmp" 
                        ToolTip="Изменить свойство карты" onclick="bCardProperty_Click" />

                    <asp:ImageButton ID="bDelCard" runat="server" ImageUrl="~/Images/del.bmp" 
                        ToolTip="Удалить карту" onclick="bDelCard_Click" /> 
                    
                    <asp:ImageButton ID="bExpertiza" runat="server" ImageUrl="~/Images/new.bmp"
                        ToolTip = "Результат экспертизы" onclick="bExpertiza_Click" />
                    
                    <asp:Panel ID="pnlPanSearch" runat="server">
                    <asp:Label runat="server" ID="lMisc">Быстрый поиск по номеру карты:</asp:Label>
                    <asp:TextBox runat="server" ID="tbPanSearch" AutoPostBack="false"></asp:TextBox>
                    <asp:Label runat="server" ID="Label1">Держатель:</asp:Label>
                    <asp:TextBox runat="server" ID="tbFioSearch"></asp:TextBox>
                    <asp:Button runat="server" ID="bPanSearch" OnClick="bPanSearch_Click" Text="Поиск" />
                    </asp:Panel>    
                    </asp:Panel>
                 </td>
                 <td align="right">
                       <asp:LinkButton ID="lbProduct" runat="server" onclick="lbProduct_Click">Продукция</asp:LinkButton>
                       <asp:LinkButton ID="lbCard" runat="server" onclick="lbCard_Click">Карты</asp:LinkButton>
                 </td>
                 </tr>
             </table>
             <div id="pBottom">
             <asp:Panel  ID="mpProducts" runat="server"> 
                <div id="pProducts" style="overflow: auto; height: 100%;">
                <asp:GridView ID="gvProducts" runat="server" BackColor="White" 
                BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="3" 
                AutoGenerateColumns="False" Width="98%" DataKeyNames="id,id_prb,cnt_new,cnt_perso,cnt_brak,id_type,prod_name,bank_name" 
                       onselectedindexchanged="gvProducts_SelectedIndexChanged">
                <FooterStyle BackColor="White" ForeColor="#000066" />
                <RowStyle ForeColor="#000066" />
                <Columns>
                    <asp:BoundField DataField="prod_name" HeaderText="Наименование" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="bank_name" HeaderText="Банк" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="cnt_new" HeaderText="Новые" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="cnt_perso" HeaderText="Персо" >
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>
                    <asp:BoundField DataField="cnt_brak" HeaderText="Брак" >
                        <ItemStyle HorizontalAlign="Center" />
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
             </asp:Panel>
             
             <asp:Panel ID="mpCards" runat="server">
                <div id="pCards" style="overflow: auto; height: 100%;">
                <asp:GridView ID="gvCards" runat="server" BackColor="White" 
                BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="3" 
                AutoGenerateColumns="False" Width="98%" DataKeyNames="id,id_doc,id_card,pan,fio,id_stat" 
                             onselectedindexchanged="gvCards_SelectedIndexChanged" 
                        onrowdatabound="gvCards_RowDataBound" >
                <FooterStyle BackColor="White" ForeColor="#000066" />
                <RowStyle ForeColor="#000066" />
                <Columns>
                    <asp:BoundField DataField="bank_name" HeaderText="Банк" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="BranchCard" HeaderText="Филиал" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="pan" HeaderText="Номер" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
<%--                    <asp:BoundField DataField="fio" HeaderText="ФИО" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>--%>
                    <asp:TemplateField HeaderText="ФИО">
                        <ItemTemplate>
                            <table style="align-content:stretch; width:100%;margin:0;">
                                <tr>
                                    <td style="align-content:flex-start"><asp:Label runat="server" Text='<%#(string)Eval("fio")%>'></asp:Label></td>
                                    <td style="vertical-align:top;align-content:flex-end; width:20px;"><asp:Image runat="server" ImageUrl="~/Images/icons8-n-16.jpg" Visible='<%#((bool)Eval("isPin"))%>' /></td>
                                </tr>
                            </table>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:BoundField DataField="passport" HeaderText="Паспорт" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="status" HeaderText="Статус" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    <asp:BoundField DataField="prop" HeaderText="Состояние" >
                        <ItemStyle HorizontalAlign="Left" />
                    </asp:BoundField>
                    
                    <asp:BoundField DataField="comment" HeaderText="Комментарий">
                        <ItemStyle HorizontalAlign="Left" />
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
             </asp:Panel>
             </div>
            </td>
        </tr>
    </table>
    </div>
  </div>  
    <div id="popup_cardnumber" title=""  class="ui-widget" style="display:none">
		<div style="display:none">
            <asp:Button ID="bSearchCardNumber" runat="server" Text="1" 
                onclick="bSearchCardNumber_Click"/>
		    <asp:Button ID="bSaveCardNumber" runat="server" Text="2" 
                onclick="bSaveCardNumber_Click" style="width: 21px"/>  
		    <asp:Button ID="bCancelCardNumber" runat="server" Text="3" 
                onclick="bCancelCardNumber_Click"/>       
		</div>
		<table cellpadding="2px" cellspacing="5px" width="100%" border="0">
		    <tr>	
		        <td>Номер карты:</td>
		        <td >
		            <asp:TextBox ID="txtCardNumber" runat = "server" 
		                         CssClass="ui-widget-content ui-corner-all" style="width:90%" />
		        </td>
		    </tr>
		    <tr>
		        <td>Владелец карты:</td>
		        <td >
		            <asp:Label ID="txtCardNumberInfo" runat = "server" style="width:90%"></asp:Label>
		        </td>
		        
		    </tr>
		    <tr id = "client_compilant">
		        <td>Жалоба клиента:</td>
		        <td>
		            <asp:TextBox ID="txtClientCompilant" runat = "server"  Rows = "5" TextMode="MultiLine" Wrap="true"
		                         CssClass="ui-widget-content ui-corner-all" style="width:90%" />
		        </td>
		        
		    </tr>
		    
		</table>
	</div>
    
    <div id="popup_clientcompilant" title=""  class="ui-widget" style="display:none">
		<div style="display:none">
            <asp:Button ID="bSaveClientCompilant" runat="server" Text="1" 
                onclick="bSaveClientCompilant_Click"/>
            <asp:Button ID="bCancelClientCompilant" runat="server" Text="2" 
                onclick="bCancelClientCompilant_Click"/>       
		</div>
		<table cellpadding="2px" cellspacing="5px" width="100%" border="0">
		    <tr>
		        <td>Жалоба клиента:</td>
		        <td>
		            <asp:TextBox ID="txtClientCompilantEdit" runat = "server"  Rows = "5" TextMode="MultiLine" Wrap="true"
		                         CssClass="ui-widget-content ui-corner-all" style="width:90%" />
		        </td>
		        
		    </tr>
		</table>
	</div>
	
	<div id="popup_cardproperty" title=""  class="ui-widget" style="display:none">
		<div style="display:none">
            <asp:Button ID="bSaveCardProperty" runat="server" Text="1" 
                onclick="bSaveCardProperty_Click"/>
            <asp:Button ID="bCancelCardProperty" runat="server" Text="2" 
                onclick="bCancelCardProperty_Click"/>       
		</div>
		<table cellpadding="2px" cellspacing="5px" width="100%" border="0">
		    <tr>
		        <td>Жалоба клиента:</td>
		        <td>
		            <asp:TextBox ID="txtClientCompilantView" runat = "server"  Rows = "5" TextMode="MultiLine" Wrap="true"
		                         CssClass="ui-widget-content ui-corner-all" style="width:90%" ReadOnly="True" />
		        </td>
		        
		    </tr>
		    <tr>
		        <td>Свойство карты:</td>
		        <td>
		            <asp:DropDownList runat="server" id="lCardsProperty" CssClass="ui-widget-content ui-corner-all" style="width:90%"></asp:DropDownList>
		        </td>
		    </tr>
		    <tr>
		        <td>Результат экспертизы:</td>
		        <td>
		            <asp:TextBox ID="txtClientCompilantResult" runat = "server"  Rows = "5" TextMode="MultiLine" Wrap="true"
		                         CssClass="ui-widget-content ui-corner-all" style="width:90%" ReadOnly="false"/>
		        </td>
		        
		    </tr>
		</table>
	</div>
    
    
    <asp:Label ID="lbSearch" runat="server" Visible="False"></asp:Label>
    <asp:Label ID="lbSort" runat="server" Visible="False"></asp:Label>
    <asp:Label ID="lbSortIndex" runat="server" Visible="False"></asp:Label>
    <asp:Label ID="lbViewP" runat="server" Visible="False"></asp:Label>
    
    <script type="text/javascript">
    
    
   
   var bSearchCardNumber="#<%=bSearchCardNumber.ClientID%>";
   var bSaveCardNumber="#<%=bSaveCardNumber.ClientID%>";
   var bCancelCardNumber="#<%=bCancelCardNumber.ClientID%>";
   var txtCardNumber = "#<%=txtCardNumber.ClientID%>";
   var txtCardNumberInfo = "#<%=txtCardNumberInfo.ClientID%>";
   var txtClientCompilant = "#<%=txtClientCompilant.ClientID%>";
   
   var bSaveClientCompilant="#<%=bSaveClientCompilant.ClientID%>";
   var bCancelClientCompilant="#<%=bCancelClientCompilant.ClientID%>";
   var txtClientCompilantEdit = "#<%=txtClientCompilantEdit.ClientID%>";
   
   var bSaveCardProperty="#<%=bSaveCardProperty.ClientID%>";
   var bCancelCardProperty="#<%=bCancelCardProperty.ClientID%>";
   var txtClientCompilantView="#<%=txtClientCompilantView.ClientID%>";
   var lCardsProperty="#<%=lCardsProperty.ClientID%>";
   var txtClientCompilantResult="#<%=txtClientCompilantResult.ClientID%>";
   
   var timerCardNumber = null;
   var countTimerCardNumber = 0;
   var cardNumberOldLen = 0;
   
   String.prototype.toCardFormat = function () 
   {
        return this.replace(/[^0-9,?=]/g, "").substr(0, 128).split("").reduce(cardFormat, "");
        function cardFormat(str, l, i) 
        {
            return str + ((!i || (i % 4)) ? "" : "-") + l;
        }
   };
      
   
   function checkTxtCardNumber()
   {
     var txt = $(txtCardNumber).val();
     if(txt.length>16 && (txt.indexOf("?")>=0 || txt.indexOf(",")>=0) && txt.indexOf("=")>=0 && txt.length==cardNumberOldLen)
     {
        countTimerCardNumber = 0;
        var txtold = txt;
        var ind1=txtold.indexOf("?");
        if(ind1<0) ind1=txtold.indexOf(",");
        ind1++;
        var ind2=txtold.indexOf("=");
        txt = txtold.substr(ind1, ind2-ind1); 
        $(txtCardNumber).val(txt.toCardFormat());
        $(txtCardNumberInfo).text("");
        $("#client_compilant").css("display","none");
     }
     cardNumberOldLen = txt.length;
     countTimerCardNumber++;
     timerCardNumber = setTimeout(checkTxtCardNumber,100);
   }
   
   function clickShowCardNumber(cardNumber, cardNumberInfo)
   {
        if(timerCardNumber!=null) 
        {
            clearTimeout(timerCardNumber);
            timerCardNumber = null;
            countTimerCardNumber = 0;
        }
        
        timerCardNumber = setTimeout(checkTxtCardNumber,100);
       
        $(txtCardNumber).val(cardNumber);
        $(txtCardNumberInfo).text(cardNumberInfo); 
        
        ShowCardNumber("Введите номер карты",function f1() { ShowWait(); $(bSearchCardNumber).click();},function f2() { ShowWait(); $(bSaveCardNumber).click();},function f3() { ShowWait();  $(bCancelCardNumber).click();});
       
        return false;
   }
   
   function clickShowClientCompilant()
   {
        ShowClientCompilant("Измените жалобу клиента",function f1() { ShowWait(); $(bSaveClientCompilant).click();},function f2() { ShowWait();  $(bCancelClientCompilant).click();});
        return false;
   }
   
   function clickShowCardProperty(priz_gen)
   {
        ShowCardProperty("Свойства карты",priz_gen,function f1() { ShowWait(); $(bSaveCardProperty).click();},function f2() { ShowWait();  $(bCancelCardProperty).click();});
        return false;
   }
   
   function ShowCardProperty(title,priz_gen,onOK,onClose)
   {
        HideWait();
        var buttonsave=
	    {
    		    text:'Сохранить',click: function()
  			    {
  				    var txtres = $(txtClientCompilantResult).val();
  				    
				    if(txtres.length<1)
				    {
				        ShowError("Необходимо указать результат выполнения экспертизы", function f() {$(txtClientCompilantResult).focus();});
				        return;
				    }
				    onClose=null;
				    $(this).dialog("close");
				    onOK();
				    return;
			    }
		};
		var buttoncancel=
	    {
		    text:'Отмена',click: function()
		    {
		            $(this).dialog("close");
					return;
		    }
	
    	};
    	var buttons=((priz_gen==0) ? [buttonsave,buttoncancel]:[buttoncancel]);
	    $("#popup_cardproperty").dialog(
	    {
	    title:title,
	    resizable:false,
	    modal:true,
	    width:500,
	    close: function(event, ui) {  $(this).dialog("destroy"); if(onClose!=null) onClose(); return true; },
	    buttons:buttons
	    });

    }
     
   function ShowClientCompilant(title,onOK,onClose)
   {
        HideWait();
        var buttonsave=
	    {
    		    text:'Сохранить жалобу',click: function()
  			    {
  				    var txt = $(txtClientCompilantEdit).val();
				    if(txt.length<1)
				    {
				        ShowError("Необходимо указать на что жалуется клиент", function f() {$(txtClientCompilantEdit).focus();});
				        return;
				    }
				    onClose=null;
				    $(this).dialog("close");
				    onOK();
				    return;
			    }
		};
		var buttoncancel=
	    {
		    text:'Отмена',click: function()
		    {
		            $(this).dialog("close");
					return;
		    }
	
    	};
    	var buttons=[buttonsave,buttoncancel];
	    $("#popup_clientcompilant").dialog(
	    {
	    title:title,
	    resizable:false,
	    modal:true,
	    width:500,
	    close: function(event, ui) {  $(this).dialog("destroy"); if(onClose!=null) onClose(); return true; },
	    buttons:buttons
	    });

    }
   
   function ShowCardNumber(title,onSearch,onOK,onClose)
    {
	    HideWait();
	    //$(txtCardNumber).mask('9999-9999-9999-99?99', {placeholder:" "});
	     if($(txtCardNumberInfo).text().length<1) $("#client_compilant").css("display","none");
	     $(txtCardNumber).keyup(function () 
	     {
                $(this).val($(this).val().toCardFormat());
        });
        
        var buttonsearch=
	    {
		    text:'Найти',click: function()
		    {
   			    var txt = $(txtCardNumber).val();
			    if(txt.length<16)
			    {
			        ShowError("Неверная длина номера карты: " + txt.length, function f() {$(txtCardNumber).focus();});
			        return;
			    }
			    onClose=null;
			    $(this).dialog("close");
			    onSearch();
			    return;
		    }
	    };
	    var buttonsave=
	    {
    		    text:'Добавить карту',click: function()
  			    {
  				    var txt = $(txtCardNumberInfo).text();
				    if(txt.length<1)
				    {
				        ShowError("Карта не найдена в БД", function f() {$(txtCardNumber).focus();});
				        return;
				    }
				    txt = $(txtClientCompilant).val();
				    if(txt.length<1)
				    {
				        ShowError("Необходимо указать на что жалуется клиент", function f() {$(txtClientCompilant).focus();});
				        return;
				    }
				    
  				    onClose=null;
				    $(this).dialog("close");
				    onOK();
				    return;
			    }
		};
		var buttoncancel=
	    {
		    text:'Отмена',click: function()
		    {
		            $(this).dialog("close");
					return;
		    }
	
    	};
    	var buttons=($(txtCardNumberInfo).text().length<1)? [buttonsearch,buttoncancel]:[buttonsearch,buttonsave,buttoncancel];
        
	    $("#popup_cardnumber").dialog(
	    {
	    title:title,
	    resizable:false,
	    modal:true,
	    width:500,
	    close: function(event, ui) {  $(this).dialog("destroy"); if(onClose!=null) onClose(); return true; },
	    buttons:buttons
	    });
	    
	    //$(this).parent().find("button:eq(1)").css("display", "none");
	    
	   $(txtCardNumber).focus();
	   var txt = $(txtCardNumber).val();
	   $(txtCardNumber).val('');
	   $(txtCardNumber).val(txt);
    }
    </script>
    
    
    
    <script type="text/javascript">
    window.onload = load_stordoc;
    </script>
</asp:Content>