<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/CardPerso.Master" CodeBehind="GenDoc.aspx.cs" Inherits="CardPerso.GenDoc" %>
<%@ Register Assembly="DatePicker" Namespace="OstCard.WebControls" TagPrefix="ost" %>

<asp:Content ID="GenDocHeader" runat="server" ContentPlaceHolderID="HeadContent">
    <title>Отчеты</title>

    <link href="<%=Page.ResolveUrl("~/Styles/DataTables/demo_page.css")%>" rel="stylesheet" type="text/css" />
    <link href="<%=Page.ResolveUrl("~/Styles/DataTables/demo_table.css")%>" rel="stylesheet" type="text/css" />
    <link href="<%=Page.ResolveUrl("~/Styles/DataTables/demo_table_jui.css")%>" rel="stylesheet" type="text/css" />
    
    <script src="<%=Page.ResolveUrl("~/javascript/jquery-2.1.1.js")%>" type="text/javascript"></script>
    <script src="<%=Page.ResolveUrl("~/javascript/jquery.dataTables.min.js")%>" type="text/javascript"></script>
    <script src="<%=Page.ResolveUrl("~/javascript/jquery-ui.js")%>" type="text/javascript"></script>
<!--
    <script src="/javascript/jquery-2.1.1.js" type="text/javascript"></script>
    <script src="/javascript/jquery.dataTables.min.js" type="text/javascript"></script>
    <script src="/javascript/jquery-ui.js" type="text/javascript"></script>           -->

    
    <script language="javascript" type="text/javascript">
	
        var oReportTable = null;

        function refresh() {
            if (oReportTable != null)
                oReportTable.fnDraw();
        }

        //setInterval("refresh()", 5000);
        
        $(document).ready(function () {

            oReportTable = $('#reports').dataTable({                
                "bLengthChange": false,
                "bInfo": false,
                "iDisplayLength": 10,
                "bFilter": false,
                "bServerSide": true,
                "sAjaxSource": "Home/ReportQueueRefresh",
                "aoColumns": [
                    { "aTargets": 0 },
                    { "aTargets": 1, "sClass": "colCenter" },
                    { "aTargets": 2, "sClass": "colCenter" },
                    {
                        "aTargets": -1,
                        "sClass": "colCenter",
                        "fnRender": function (data, type, row) {
                            res = '<div>';
                            if (data.aData[3] == "3")
                                res = res + '<input type="image" src="Images/download-file.png" onclick=DownloadReport(' + data.aData[4] + ',' + data.aData[5]+ ')>  </input>';
                            res = res + '<input type="image" src="Images/delete-file.png" onclick=DeleteReport(' + data.aData[4] + ')></input>';
                            res = res + '</div>';
                            return res;                            
                        }
                    },
                    {
                        "aTarget": 4,
                        "bVisible": false
                    },
                    {
                        "aTarget": 5,
                        "bVisible": false
                    }
                ]
            });
        });

        function DeleteReport(ReportId) {
            if (!confirm("вы уверены, что хотите удалить отчет?"))
                return;
            var url = "Home/DeleteReport";
            $.post(url, { reportId: ReportId }, function (data) {               
            });
        }
        function DownloadReport(ReportId, ReportType) {
            var url = "Home/DownloadReport";
		var sss = "";
            $.ajax({
                type: "post",
                url: url,
                data: {reportId:ReportId},
		success: function(res) {sss="Home/Download?filename="+res+"&type="+ReportType;window.location.href=sss;},
		error: function(e){confirm("error:" + e);}
		 });                     		
		$(document).submit(function(event){event.preventDefault();});

        }
        
    </script>
    <style>
        .colCenter
        {
            text-align:center;
            vertical-align:middle;
        }
    </style>

</asp:Content>
<asp:Content ID="GenDocContent" runat="server" ContentPlaceHolderID="MainContent">
    <table width="100%">
        <tr style="display:none">
        <td colspan="4">
            <asp:TextBox ID="dirDBF" runat="server"></asp:TextBox>
            <asp:Button ID="bOkDirDBF" runat="server" onclick="bOkDirDBF_click"></asp:Button>
        </td>
        </tr>
        <tr>
            <td width="15%"></td>
            <td align="center" width="65%">
            <asp:DropDownList ID="dListDoc" runat="server" Width="100%"  
                    onselectedindexchanged="dListDoc_SelectedIndexChanged" AutoPostBack="True">
                </asp:DropDownList></td>
            <td width="5%" align="left" valign="middle">
               <asp:ImageButton ID="bExcel" runat="server" ImageUrl="~/Images/excel.bmp" ToolTip="Документ по старому" onclick="bExcel_Click" />
               <asp:ImageButton ID="bReportQuery" runat="server" ImageUrl="~/Images/excel.bmp" ToolTip="Документ через сервис" onclick="bReportQuery_Click"/>
               <asp:ImageButton ID="bEMail" runat="server" ImageUrl="~/Images/mail.png" ToolTip="Отправить по рассылке" onclick="bMail_Click" />
               <asp:ImageButton ID="bExportDBF" runat="server" ImageUrl="~/Images/excel2.bmp" ToolTip="Экспорт в DBF" onclick="bExportDBF_Click" OnClientClick="ShowWait('Пожалуйста, подождите завершения экспорта...');return true;" />          

               <asp:Label ID="lbInform" runat="server"></asp:Label>
               
            </td>
            <td width="15%"><asp:Label ID="lInfo" runat="server" style="width:100%"></asp:Label></td>
        </tr>    
        <tr><td></td></tr>
        <asp:Panel runat="server" ID="PanelTwoPeriod" Visible="false">
        <tr>
            <td></td>
            <td align="center" colspan="2"><table><tr><td>За период с</td><td><ost:DatePicker  ID="DatePickerStart" runat="server" AutoPostBack="true" PaneWidth="150px">
        <PaneTableStyle BorderColor="#707070" BorderWidth="1px" BorderStyle="Solid" />
        <PaneHeaderStyle BackColor="#0099FF" />
        <TitleStyle ForeColor="White" Font-Bold="true" />
        <NextPrevMonthStyle ForeColor="White" Font-Bold="true" />
        <NextPrevYearStyle ForeColor="#E0E0E0" Font-Bold="true" />
        <DayHeaderStyle BackColor="#E8E8E8" />
        <TodayStyle BackColor="#FFFFCC" ForeColor="#000000" Font-Underline="false" BorderColor="#FFCC99"/>
        <AlternateMonthStyle BackColor="#F0F0F0" ForeColor="#707070" Font-Underline="false"/>
        <MonthStyle BackColor="" ForeColor="#000000" Font-Underline="false"/>            
            </ost:DatePicker>
            </td>
            <td>по</td>
            <td><ost:DatePicker ID="DatePickerEnd" runat="server" AutoPostBack="true" PaneWidth="150px">
        <PaneTableStyle BorderColor="#707070" BorderWidth="1px" BorderStyle="Solid" />
        <PaneHeaderStyle BackColor="#0099FF" />
        <TitleStyle ForeColor="White" Font-Bold="true" />
        <NextPrevMonthStyle ForeColor="White" Font-Bold="true" />
        <NextPrevYearStyle ForeColor="#E0E0E0" Font-Bold="true" />
        <DayHeaderStyle BackColor="#E8E8E8" />
        <TodayStyle BackColor="#FFFFCC" ForeColor="#000000" Font-Underline="false" BorderColor="#FFCC99"/>
        <AlternateMonthStyle BackColor="#F0F0F0" ForeColor="#707070" Font-Underline="false"/>
        <MonthStyle BackColor="" ForeColor="#000000" Font-Underline="false"/>            
            </ost:DatePicker>            
            </td>
            <td align="left"><asp:Label Font-Size="X-Large" ID="lDatePickerTwoError" runat="server" Visible="false" Text="*" ForeColor="Red"></asp:Label></td>
            </tr>
            </table>
            </td>
            <td></td>
        </tr>

        <tr>
            <td colspan="5" align="center"><div><label style="color:Red"><%=excludeMemoricProduct%></label></div></td>
        </tr>   
        
        </asp:Panel>
        <asp:Panel runat="server" ID="PanelOnePeriod" Visible="false">
        <tr>
        <td></td>
        <td colspan="2" align="center">
        Выберите день <ost:DatePicker ID="DatePickerOne" runat="server" AutoPostBack="true" PaneWidth="150px" OnSelectedDateChanged="DatePickerOne_OnSelectedDateChanged">
        <PaneTableStyle BorderColor="#707070" BorderWidth="1px" BorderStyle="Solid" />
        <PaneHeaderStyle BackColor="#0099FF" />
        <TitleStyle ForeColor="White" Font-Bold="true" />
        <NextPrevMonthStyle ForeColor="White" Font-Bold="true" />
        <NextPrevYearStyle ForeColor="#E0E0E0" Font-Bold="true" />
        <DayHeaderStyle BackColor="#E8E8E8" />
        <TodayStyle BackColor="#FFFFCC" ForeColor="#000000" Font-Underline="false" BorderColor="#FFCC99"/>
        <AlternateMonthStyle BackColor="#F0F0F0" ForeColor="#707070" Font-Underline="false"/>
        <MonthStyle BackColor="" ForeColor="#000000" Font-Underline="false"/>            
            </ost:DatePicker>             
        </td>
        <td align="left"><asp:Label Font-Size="X-Large" ID="lDatePickerOneError" runat="server" Visible="false" Text="*" ForeColor="Red"></asp:Label></td>
        </tr>
        <tr>
        <td colspan="4" align="center"><div><label><%=OneDataMessage()%></label></div></td>
        </tr>
        <%if (isMemoricType()==false) {%>
        <tr style="display:none">
        <%} else { %>
        <tr valign="middle" style="display:none">
        <%}%>
           <td colspan="4" align="center">
               <asp:Label ID="userbranch"  Text="Выводить по: " runat="server"/> 
               <asp:RadioButton Text="текущий пользователь" ID="userradio" runat="server" GroupName="userorbranch" AutoPostBack="true"/>
               <asp:RadioButton Text="текущее подразделение" ID="branchradio" runat="server" GroupName="userorbranch" AutoPostBack="true"/>
           </td>
        </tr>
        <%if (isTwoDay == false || isMemoricType()==false)
          {%>
        <tr style="display:none">
        <%} else { %>
        <tr valign="middle">
        <%}%>
           <td colspan="4" align="center">
                <asp:Label ID="rangeout" Text="Диапазон вывода: " runat="server" />
                <asp:RadioButton ID="allrangeradio" Text="все" runat="server" GroupName="rangeradio" Font-Underline="True" Checked="True" />     
                <asp:RadioButton ID="firstrangeradio" Text="вчера" runat="server" GroupName="rangeradio" />     
                <asp:RadioButton ID="nextrangeradio" Text="сегодня" runat="server" GroupName="rangeradio" />     
           </td>
        </tr>
                
        <tr>
            <td colspan="4" align="center"><div><label style="color:Red"><%=excludeMemoricProduct%></label></div></td>
        </tr>   
        </asp:Panel>        
        <asp:Panel runat="server" ID="PanelDinamyc" Visible="false">
        <tr>
        <td></td>
        <td colspan="2" align="center">
            <asp:RadioButtonList ID="rbPeriods" runat="server">
                <asp:ListItem Text="по месяцам" Value="1"></asp:ListItem>
                <asp:ListItem Text="по дням" Value="2" Selected="True"></asp:ListItem>
            </asp:RadioButtonList>
            Число дней в периоде <asp:TextBox ID="tbDays" runat="server" Text="1"></asp:TextBox><asp:Label Text="*" runat="server" ID="lDaysWarning" ForeColor="Red" Visible="false"></asp:Label>
        </td>
        <td></td>
        </tr>
        <tr>
            <td></td>
            <td colspan="2" align="center">по продуктам <asp:DropDownList ID="ddlProducts" runat="server"></asp:DropDownList></td>
            <td></td>
        </tr>
        <tr>
            <td colspan="4" align="center"></td>
        </tr>
        </asp:Panel>
        <asp:Panel runat="server" ID="PanelTreasures" Visible="false">
        <tr>
        <td></td>
        <td colspan="2" align="center">
            <asp:DropDownList runat="server" ID="ddlBranch"></asp:DropDownList>
        </td>
        <td></td>
        </tr>
        <tr>
            <td colspan="4">
                
            </td>
        </tr>
        </asp:Panel>
        <asp:Panel runat="server" ID="PanelPersons" Visible="false">
            <tr>
                <td></td>
                <td colspan="2" align="center">
                    <asp:DropDownList runat="server" ID="ddlPersons"></asp:DropDownList>
                </td>
                <td></td>
            </tr>
            <tr>
                <td colspan="4">
                
                </td>
            </tr>
        </asp:Panel>



<%--        <asp:Timer runat="server" ID="Timer" OnTick="Timer_OnTick" Interval="1000"/>--%>
        <tr >
            <td colspan="4">
            <table id="reports" class="display" style="display:none">
            <thead>
                <tr>
                    <th>Название отчета</th>
                    <th>Дата заявки</th>
                    <th>Статус отчета</th>
                    <th>Действия</th>
                    <th></th>
                    <th></th>
                </tr>
            </thead>
            <tbody></tbody>
        </table>
            </td>
        </tr>

        
            <asp:Label ID="lbCount" runat="server"/>
</table>
        
    <div id="result_f"></div>
    
</asp:Content>