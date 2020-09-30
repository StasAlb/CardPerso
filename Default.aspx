<%@ Page Title="CardPerso" Language="C#" AutoEventWireup="true" MasterPageFile="~/CardPerso.master" CodeBehind="Default.aspx.cs" Inherits="OstCard.CardPerso._Default" %>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
<!--  
  <script>
  
  
  
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
    $( "#spinner" ).timespinner();
  });

  </script>


  <input id="spinner" value="08:30"><br/>
  <label id="lb">12345</label>
  -->
  

</asp:Content>