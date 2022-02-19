<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:Label ID="LabelMensaje" runat="server" Text="Label"></asp:Label>
    
    </div>
     <br />
      <br />
    <asp:Label ID="Label1" runat="server" Text="De:"></asp:Label>
    <asp:TextBox ID="TextBox1" runat="server" Width="303px">soportesag@intranet.gendarmeria.gov.ar</asp:TextBox>
    <br />
     <br />
      <br />
    <asp:Label ID="Label2" runat="server" Text="Para:"></asp:Label>
    <asp:TextBox ID="TextBox2" runat="server" Width="294px">soportesag@intranet.gendarmeria.gov.ar</asp:TextBox>
    <br />
     <br />
      <br />
    <asp:TextBox ID="TextBox3" runat="server" Height="225px" Width="549px">PROBANDO TEXTO PARAENVIO DE MAIL -PROYECTO CALL CENTER</asp:TextBox>
    <br />
    <br />
    <asp:Button ID="Button1" runat="server" Text="Enviar" />
    </form>
</body>
</html>
