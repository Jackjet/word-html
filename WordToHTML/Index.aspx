<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Index.aspx.cs" Inherits="WordToHTML.Index" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:FileUpload ID="FileUpload1" runat="server" Width="400px" />&nbsp;
        <asp:Button ID="btnChange" runat="server" Text="转换成HTML" onclick="btnChange_Click" />
        <br />
        <iframe id="ifrm" runat="server" width="100%" height="300px"></iframe> 
    </div>
    </form>
</body>
</html>
