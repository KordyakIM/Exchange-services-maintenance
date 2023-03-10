<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %> 
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="VisualWebPart1UserControl.ascx.cs" Inherits="ExchangeServices.VisualWebPart1.VisualWebPart1UserControl" %>
<asp:Label ID="Label1" runat="server" Text="Input Code:"></asp:Label>
<asp:Textbox ID="TextBox1" runat="server" Text="" ForeColor="Gray" TextMode="Password"></asp:Textbox>
<asp:Table id="Table2" runat="server">
    <asp:TableRow>
        <asp:TableCell>Servers</asp:TableCell>
        <asp:TableCell>Services</asp:TableCell>
    </asp:TableRow>
    <asp:TableRow>
        <asp:TableCell VerticalAlign="Top">
            <asp:Label ID="Label3" runat="server" Text="Array01:"></asp:Label>
            <asp:CheckBoxList ID="CheckBoxList_Array01" runat="server" Font-Size="Small">
                <asp:ListItem Text=Server01>Server01</asp:ListItem>
                <asp:ListItem Text=Server02>Server02</asp:ListItem>
                <asp:ListItem Text=Server03>Server03</asp:ListItem>
                <asp:ListItem Text=Server04>Server04</asp:ListItem>
                <asp:ListItem Text=Server05>Server05</asp:ListItem>
            </asp:CheckBoxList>
            <asp:Label ID="Label5" runat="server" Text="Array02:"></asp:Label>
            <asp:CheckBoxList ID="CheckBoxList_Array02" runat="server" Font-Size="Small">
                <asp:ListItem Text=Server01>Server01</asp:ListItem>
                <asp:ListItem Text=Server02>Server02</asp:ListItem>
                <asp:ListItem Text=Server03>Server03</asp:ListItem>
                <asp:ListItem Text=Server04>Server04</asp:ListItem>
                <asp:ListItem Text=Server05>Server05</asp:ListItem>
            </asp:CheckBoxList>
        </asp:TableCell>
        <asp:TableCell VerticalAlign="top">
            <asp:CheckBoxList ID="CheckBoxList_Services" runat="server" Font-Size="Small">
                <asp:ListItem Text=MSExchangeIS>Microsoft Exchange Information Store (MSExchangeIS)</asp:ListItem>
                <asp:ListItem Text=MSExchangeTransport>Microsoft Exchange Transport (MSExchangeTransport)</asp:ListItem>
                <asp:ListItem Text=MSExchangeMailboxAssistants>Microsoft Exchange Mailbox Assistants (MSExchangeMailboxAssistants)</asp:ListItem>
            </asp:CheckBoxList>
        </asp:TableCell>
    </asp:TableRow>
    
</asp:Table>

<br />
<asp:Button OnClick="Start_click" ID="button_start" runat="server" Text="Start"></asp:Button>
<asp:Button OnClick="Stop_click" ID="button_stop" runat="server" Text="Stop"></asp:Button>
<asp:Button OnClick="Restart_click" ID="button_restart" runat="server" Text="Restart"></asp:Button>
<asp:Button OnClick="CheckService_click" ID="button_checkService" runat="server" Text="Check Service"></asp:Button>
<br />
<asp:Label ID="Label2" runat="server" Text="Result:"></asp:Label>
<br />
<asp:Textbox ID="ResultBox" TextMode="MultiLine" runat="server" Height="250px" Width="800px" Text=""></asp:Textbox>