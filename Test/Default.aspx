<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="Test._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="jumbotron">
        <h3><b>Employee salary increments</b></h3>
       <p>Use this page to upload employees excel file and then send then generate and send them increment letters automatically.</p>
    </div>
    <table>

        <tr>
            <td colspan="2">Download the excel sample file from <a href="#">Here</a></td>

        </tr>
        <tr>
            <td>
                <asp:FileUpload ID="FileUploadtoServer" Width="200px" runat="server" /></td>
            <td>
                <asp:Button ID="btnUpload" runat="server" Text="Upload File" OnClick="btnUpload_click" ValidationGroup="vg" style="width:99px;" /></td>
        </tr>
    </table>
    <br />
    <asp:Label ID="lblMsg" runat="server"></asp:Label>
    <br />
    <asp:GridView ID="GridView1" runat="server" EmptyDataText="No record found" Height="25px" CellPadding="4" ForeColor="#333333" GridLines="None">
        <AlternatingRowStyle BackColor="White" />
        <EditRowStyle BackColor="#2461BF" />
        <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
        <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
        <RowStyle BackColor="#EFF3FB" />
        <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
        <SortedAscendingCellStyle BackColor="#F5F7FB" />
        <SortedAscendingHeaderStyle BackColor="#6D95E1" />
        <SortedDescendingCellStyle BackColor="#E9EBEF" />
        <SortedDescendingHeaderStyle BackColor="#4870BE" />
    </asp:GridView>

    <asp:Button ID="btnIncrementLetter" runat="server" Visible="false" Text="Send increment letters" OnClick="btnIncrementLetter_Click" />
   
    <br />
    <asp:Label ID="lblEmailAlertMsg" runat="server"></asp:Label>
   
</asp:Content>
