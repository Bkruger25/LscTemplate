<%@ Page Title="LSC" Language="VB" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="jumbotron">
        <h1 style="text-align:center">LSC Template Uploader</h1>   
    </div>
    <asp:Label ID="lblSuccess" Visible="false" ForeColor="Green" runat="server" Text=""></asp:Label>
    <div class="row">
        <div class="col-md-4">
           <asp:FileUpload ID="FileUpload1" runat="server" />
        </div>
        <div class="col-md-4" style="text-align:center">
            <asp:DropDownList ID="DropDownList1" runat="server">
                <asp:ListItem>No Value Selected</asp:ListItem>
                <asp:ListItem>Advanced Core Study</asp:ListItem>
                <asp:ListItem>Atterberg</asp:ListItem>
                <asp:ListItem>Geochemistry</asp:ListItem>
                <asp:ListItem>Gravel Sand Silt and Clay</asp:ListItem>
                <asp:ListItem>Particle Size Characteristics</asp:ListItem>
                <asp:ListItem>RBRC</asp:ListItem>
                <asp:ListItem>Sieves Passing</asp:ListItem>
            </asp:DropDownList>
            <asp:Label ID="lblError" Visible="false" ForeColor="Red" runat="server" Text="Please Select value"></asp:Label>
         </div>
        <div class="col-md-4" style="text-align:center">
            <asp:Button ID="btnUpload" class="btn" runat="server" Text="Upload" />
        </div>        
    </div>
    <br />
    <div class="row">
        <div class="col-md-12">
            <asp:GridView ID="GridView1" runat="server">
             </asp:GridView>
        </div>
    </div>

</asp:Content>
