<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="Marek_Ptak._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="jumbotron">
        <h1>Dostawy materiałów</h1>
        

    </div>
    <div class="jumbotron">
        <p class="lead">Naciśnij przycisk "Utwórz plik .xlsx" aby utworzyć nowy plik .xlsx w lokalizacji
            C:\Dostawy .<br/> Plik bedzie zawierał kolumny wypełnione przykładowymi danymi. Jeżeli katalog docelowy nie istnieje, zostanie on utworzony. 
            Jeżeli w tej lokalizacji istnieje już plik o nazwie Dostawy materiałów.xlsx, w raporcie zostanie wyświetlony stosowny komunikat.</p>
        
        <asp:Button ID="Button1" runat="server" Text="Dodaj do pliku .xlsx" OnClick="Button1_Click" />
            
    </div>
    <div class="jumbotron">
        
        <p>Raport</p>
        <asp:Label ID="Label2" runat="server" Text="Brak działań"></asp:Label>
        <asp:Button ID="Button2" runat="server" Text="Otwórz plik .xlsx" OnClick="Button2_Click" />

            
    </div>
   

</asp:Content>
