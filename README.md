# SCCM AD Preparation Tool

WPF-Anwendung zur automatisierten Active-Directory-Vorbereitung fuer SCCM/MECM-Deployments.

## Ueberblick

Das Tool automatisiert alle AD-relevanten Vorbereitungsschritte fuer eine SCCM/MECM-Installation:

- **Prerequisites pruefen** - AD-Zugriff, Domain-Konnektivitaet, SCCM-Server, Schema-Admin-Rechte
- **OU-Struktur anlegen** - Konfigurierbarer Root-OU mit Sub-OUs (Users, Groups, Servers, Workstations, ...)
- **System Management Container** - Erstellt `CN=System Management` mit ACL-Konfiguration
- **Security Groups** - Benutzer- und Geraetegruppen (Global/DomainLocal)
- **Service Accounts** - AD-Benutzerkonten mit automatischer Passwort-Generierung (24 Zeichen)
- **gMSA** - Group Managed Service Accounts und KDS Root Key (Lab/Produktion)
- **Schema Extension** - Fuehrt `extadsch.exe` aus und validiert das Ergebnis
- **DNS & Firewall Hints** - Referenz-Dokumentation fuer Netzwerkkonfiguration

## Features

- **Dry-Run-Modus** - Simuliert alle Aenderungen ohne Ausfuehrung
- **Einzelschritt-Ausfuehrung** - Jeder Abschnitt kann einzeln ausgefuehrt werden
- **OU-Browser** - Interaktiver Dialog zum Durchsuchen der AD-OU-Struktur
- **Passwort-Export** - Generierte Passwoerter werden als Datei auf dem Desktop gespeichert
- **Log-Export** - Ausfuehrungsprotokolle als `.txt` oder `.log` speichern
- **Domain-Auto-Erkennung** - Erkennt die lokale AD-Domain automatisch beim Start
- **Select All** - Alle Optionen mit einem Klick aktivieren/deaktivieren

## Voraussetzungen

- Windows (WPF-Anwendung)
- .NET 9.0 Runtime (oder Self-Contained-Publish nutzen)
- Active-Directory-Zugriff (Domain-joined oder DC-Konnektivitaet)
- PowerShell 5.1 (Windows PowerShell) mit ActiveDirectory-Modul
- RSAT-Tools empfohlen
- Administrator-Rechte fuer die meisten Operationen
- `extadsch.exe` (aus SCCM-Medium, `SMSSETUP\BIN\X64\`) im Anwendungsordner fuer Schema-Extension

## Build

```bash
dotnet build
```

## Publish (Self-Contained)

```bash
dotnet publish -c Release -r win-x64 --self-contained true -o ./publish
```

Die publizierte EXE (~142 MB) beinhaltet die .NET Runtime und laeuft ohne separate Installation.

## Projektstruktur

```
SCCMAdPrep/
  SCCMAdPrep.csproj        Projektdatei (.NET 9, WPF)
  App.xaml / App.xaml.cs    Application-Einstiegspunkt (WPF-UI Light Theme)
  app.ico                   Anwendungs-Icon
  Models/
    AdConfig.cs             Datenmodelle (OuEntry, GroupEntry, ServiceAccountEntry, ...)
  Services/
    AdService.cs            AD-Operationen (LDAP + PowerShell)
    LogService.cs           Thread-sicheres Logging (ConcurrentQueue)
  ViewModels/
    MainViewModel.cs        MVVM ViewModel, orchestriert die gesamte UI-Logik
  Views/
    RefinedView.xaml/.cs    Hauptfenster mit Sidebar-Navigation
    OuBrowserDialog.xaml/.cs  OU-Browser-Dialog
```

## Architektur

- **MVVM** mit CommunityToolkit.Mvvm (ObservableProperty, RelayCommand)
- **WPF-UI 3.0** (Fluent-Design, FluentWindow)
- **System.DirectoryServices** fuer LDAP-Leseoperationen
- **powershell.exe** (Windows PS 5.1) fuer AD-Cmdlets (New-ADOrganizationalUnit, New-ADGroup, etc.)
- **Async/Await** durchgehend fuer nicht-blockierende Ausfuehrung

## NuGet-Abhaengigkeiten

| Paket | Version | Zweck |
|-------|---------|-------|
| CommunityToolkit.Mvvm | 8.4.0 | MVVM-Framework |
| WPF-UI | 3.0.5 | Modernes UI-Framework |
| System.DirectoryServices | 9.0.3 | LDAP/AD-Zugriff |
| System.DirectoryServices.AccountManagement | 9.0.3 | AD-Kontenverwaltung |
| System.Management.Automation | 7.5.0 | PowerShell-Integration |

## Standardmaessig erstellte Objekte

### Sub-OUs (unter Root-OU)
Users, Groups, ServiceAccounts, Servers, Workstations, Clients, Administrators

### Security Groups
- `SCCM-Admins` (Full Administrators)
- `SCCM-RemoteConsole` (Remote Console Access)
- `SCCM-Servers` (Site Servers)
- `SCCM-SiteServers` (Site Systems)

### Service Accounts
- `svc-SCCM-Admin` - Full SCCM Administrator
- `svc-SCCM-ClientPush` - Client Push Installation
- `svc-SCCM-DomJoin` - OSD Domain Join
- `svc-SCCM-Reporting` - SSRS Reporting Point
- `svc-SCCM-NAA` - Network Access Account (deprecated seit CB 2403+, standardmaessig deaktiviert)

## Lizenz

Intern - IT Administration
