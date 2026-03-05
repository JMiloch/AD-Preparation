using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using SCCMAdPrep.Models;
using SCCMAdPrep.Services;
using SCCMAdPrep.Views;

namespace SCCMAdPrep.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private readonly LogService _logService;
    private readonly AdService _adService;
    private readonly DispatcherTimer _logTimer;

    public MainViewModel()
    {
        _logService = new LogService();
        _adService = new AdService(_logService);

        // Log timer: polls queue every 80ms
        _logTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(80) };
        _logTimer.Tick += OnLogTimerTick;
        _logTimer.Start();

        // Default values
        InitializeDefaults();

        // Domain auto-detect
        DetectDomainAsync();
    }

    // ========================================================================
    // Properties - Server Config
    // ========================================================================

    [ObservableProperty] private string _sccmServer = "";
    [ObservableProperty] private string _rootOu = "Management";
    [ObservableProperty] private string _domainDisplay = "Domain: detecting...";
    [ObservableProperty] private string _gmsaDns = "";

    /// <summary>
    /// The parent DN where the root OU should be created.
    /// Null = domain root (default). Set by Browse dialog.
    /// </summary>
    private string? _rootOuParentDn;

    // Dynamic path display (reacts to RootOu changes)
    public string OuPathDisplay => string.IsNullOrEmpty(_rootOuParentDn)
        ? $"OU={RootOu},DC=..."
        : $"OU={RootOu},{ShortenDn(_rootOuParentDn)}";
    public string GroupsPathDisplay => $"OU=Groups,OU={RootOu}";
    public string AccountsPathDisplay => $"OU=ServiceAccounts,OU={RootOu}";

    partial void OnRootOuChanged(string value)
    {
        // When user manually edits the name, clear the parent DN (back to domain root)
        if (!_browsingOu)
            _rootOuParentDn = null;
        OnPropertyChanged(nameof(OuPathDisplay));
        OnPropertyChanged(nameof(GroupsPathDisplay));
        OnPropertyChanged(nameof(AccountsPathDisplay));
    }

    private bool _browsingOu;

    /// <summary>
    /// Shortens a DN for display purposes (e.g. "OU=IT,DC=contoso,DC=local" → "OU=IT,DC=...")
    /// </summary>
    private static string ShortenDn(string dn)
    {
        // Find the first DC= component and shorten
        var dcIdx = dn.IndexOf("DC=", StringComparison.OrdinalIgnoreCase);
        if (dcIdx > 0)
            return dn.Substring(0, dcIdx) + "DC=...";
        return dn.Length > 30 ? dn.Substring(0, 30) + "..." : dn;
    }

    // ========================================================================
    // Properties - Execution State
    // ========================================================================

    [ObservableProperty] private bool _isRunning;
    [ObservableProperty] private string _logText = "";
    [ObservableProperty] private string _statusText = "Ready";

    /// <summary>
    /// When IsRunning changes, refresh all commands
    /// </summary>
    partial void OnIsRunningChanged(bool value)
    {
        RunCommand.NotifyCanExecuteChanged();
        DryRunCommand.NotifyCanExecuteChanged();
        RunPrerequisitesOnlyCommand.NotifyCanExecuteChanged();
        RunOusOnlyCommand.NotifyCanExecuteChanged();
        RunContainerOnlyCommand.NotifyCanExecuteChanged();
        RunGroupsOnlyCommand.NotifyCanExecuteChanged();
        RunAccountsOnlyCommand.NotifyCanExecuteChanged();
        RunGmsaOnlyCommand.NotifyCanExecuteChanged();
        RunExtrasOnlyCommand.NotifyCanExecuteChanged();
    }

    // ========================================================================
    // Properties - Checkboxes Prerequisites
    // ========================================================================

    [ObservableProperty] private bool _checkAdModule = true;
    [ObservableProperty] private bool _checkDomain = true;
    [ObservableProperty] private bool _checkSccmServer = true;
    [ObservableProperty] private bool _checkSchemaAdmin = true;

    // ========================================================================
    // Properties - OU Structure
    // ========================================================================

    [ObservableProperty] private bool _createOus = true;
    [ObservableProperty] private bool _disableProtectFromDeletion = false;

    public ObservableCollection<OuEntry> SubOus { get; } = new();

    // ========================================================================
    // Properties - System Management
    // ========================================================================

    [ObservableProperty] private bool _createSysMgmt = true;
    [ObservableProperty] private bool _setSysMgmtAcl = true;

    // ========================================================================
    // Properties - Groups (User + Device separated)
    // ========================================================================

    public ObservableCollection<GroupEntry> UserGroups { get; } = new();
    public ObservableCollection<GroupEntry> DeviceGroups { get; } = new();
    private IEnumerable<GroupEntry> AllGroups => UserGroups.Concat(DeviceGroups);
    [ObservableProperty] private bool _addServerToGroup = true;

    // ========================================================================
    // Properties - Service Accounts
    // ========================================================================

    public ObservableCollection<ServiceAccountEntry> ServiceAccounts { get; } = new();
    [ObservableProperty] private bool _generatePasswords = true;
    [ObservableProperty] private string _customAccountName = "";
    [ObservableProperty] private string _customAccountDescription = "";

    // ========================================================================
    // Properties - gMSA
    // ========================================================================

    [ObservableProperty] private bool _createKdsKey = true;
    [ObservableProperty] private bool _kdsLabMode = false;
    [ObservableProperty] private bool _createGmsaSql = true;

    // ========================================================================
    // Properties - System Management (Schema Extension)
    // ========================================================================

    [ObservableProperty] private bool _extendSchema = false;

    // ========================================================================
    // Properties - Extras
    // ========================================================================

    [ObservableProperty] private bool _checkSchema = true;
    [ObservableProperty] private bool _showDnsHints = true;
    [ObservableProperty] private bool _showFwHints = true;

    // ========================================================================
    // Properties - Select All
    // ========================================================================

    [ObservableProperty] private bool _selectAll = true;

    partial void OnSelectAllChanged(bool value)
    {
        CheckAdModule = value;
        CheckDomain = value;
        CheckSccmServer = value;
        CheckSchemaAdmin = value;
        CreateOus = value;
        foreach (var ou in SubOus) ou.IsSelected = value;
        CreateSysMgmt = value;
        SetSysMgmtAcl = value;
        ExtendSchema = value;
        foreach (var g in UserGroups) g.IsSelected = value;
        foreach (var g in DeviceGroups) g.IsSelected = value;
        AddServerToGroup = value;
        foreach (var a in ServiceAccounts) a.IsSelected = value;
        GeneratePasswords = value;
        CreateKdsKey = value;
        CreateGmsaSql = value;
        CheckSchema = value;
        ShowDnsHints = value;
        ShowFwHints = value;
    }

    // ========================================================================
    // Commands - Global (all steps)
    // ========================================================================

    [RelayCommand(CanExecute = nameof(CanExecute))]
    private async Task RunAsync()
    {
        await ExecuteAsync(dryRun: false);
    }

    [RelayCommand(CanExecute = nameof(CanExecute))]
    private async Task DryRunAsync()
    {
        await ExecuteAsync(dryRun: true);
    }

    private bool CanExecute() => !IsRunning;

    [RelayCommand]
    private void ClearLog()
    {
        LogText = "";
    }

    [RelayCommand]
    private async Task ExportLogAsync()
    {
        var dlg = new SaveFileDialog
        {
            Filter = "Text (*.txt)|*.txt|Log (*.log)|*.log",
            FileName = $"SCCM-AD-Prep_{DateTime.Now:yyyy-MM-dd_HH-mm}.txt"
        };
        if (dlg.ShowDialog() == true)
        {
            File.WriteAllText(dlg.FileName, LogText);

            var dialog = new Wpf.Ui.Controls.MessageBox
            {
                Title = "Export Successful",
                Content = new System.Windows.Controls.TextBlock
                {
                    Text = $"Log saved to:\n{dlg.FileName}",
                    TextWrapping = System.Windows.TextWrapping.Wrap,
                    FontSize = 13
                },
                CloseButtonText = "OK",
                CloseButtonAppearance = Wpf.Ui.Controls.ControlAppearance.Primary
            };
            await dialog.ShowDialogAsync();
        }
    }

    // ========================================================================
    // Commands - Custom Account
    // ========================================================================

    [RelayCommand]
    private void AddCustomAccount()
    {
        if (string.IsNullOrWhiteSpace(CustomAccountName)) return;

        var name = CustomAccountName.Trim();
        var desc = string.IsNullOrWhiteSpace(CustomAccountDescription)
            ? "Custom service account"
            : CustomAccountDescription.Trim();

        // Prefix svc- if not already present
        if (!name.StartsWith("svc-", StringComparison.OrdinalIgnoreCase))
            name = $"svc-{name}";

        ServiceAccounts.Add(new ServiceAccountEntry
        {
            Name = name,
            DisplayName = name,
            Description = desc,
            IsSelected = true
        });

        _logService.Info($"Custom account added: {name}");
        StatusText = $"Account '{name}' added";

        CustomAccountName = "";
        CustomAccountDescription = "";
    }

    // ========================================================================
    // Commands - Browse OU
    // ========================================================================

    [RelayCommand]
    private async Task BrowseOuAsync()
    {
        var previousStatus = StatusText;
        StatusText = "Loading OU structure...";

        var tree = await Task.Run(() => _adService.GetOuTree());

        StatusText = previousStatus;

        if (tree == null)
        {
            var dialog = new Wpf.Ui.Controls.MessageBox
            {
                Title = "Browse Failed",
                Content = new System.Windows.Controls.TextBlock
                {
                    Text = "Could not load OU structure from Active Directory.\nPlease check domain connectivity and RSAT tools.",
                    TextWrapping = System.Windows.TextWrapping.Wrap,
                    FontSize = 13
                },
                CloseButtonText = "OK",
                CloseButtonAppearance = Wpf.Ui.Controls.ControlAppearance.Primary
            };
            await dialog.ShowDialogAsync();
            return;
        }

        var browser = new OuBrowserDialog();
        browser.Owner = Application.Current.MainWindow;
        browser.LoadTree(tree);

        if (browser.ShowDialog() == true && !string.IsNullOrEmpty(browser.SelectedOuName))
        {
            _browsingOu = true;
            RootOu = browser.SelectedOuName;
            _rootOuParentDn = browser.ParentDn;
            _browsingOu = false;
            OnPropertyChanged(nameof(OuPathDisplay));
        }
    }

    // ========================================================================
    // Commands - Individual sections
    // ========================================================================

    [RelayCommand(CanExecute = nameof(CanExecute))]
    private async Task RunPrerequisitesOnly()
    {
        // Only validate server name when the check is active
        if (CheckSccmServer && !await ValidateSccmServerAsync()) return;

        IsRunning = true;
        StatusText = "Checking prerequisites...";

        try
        {
            _logService.Section("=== CHECKING PREREQUISITES ===");
            var ok = await _adService.RunPrerequisites(
                SccmServer, CheckAdModule, CheckDomain, CheckSccmServer, CheckSchemaAdmin, false);
            _logService.Section($"=== RESULT: {(ok ? "PASSED" : "FAILED")} ===");
        }
        catch (Exception ex)
        {
            _logService.Error($"Error: {ex.Message}");
        }
        finally
        {
            IsRunning = false;
            StatusText = "Ready";
        }
    }

    [RelayCommand(CanExecute = nameof(CanExecute))]
    private Task RunOusOnly() => RunSectionAsync("Create OUs", async () =>
    {
        if (CreateOus)
            await _adService.CreateOuStructure(RootOu, _rootOuParentDn, SubOus, DisableProtectFromDeletion, false);
        else
            _logService.Info("OU creation disabled");
    });

    [RelayCommand(CanExecute = nameof(CanExecute))]
    private Task RunContainerOnly() => RunSectionAsync("System Management", async () =>
    {
        if (CreateSysMgmt)
            await _adService.CreateSystemManagementContainer(SccmServer, SetSysMgmtAcl, false);
        else
            _logService.Info("Container disabled");

        if (ExtendSchema)
            await _adService.RunExtadsch(false);
        else
            _logService.Info("Schema Extension disabled");
    }, needsSccmServer: true);

    [RelayCommand(CanExecute = nameof(CanExecute))]
    private Task RunGroupsOnly() => RunSectionAsync("Security Groups", async () =>
    {
        await _adService.CreateSecurityGroups(RootOu, _rootOuParentDn, AllGroups, SccmServer, AddServerToGroup, false);
    }, needsSccmServer: true);

    [RelayCommand(CanExecute = nameof(CanExecute))]
    private Task RunAccountsOnly() => RunSectionAsync("Service Accounts", async () =>
    {
        await _adService.CreateServiceAccounts(RootOu, _rootOuParentDn, ServiceAccounts, GeneratePasswords, false);
    });

    [RelayCommand(CanExecute = nameof(CanExecute))]
    private Task RunGmsaOnly() => RunSectionAsync("gMSA", async () =>
    {
        await _adService.CreateGmsa(SccmServer, CreateKdsKey, KdsLabMode, CreateGmsaSql, GmsaDns, false);
    });

    [RelayCommand(CanExecute = nameof(CanExecute))]
    private Task RunExtrasOnly() => RunSectionAsync("Extras", async () =>
    {
        if (CheckSchema)
            await _adService.CheckSchema();
        if (ShowDnsHints)
            _adService.ShowDnsHints(SccmServer);
        if (ShowFwHints)
            _adService.ShowFirewallHints();
        if (!CheckSchema && !ShowDnsHints && !ShowFwHints)
            _logService.Info("No extras enabled");
    });

    // ========================================================================
    // Section Helper
    // ========================================================================

    /// <summary>
    /// Executes a single section.
    /// Domain is automatically detected if needed.
    /// </summary>
    private async Task RunSectionAsync(string sectionName, Func<Task> action, bool needsSccmServer = false)
    {
        if (IsRunning) return;

        if (needsSccmServer && !await ValidateSccmServerAsync()) return;

        IsRunning = true;
        StatusText = $"{sectionName}...";

        try
        {
            _logService.Section($"=== {sectionName.ToUpper()} ===");

            // Auto-detect domain if needed
            if (!_adService.IsDomainDetected)
            {
                _logService.Info("Auto-detecting domain...");
                var domain = await Task.Run(() => _adService.DetectDomain());
                if (domain?.IsDetected != true)
                {
                    _logService.Error("Domain not detected - please check prerequisites first");
                    return;
                }
                _logService.Ok($"Domain detected: {domain.DnsRoot}");
            }

            // Find SCCM server if needed
            if (needsSccmServer && !string.IsNullOrWhiteSpace(SccmServer))
            {
                _adService.TryFindServer(SccmServer);
            }

            await action();
            _logService.Section("=== COMPLETED ===");
        }
        catch (Exception ex)
        {
            _logService.Error($"Error in {sectionName}: {ex.Message}");
        }
        finally
        {
            IsRunning = false;
            StatusText = "Ready";
        }
    }

    private async Task<bool> ValidateSccmServerAsync()
    {
        if (!string.IsNullOrWhiteSpace(SccmServer)) return true;

        var dialog = new Wpf.Ui.Controls.MessageBox
        {
            Title = "Input Required",
            Content = new System.Windows.Controls.TextBlock
            {
                Text = "Please enter the SCCM server name before proceeding.",
                TextWrapping = System.Windows.TextWrapping.Wrap,
                FontSize = 13
            },
            CloseButtonText = "OK",
            CloseButtonAppearance = Wpf.Ui.Controls.ControlAppearance.Primary
        };
        await dialog.ShowDialogAsync();
        return false;
    }

    // ========================================================================
    // Core Execution (all steps)
    // ========================================================================

    private async Task ExecuteAsync(bool dryRun)
    {
        if (!await ValidateSccmServerAsync()) return;

        IsRunning = true;
        StatusText = dryRun ? "Dry run in progress..." : "Execution in progress...";

        try
        {
            var mode = dryRun ? "DRY RUN" : "EXECUTION";
            _logService.Section($"=== {mode} STARTED ===");
            _logService.Info($"SCCM Server: {SccmServer} | Root OU: {RootOu}");
            if (!string.IsNullOrEmpty(_rootOuParentDn))
                _logService.Info($"Parent path: {_rootOuParentDn}");
            _logService.Info($"Mode: {(dryRun ? "Simulation (no changes)" : "Live execution")}");

            // 1. Prerequisites
            _logService.Info("--- Step 1/10: Prerequisites ---");
            var ok = await _adService.RunPrerequisites(
                SccmServer, CheckAdModule, CheckDomain, CheckSccmServer, CheckSchemaAdmin, dryRun);
            if (!ok && !dryRun)
            {
                _logService.Error("Prerequisites NOT met - aborting!");
                return;
            }

            // 2. OU Structure
            _logService.Info("--- Step 2/10: OU Structure ---");
            if (CreateOus)
                await _adService.CreateOuStructure(RootOu, _rootOuParentDn, SubOus, DisableProtectFromDeletion, dryRun);
            else
                _logService.Info("OU creation disabled - skipped");

            // 3. System Management Container
            _logService.Info("--- Step 3/10: System Management Container ---");
            if (CreateSysMgmt)
                await _adService.CreateSystemManagementContainer(SccmServer, SetSysMgmtAcl, dryRun);
            else
                _logService.Info("System Management Container disabled - skipped");

            // 3b. Schema Extension
            _logService.Info("--- Step 3b/10: AD Schema Extension ---");
            if (ExtendSchema)
                await _adService.RunExtadsch(dryRun);
            else
                _logService.Info("Schema Extension disabled - skipped");

            // 4. Security Groups
            _logService.Info("--- Step 4/10: Security Groups ---");
            await _adService.CreateSecurityGroups(RootOu, _rootOuParentDn, AllGroups, SccmServer, AddServerToGroup, dryRun);

            // 5. Service Accounts
            _logService.Info("--- Step 5/10: Service Accounts ---");
            await _adService.CreateServiceAccounts(RootOu, _rootOuParentDn, ServiceAccounts, GeneratePasswords, dryRun);

            // 6. gMSA
            _logService.Info("--- Step 6/10: gMSA Accounts ---");
            await _adService.CreateGmsa(SccmServer, CreateKdsKey, KdsLabMode, CreateGmsaSql, GmsaDns, dryRun);

            // 7. Schema Check
            _logService.Info("--- Step 7/10: Schema Check ---");
            if (CheckSchema)
                await _adService.CheckSchema();
            else
                _logService.Info("Schema Check disabled - skipped");

            // 8. DNS Hints
            _logService.Info("--- Step 8/10: DNS Hints ---");
            if (ShowDnsHints)
                _adService.ShowDnsHints(SccmServer);
            else
                _logService.Info("DNS Hints disabled - skipped");

            // 9. Firewall Hints
            _logService.Info("--- Step 9/10: Firewall Reference ---");
            if (ShowFwHints)
                _adService.ShowFirewallHints();
            else
                _logService.Info("Firewall Reference disabled - skipped");

            _logService.Section($"=== {mode} COMPLETED ===");
        }
        catch (Exception ex)
        {
            _logService.Error($"Unexpected error: {ex.Message}");
            _logService.Error($"Details: {ex.GetType().Name}");
        }
        finally
        {
            IsRunning = false;
            StatusText = "Ready";
        }
    }

    // ========================================================================
    // Log Timer
    // ========================================================================

    private void OnLogTimerTick(object? sender, EventArgs e)
    {
        while (_logService.TryDequeue(out var entry))
        {
            if (entry == null) continue;
            LogText += $"[{entry.FormattedTime}] {entry.Prefix} {entry.Message}\r\n";
        }
    }

    // ========================================================================
    // Init
    // ========================================================================

    private void InitializeDefaults()
    {
        SubOus.Add(new OuEntry { Name = "Users", Description = "Management users" });
        SubOus.Add(new OuEntry { Name = "Groups", Description = "Security groups" });
        SubOus.Add(new OuEntry { Name = "ServiceAccounts", Description = "Service accounts" });
        SubOus.Add(new OuEntry { Name = "Servers", Description = "Server objects" });
        SubOus.Add(new OuEntry { Name = "Workstations", Description = "Workstations" });
        SubOus.Add(new OuEntry { Name = "Clients", Description = "SCCM-managed clients" });
        SubOus.Add(new OuEntry { Name = "Administrators", Description = "SCCM administrators" });

        // User Groups
        UserGroups.Add(new GroupEntry { Name = "SCCM-Admins", Description = "SCCM Full Administrators" });
        UserGroups.Add(new GroupEntry { Name = "SCCM-RemoteConsole", Description = "Remote console access" });

        // Device Groups
        DeviceGroups.Add(new GroupEntry { Name = "SCCM-Servers", Description = "All SCCM Site Servers" });
        DeviceGroups.Add(new GroupEntry { Name = "SCCM-SiteServers", Description = "All SCCM site systems" });

        ServiceAccounts.Add(new ServiceAccountEntry
        {
            Name = "svc-SCCM-Admin", DisplayName = "SCCM Administrator Account",
            Description = "Full SCCM administrator"
        });
        ServiceAccounts.Add(new ServiceAccountEntry
        {
            Name = "svc-SCCM-ClientPush", DisplayName = "SCCM Client Push Account",
            Description = "Client push installation"
        });
        ServiceAccounts.Add(new ServiceAccountEntry
        {
            Name = "svc-SCCM-DomJoin", DisplayName = "SCCM Domain Join Account",
            Description = "OSD Domain Join"
        });
        ServiceAccounts.Add(new ServiceAccountEntry
        {
            Name = "svc-SCCM-Reporting", DisplayName = "SCCM Reporting Account",
            Description = "SSRS Reporting Point"
        });
        ServiceAccounts.Add(new ServiceAccountEntry
        {
            Name = "svc-SCCM-NAA", DisplayName = "SCCM Network Access Account",
            Description = "Network access account — deprecated since CB 2403+ (use Enhanced HTTP)",
            IsSelected = false
        });
    }

    private async void DetectDomainAsync()
    {
        await Task.Run(() =>
        {
            var domain = _adService.DetectDomain();
            Application.Current?.Dispatcher.Invoke(() =>
            {
                if (domain?.IsDetected == true)
                {
                    DomainDisplay = $"Domain: {domain.DnsRoot} ({domain.NetBiosName})";
                    if (string.IsNullOrWhiteSpace(GmsaDns))
                        GmsaDns = $"gmsa-SQL.{domain.DnsRoot}";
                }
                else
                {
                    DomainDisplay = "Domain: Not detected (no DC / RSAT missing)";
                }
            });
        });
    }
}
