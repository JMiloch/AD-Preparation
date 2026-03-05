using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.IO;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using SCCMAdPrep.Models;

namespace SCCMAdPrep.Services;

/// <summary>
/// Active Directory operations for SCCM preparation.
/// Uses System.DirectoryServices for read operations and
/// powershell.exe (Windows PS 5.1) for AD cmdlets.
/// </summary>
public class AdService
{
    private readonly LogService _log;
    private DomainInfo? _domain;
    private DirectoryEntry? _sccmComputer;

    public AdService(LogService log)
    {
        _log = log;
    }

    /// <summary>
    /// Indicates whether the domain has already been detected
    /// </summary>
    public bool IsDomainDetected => _domain?.IsDetected == true;

    /// <summary>
    /// Attempts to find the SCCM server in AD (if not already done)
    /// </summary>
    public void TryFindServer(string sccmServer)
    {
        if (_sccmComputer != null) return;
        if (_domain == null) return;
        var name = sccmServer.Split('.')[0];
        try
        {
            _sccmComputer = FindComputer(name);
            if (_sccmComputer != null)
                _log.Info($"SCCM server found: {name}");
        }
        catch
        {
            _log.Warn($"SCCM server '{name}' not found in AD");
        }
    }

    // ========================================================================
    // Detect domain
    // ========================================================================

    public DomainInfo? DetectDomain()
    {
        try
        {
            using var domain = Domain.GetCurrentDomain();
            using var rootDse = new DirectoryEntry("LDAP://RootDSE");
            var defaultNamingContext = rootDse.Properties["defaultNamingContext"].Value?.ToString() ?? "";

            _domain = new DomainInfo
            {
                DnsRoot = domain.Name,
                NetBiosName = GetNetBiosName(defaultNamingContext),
                DistinguishedName = defaultNamingContext,
                IsDetected = true
            };
            return _domain;
        }
        catch
        {
            return null;
        }
    }

    private string GetNetBiosName(string dn)
    {
        try
        {
            using var entry = new DirectoryEntry($"LDAP://CN=Partitions,CN=Configuration,{dn}");
            using var searcher = new DirectorySearcher(entry)
            {
                Filter = $"(&(objectClass=crossRef)(nCName={dn}))"
            };
            searcher.PropertiesToLoad.Add("nETBIOSName");
            var result = searcher.FindOne();
            return result?.Properties["nETBIOSName"][0]?.ToString() ?? "";
        }
        catch { return ""; }
    }

    // ========================================================================
    // OU Tree Browser
    // ========================================================================

    /// <summary>
    /// Loads the OU tree from Active Directory for the browse dialog.
    /// Returns the domain root node with all child OUs (max 5 levels deep).
    /// </summary>
    public OuTreeItem? GetOuTree()
    {
        if (_domain == null || !_domain.IsDetected)
        {
            DetectDomain();
            if (_domain == null || !_domain.IsDetected) return null;
        }

        var root = new OuTreeItem
        {
            Name = _domain.DnsRoot,
            DistinguishedName = _domain.DistinguishedName,
            IsExpanded = true
        };

        try
        {
            using var domainEntry = new DirectoryEntry($"LDAP://{_domain.DistinguishedName}");
            LoadOuChildren(domainEntry, root.Children, 0, 5);
        }
        catch
        {
            // Return root even if children can't be loaded
        }

        return root;
    }

    private void LoadOuChildren(DirectoryEntry parent, System.Collections.ObjectModel.ObservableCollection<OuTreeItem> children, int depth, int maxDepth)
    {
        if (depth >= maxDepth) return;

        try
        {
            using var searcher = new DirectorySearcher(parent)
            {
                Filter = "(objectClass=organizationalUnit)",
                SearchScope = SearchScope.OneLevel
            };
            searcher.PropertiesToLoad.Add("name");
            searcher.PropertiesToLoad.Add("distinguishedName");
            searcher.Sort = new SortOption("name", SortDirection.Ascending);

            foreach (SearchResult sr in searcher.FindAll())
            {
                var item = new OuTreeItem
                {
                    Name = sr.Properties["name"][0]?.ToString() ?? "",
                    DistinguishedName = sr.Properties["distinguishedName"][0]?.ToString() ?? "",
                    IsExpanded = depth < 1  // Auto-expand first level
                };

                using var childEntry = sr.GetDirectoryEntry();
                LoadOuChildren(childEntry, item.Children, depth + 1, maxDepth);

                children.Add(item);
            }
        }
        catch
        {
            // Silently skip inaccessible OUs
        }
    }

    // ========================================================================
    // Prerequisites
    // ========================================================================

    public async Task<bool> RunPrerequisites(string sccmServer, bool checkAd, bool checkDomain,
        bool checkServer, bool checkSchema, bool dryRun)
    {
        _log.Section("CHECKING PREREQUISITES");
        bool ok = true;

        // AD module check
        if (checkAd)
        {
            _log.Info("Checking AD access and PowerShell module...");
            try
            {
                // Primary: Test DirectoryServices LDAP access
                using var rootDse = new DirectoryEntry("LDAP://RootDSE");
                var ctx = rootDse.Properties["defaultNamingContext"].Value?.ToString();
                if (string.IsNullOrEmpty(ctx)) throw new Exception("LDAP not reachable");
                _log.Ok("LDAP access successful (RootDSE reachable)");

                // Secondary: Test PS ActiveDirectory module via powershell.exe
                _log.Info("Testing PowerShell ActiveDirectory module...");
                var psResult = await RunPsCommand(
                    "Import-Module ActiveDirectory -ErrorAction Stop; Write-Output 'OK'");
                if (psResult?.Contains("OK") == true)
                {
                    _log.Ok("ActiveDirectory module loaded (powershell.exe)");
                }
                else
                {
                    _log.Warn($"PS ActiveDirectory module issue - output: '{psResult ?? "(empty)"}'");
                    _log.Info("LDAP fallback available, execution will continue");
                }
            }
            catch (Exception ex)
            {
                _log.Error($"AD access NOT possible: {ex.Message}");
                ok = false;
            }
        }
        else
        {
            _log.Info("AD module check skipped (disabled)");
        }

        if (checkDomain && ok)
        {
            _log.Info("Checking domain connectivity...");
            _domain = DetectDomain();
            if (_domain?.IsDetected == true)
            {
                _log.Ok($"Domain detected: {_domain.DnsRoot} ({_domain.NetBiosName})");
                _log.Info($"Distinguished Name: {_domain.DistinguishedName}");
            }
            else
            {
                _log.Error("Domain could not be detected!");
                ok = false;
            }
        }
        else if (!checkDomain)
        {
            _log.Info("Domain check skipped (disabled)");
            // Still detect domain for subsequent steps
            if (_domain == null)
            {
                _domain = DetectDomain();
                if (_domain?.IsDetected == true)
                    _log.Info($"Domain automatically detected: {_domain.DnsRoot}");
            }
        }

        if (checkServer && ok)
        {
            _log.Info($"Searching for SCCM server '{sccmServer}' in AD...");
            var serverName = sccmServer.Split('.')[0];
            try
            {
                _sccmComputer = FindComputer(serverName);
                if (_sccmComputer != null)
                    _log.Ok($"SCCM server found: {serverName}");
                else
                    throw new Exception("Not found");
            }
            catch
            {
                _log.Warn($"SCCM server '{serverName}' NOT found in AD!");
                _sccmComputer = null;
                if (!dryRun) ok = false;
            }
        }
        else if (!checkServer)
        {
            _log.Info("Server check skipped (disabled)");
        }

        if (checkSchema && ok)
        {
            _log.Info("Checking Schema Admin permissions...");
            try
            {
                var result = await RunPsCommand(@"
                    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
                    $m = Get-ADGroupMember 'Schema Admins' -ErrorAction SilentlyContinue |
                         Where-Object { $_.SamAccountName -eq $env:USERNAME }
                    if ($m) { 'YES' } else { 'NO' }
                ");
                if (result?.Contains("YES") == true)
                    _log.Ok("Current user is Schema Admin");
                else
                    _log.Warn("Current user is NOT Schema Admin");
            }
            catch (Exception ex)
            {
                _log.Warn($"Schema Admin check failed: {ex.Message}");
            }
        }
        else if (!checkSchema)
        {
            _log.Info("Schema Admin check skipped (disabled)");
        }

        _log.Info($"Prerequisite check: {(ok ? "PASSED" : "FAILED")}");
        return ok;
    }

    // ========================================================================
    // OU Structure
    // ========================================================================

    public async Task CreateOuStructure(string rootOu, string? parentDn, IEnumerable<OuEntry> subOus, bool disableProtection, bool dryRun)
    {
        if (_domain == null || string.IsNullOrWhiteSpace(_domain.DistinguishedName))
        {
            _log.Error("Domain not detected - OU creation not possible");
            return;
        }
        _log.Section("CREATING OU STRUCTURE");

        var baseDn = !string.IsNullOrWhiteSpace(parentDn) ? parentDn : _domain.DistinguishedName;
        var rootDn = $"OU={rootOu},{baseDn}";
        if (!string.IsNullOrWhiteSpace(parentDn))
            _log.Info($"Custom parent path: {parentDn}");

        // Root OU
        _log.Info($"Checking root OU: {rootDn}");
        if (OuExists(rootDn))
        {
            _log.Info($"Root OU '{rootOu}' already exists");
        }
        else if (dryRun)
        {
            _log.DryRun($"Would create OU: {rootDn}");
        }
        else
        {
            var protectStr = disableProtection ? "$false" : "$true";
            _log.Info($"Creating root OU '{rootOu}'{(disableProtection ? " (deletion protection disabled)" : "")}...");
            var result = await RunPsCommand($@"
                Import-Module ActiveDirectory -ErrorAction Stop
                New-ADOrganizationalUnit -Name '{rootOu}' -Path '{baseDn}' -Description 'SCCM Management OU' -ProtectedFromAccidentalDeletion {protectStr}
                Write-Output 'CREATED'
            ");
            if (result?.Contains("CREATED") == true)
                _log.Ok($"Root OU created: {rootDn}");
            else
                _log.Error($"Root OU error - output: '{result ?? "(empty)"}'");
        }

        // Sub-OUs
        var selectedOus = subOus.Where(s => s.IsSelected).ToList();
        _log.Info($"{selectedOus.Count} sub-OUs selected");
        if (disableProtection)
            _log.Warn("Deletion protection disabled - OUs can be deleted without restriction");

        foreach (var sub in selectedOus)
        {
            var subDn = $"OU={sub.Name},{rootDn}";
            if (OuExists(subDn))
            {
                _log.Info($"Sub-OU '{sub.Name}' already exists");
            }
            else if (dryRun)
            {
                _log.DryRun($"Would create OU: {sub.Name} ({sub.Description})");
            }
            else
            {
                var subProtectStr = disableProtection ? "$false" : "$true";
                _log.Info($"Creating sub-OU '{sub.Name}'...");
                var result = await RunPsCommand($@"
                    Import-Module ActiveDirectory -ErrorAction Stop
                    New-ADOrganizationalUnit -Name '{sub.Name}' -Path '{rootDn}' -Description '{sub.Description}' -ProtectedFromAccidentalDeletion {subProtectStr}
                    Write-Output 'CREATED'
                ");
                if (result?.Contains("CREATED") == true)
                    _log.Ok($"Sub-OU created: {sub.Name}");
                else
                    _log.Error($"Sub-OU '{sub.Name}' error - output: '{result ?? "(empty)"}'");
            }
        }
    }

    // ========================================================================
    // System Management Container
    // ========================================================================

    public async Task CreateSystemManagementContainer(string sccmServer, bool setAcl, bool dryRun)
    {
        if (_domain == null || string.IsNullOrWhiteSpace(_domain.DistinguishedName))
        {
            _log.Error("Domain not detected - container creation not possible");
            return;
        }
        _log.Section("SYSTEM MANAGEMENT CONTAINER");

        var systemDn = $"CN=System,{_domain.DistinguishedName}";
        var sysManDn = $"CN=System Management,{systemDn}";

        _log.Info($"Checking container: {sysManDn}");
        if (AdObjectExists(sysManDn))
        {
            _log.Info("System Management container already exists");
        }
        else if (dryRun)
        {
            _log.DryRun($"Would create container: {sysManDn}");
        }
        else
        {
            try
            {
                _log.Info("Creating System Management container...");
                using var parent = new DirectoryEntry($"LDAP://{systemDn}");
                var container = parent.Children.Add("CN=System Management", "container");
                container.Properties["description"].Value = "SCCM System Management Container";
                container.CommitChanges();
                _log.Ok($"Container created: {sysManDn}");
            }
            catch (Exception ex)
            {
                _log.Error($"Container error: {ex.Message}");
            }
        }

        // Set ACL
        if (setAcl && _sccmComputer != null)
        {
            var serverName = sccmServer.Split('.')[0];
            if (dryRun)
            {
                _log.DryRun($"Would set Full Control for '{serverName}'");
            }
            else
            {
                _log.Info($"Setting Full Control ACL for '{serverName}'...");
                var result = await RunPsCommand($@"
                    Import-Module ActiveDirectory -ErrorAction Stop
                    $sysManDN = '{sysManDn}'
                    $comp = Get-ADComputer -Identity '{serverName}'
                    $acl = Get-Acl -Path ""AD:\$sysManDN""
                    $sid = New-Object System.Security.Principal.SecurityIdentifier $comp.SID
                    $ace = New-Object System.DirectoryServices.ActiveDirectoryAccessRule(
                        $sid,
                        [System.DirectoryServices.ActiveDirectoryRights]'GenericAll',
                        [System.Security.AccessControl.AccessControlType]'Allow',
                        [System.DirectoryServices.ActiveDirectorySecurityInheritance]'All'
                    )
                    $acl.AddAccessRule($ace)
                    Set-Acl -Path ""AD:\$sysManDN"" -AclObject $acl
                    Write-Output 'ACL_SET'
                ");
                if (result?.Contains("ACL_SET") == true)
                    _log.Ok($"Full Control set for '{serverName}' (incl. inheritance)");
                else
                    _log.Error($"ACL could not be set - output: '{result ?? "(empty)"}'");
            }
        }
        else if (setAcl)
        {
            _log.Warn("SCCM server not found - ACL must be set manually");
        }
        else
        {
            _log.Info("ACL setting skipped (disabled)");
        }
    }

    // ========================================================================
    // Security Groups
    // ========================================================================

    public async Task CreateSecurityGroups(string rootOu, string? parentDn, IEnumerable<GroupEntry> groups,
        string sccmServer, bool addServerToGroup, bool dryRun)
    {
        if (_domain == null || string.IsNullOrWhiteSpace(_domain.DistinguishedName))
        {
            _log.Error("Domain not detected - groups creation not possible");
            return;
        }

        var selected = groups.Where(g => g.IsSelected).ToList();
        if (!selected.Any())
        {
            _log.Info("No security groups selected - step skipped");
            return;
        }

        _log.Section("CREATING SECURITY GROUPS");
        var baseDn = !string.IsNullOrWhiteSpace(parentDn) ? parentDn : _domain.DistinguishedName;
        var targetPath = $"OU=Groups,OU={rootOu},{baseDn}";
        _log.Info($"Target path: {targetPath}");
        _log.Info($"{selected.Count} groups selected");

        foreach (var g in selected)
        {
            if (GroupExists(g.Name))
            {
                _log.Info($"Group '{g.Name}' already exists");
            }
            else if (dryRun)
            {
                _log.DryRun($"Would create group: {g.Name} ({g.Scope}/{g.Category})");
            }
            else
            {
                _log.Info($"Creating group '{g.Name}'...");
                var result = await RunPsCommand($@"
                    Import-Module ActiveDirectory -ErrorAction Stop
                    New-ADGroup -Name '{g.Name}' -GroupScope {g.Scope} -GroupCategory {g.Category} -Path '{targetPath}' -Description '{g.Description}'
                    Write-Output 'CREATED'
                ");
                if (result?.Contains("CREATED") == true)
                    _log.Ok($"Group created: {g.Name}");
                else
                    _log.Error($"Group '{g.Name}' error - output: '{result ?? "(empty)"}'");
            }
        }

        if (addServerToGroup && _sccmComputer != null)
        {
            var serverName = sccmServer.Split('.')[0];
            if (dryRun)
            {
                _log.DryRun($"Would add '{serverName}' to 'SCCM-Servers'");
            }
            else
            {
                _log.Info($"Adding '{serverName}' to 'SCCM-Servers'...");
                var result = await RunPsCommand($@"
                    Import-Module ActiveDirectory -ErrorAction Stop
                    try {{
                        Add-ADGroupMember -Identity 'SCCM-Servers' -Members (Get-ADComputer '{serverName}') -ErrorAction Stop
                        Write-Output 'ADDED'
                    }} catch {{
                        if ($_.Exception.Message -match 'already a member') {{ Write-Output 'ALREADY' }} else {{ throw }}
                    }}
                ");
                if (result?.Contains("ADDED") == true)
                    _log.Ok($"'{serverName}' added to 'SCCM-Servers'");
                else if (result?.Contains("ALREADY") == true)
                    _log.Info($"'{serverName}' is already a member of 'SCCM-Servers'");
                else
                    _log.Error($"Server group membership error - output: '{result ?? "(empty)"}'");
            }
        }
        else if (addServerToGroup)
        {
            _log.Warn("SCCM server not found - group membership must be set manually");
        }
    }

    // ========================================================================
    // Service Accounts
    // ========================================================================

    public async Task CreateServiceAccounts(string rootOu, string? parentDn, IEnumerable<ServiceAccountEntry> accounts,
        bool generatePasswords, bool dryRun)
    {
        if (_domain == null || string.IsNullOrWhiteSpace(_domain.DistinguishedName))
        {
            _log.Error("Domain not detected - account creation not possible");
            return;
        }

        var selected = accounts.Where(a => a.IsSelected).ToList();
        if (!selected.Any())
        {
            _log.Info("No service accounts selected - step skipped");
            return;
        }

        _log.Section("CREATING SERVICE ACCOUNTS");
        var baseDn = !string.IsNullOrWhiteSpace(parentDn) ? parentDn : _domain.DistinguishedName;
        var targetPath = $"OU=ServiceAccounts,OU={rootOu},{baseDn}";
        _log.Info($"Target path: {targetPath}");
        _log.Info($"{selected.Count} accounts selected");

        var createdPasswords = new List<(string Name, string Password)>();

        foreach (var acc in selected)
        {
            var upn = $"{acc.Name}@{_domain.DnsRoot}";
            if (UserExists(acc.Name))
            {
                _log.Info($"Account '{acc.Name}' already exists");
                continue;
            }
            if (dryRun)
            {
                _log.DryRun($"Would create account: {acc.Name} ({acc.Description})");
                continue;
            }

            _log.Info($"Creating account '{acc.Name}'...");
            var pw = GeneratePassword(24);
            acc.GeneratedPassword = pw;
            createdPasswords.Add((acc.Name, pw));

            var result = await RunPsCommand($@"
                Import-Module ActiveDirectory -ErrorAction Stop
                $pw = ConvertTo-SecureString '{pw}' -AsPlainText -Force
                New-ADUser -Name '{acc.Name}' -SamAccountName '{acc.Name}' -UserPrincipalName '{upn}' -DisplayName '{acc.DisplayName}' -Description '{acc.Description}' -Path '{targetPath}' -AccountPassword $pw -Enabled $true -PasswordNeverExpires $true -CannotChangePassword $true -ChangePasswordAtLogon $false
                Write-Output 'CREATED'
            ");
            if (result?.Contains("CREATED") == true)
            {
                _log.Ok($"Account created: {acc.Name} (UPN: {upn})");
                // NAA deprecation notice
                if (acc.Name.Contains("NAA", StringComparison.OrdinalIgnoreCase))
                    _log.Warn("Note: The Network Access Account (NAA) is deprecated since SCCM CB 2403+. Consider using Enhanced HTTP or PKI instead.");
            }
            else
                _log.Error($"Account '{acc.Name}' error - output: '{result ?? "(empty)"}'");
        }

        if (generatePasswords && createdPasswords.Any())
        {
            _log.Warn("GENERATED PASSWORDS:");
            foreach (var (name, pw) in createdPasswords)
                _log.Warn($"  {name}: {pw}");
            SavePasswordFile(createdPasswords);
        }
    }

    // ========================================================================
    // gMSA
    // ========================================================================

    public async Task CreateGmsa(string sccmServer, bool createKds, bool labMode,
        bool createSql, string gmsaDns, bool dryRun)
    {
        if (!createKds && !createSql)
        {
            _log.Info("gMSA completely disabled - step skipped");
            return;
        }
        _log.Section("gMSA ACCOUNTS");

        if (createKds)
        {
            if (dryRun)
            {
                var modeStr = labMode ? "(lab mode: -EffectiveTime -10h)" : "(production mode)";
                _log.DryRun($"Would create KDS Root Key {modeStr}");
            }
            else
            {
                var modeStr = labMode ? "lab mode" : "production mode";
                _log.Info($"Checking/creating KDS Root Key ({modeStr})...");

                var kdsScript = labMode
                    ? @"
                        $existing = Get-KdsRootKey -ErrorAction SilentlyContinue
                        if ($existing) { Write-Output 'EXISTS' }
                        else {
                            Add-KdsRootKey -EffectiveTime ((Get-Date).AddHours(-10))
                            Write-Output 'CREATED_LAB'
                        }
                    "
                    : @"
                        $existing = Get-KdsRootKey -ErrorAction SilentlyContinue
                        if ($existing) { Write-Output 'EXISTS' }
                        else {
                            Add-KdsRootKey -EffectiveImmediately
                            Write-Output 'CREATED_PROD'
                        }
                    ";

                var result = await RunPsCommand(kdsScript);
                if (result?.Contains("EXISTS") == true)
                    _log.Ok("KDS Root Key already exists");
                else if (result?.Contains("CREATED_LAB") == true)
                    _log.Ok("KDS Root Key created (lab mode: immediately available)");
                else if (result?.Contains("CREATED_PROD") == true)
                {
                    _log.Ok("KDS Root Key created (production mode)");
                    _log.Warn("Available after DC replication (~10 hours)");
                }
                else
                    _log.Error($"KDS Root Key could not be created - output: '{result ?? "(empty)"}'");
            }
        }
        else
        {
            _log.Info("KDS Root Key skipped (disabled)");
        }

        if (createSql)
        {
            if (string.IsNullOrWhiteSpace(gmsaDns) && _domain != null)
                gmsaDns = $"gmsa-SQL.{_domain.DnsRoot}";

            var serverName = sccmServer.Split('.')[0];

            if (dryRun)
            {
                _log.DryRun($"Would create gMSA: gmsa-SQL (DNS: {gmsaDns})");
            }
            else
            {
                _log.Info($"Checking/creating gMSA 'gmsa-SQL' (DNS: {gmsaDns})...");
                var result = await RunPsCommand($@"
                    Import-Module ActiveDirectory -ErrorAction Stop
                    try {{
                        Get-ADServiceAccount -Identity 'gmsa-SQL' -ErrorAction Stop | Out-Null
                        Write-Output 'EXISTS'
                    }} catch {{
                        New-ADServiceAccount -Name 'gmsa-SQL' -DNSHostName '{gmsaDns}' -PrincipalsAllowedToRetrieveManagedPassword '{serverName}$' -Enabled $true
                        Write-Output 'CREATED'
                    }}
                ");
                if (result?.Contains("EXISTS") == true)
                    _log.Info("gMSA 'gmsa-SQL' already exists");
                else if (result?.Contains("CREATED") == true)
                {
                    _log.Ok($"gMSA created: gmsa-SQL (DNS: {gmsaDns})");
                    _log.Info($"On target server: Install-ADServiceAccount -Identity gmsa-SQL");
                }
                else
                    _log.Error($"gMSA could not be created - output: '{result ?? "(empty)"}'");
            }
        }
        else
        {
            _log.Info("gMSA SQL skipped (disabled)");
        }
    }

    // ========================================================================
    // Schema Extension (extadsch.exe)
    // ========================================================================

    public async Task RunExtadsch(bool dryRun)
    {
        _log.Section("AD SCHEMA EXTENSION");

        // Find extadsch.exe next to our exe
        var exeDir = AppDomain.CurrentDomain.BaseDirectory;
        var extadschPath = Path.Combine(exeDir, "extadsch.exe");

        if (!File.Exists(extadschPath))
        {
            _log.Warn($"extadsch.exe not found in: {exeDir}");
            _log.Info("Please copy extadsch.exe from SMSSETUP\\BIN\\X64\\ to the application folder.");
            return;
        }

        _log.Info($"Found: {extadschPath}");

        if (dryRun)
        {
            _log.DryRun("Would run extadsch.exe to extend the AD schema");
            return;
        }

        _log.Info("Running extadsch.exe (this may take a moment)...");
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = extadschPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var process = Process.Start(psi);
            if (process == null)
            {
                _log.Error("extadsch.exe could not be started");
                return;
            }

            var output = await process.StandardOutput.ReadToEndAsync();
            var error = await process.StandardError.ReadToEndAsync();

            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(120));
            try
            {
                await process.WaitForExitAsync(cts.Token);
            }
            catch (OperationCanceledException)
            {
                _log.Error("extadsch.exe timeout after 120 seconds");
                try { process.Kill(true); } catch { }
                return;
            }

            if (!string.IsNullOrWhiteSpace(output))
                _log.Info($"extadsch output: {output.Trim()}");
            if (!string.IsNullOrWhiteSpace(error))
                _log.Error($"extadsch error: {error.Trim()}");

            _log.Info($"Exit code: {process.ExitCode}");

            // Check extadsch.log
            await CheckExtadschLog();
        }
        catch (Exception ex)
        {
            _log.Error($"extadsch.exe failed: {ex.Message}");
        }
    }

    private async Task CheckExtadschLog()
    {
        var logPath = @"C:\extadsch.log";
        _log.Info($"Checking log: {logPath}");

        await Task.Delay(1000); // Short delay for file write

        if (!File.Exists(logPath))
        {
            _log.Warn("extadsch.log not found at C:\\");
            return;
        }

        try
        {
            var content = File.ReadAllText(logPath);
            if (content.Contains("Successfully extended the Active Directory schema", StringComparison.OrdinalIgnoreCase))
            {
                _log.Ok("Schema extension successful: 'Successfully extended the Active Directory schema'");
            }
            else
            {
                _log.Warn("Schema extension may have failed - success message not found in log");
                // Show log content for troubleshooting
                foreach (var line in content.Split('\n', StringSplitOptions.RemoveEmptyEntries))
                {
                    var trimmed = line.Trim();
                    if (!string.IsNullOrEmpty(trimmed))
                        _log.Info($"  extadsch.log: {trimmed}");
                }
            }
        }
        catch (Exception ex)
        {
            _log.Error($"Could not read extadsch.log: {ex.Message}");
        }
    }

    // ========================================================================
    // Schema check (read-only)
    // ========================================================================

    public async Task CheckSchema()
    {
        _log.Section("AD SCHEMA CHECK");
        _log.Info("Checking AD schema for SCCM extension (mS-SMS-Site-Code)...");
        try
        {
            var result = await RunPsCommand(@"
                Import-Module ActiveDirectory -ErrorAction Stop
                $schema = (Get-ADRootDSE).schemaNamingContext
                $attr = Get-ADObject -SearchBase $schema -Filter ""Name -eq 'mS-SMS-Site-Code'"" -ErrorAction SilentlyContinue
                if ($attr) { 'EXTENDED' } else { 'NOT_EXTENDED' }
            ");
            if (result?.Contains("EXTENDED") == true && !result.Contains("NOT_EXTENDED"))
                _log.Ok("AD schema extended for SCCM (mS-SMS-Site-Code found)");
            else
            {
                _log.Warn("AD schema NOT extended for SCCM");
                _log.Info(@"Run: SMSSETUP\BIN\X64\extadsch.exe");
            }
        }
        catch (Exception ex)
        {
            _log.Warn($"Schema check failed: {ex.Message}");
        }
    }

    public void ShowDnsHints(string sccmServer)
    {
        _log.Section("DNS CONFIGURATION");
        _log.Info("Ensure the following DNS records exist:");
        _log.Info($"  A/AAAA Record for: {sccmServer}");
        _log.Info("  SRV Record: _mssms_mp (Management Point)");
        _log.Info("  Reverse DNS (PTR) for SCCM server IP");
    }

    public void ShowFirewallHints()
    {
        _log.Section("FIREWALL PORT REFERENCE");
        _log.Info("Required ports for SCCM:");
        _log.Info("  Client -> MP:     TCP 80, 443");
        _log.Info("  Client -> DP:     TCP 80, 443");
        _log.Info("  Client -> SUP:    TCP 8530, 8531 (WSUS)");
        _log.Info("  Site <-> Site:    TCP 4022 (SQL SSB)");
        _log.Info("  Console -> Site:  TCP 135, 445 + RPC");
        _log.Info("  Client Push:      TCP 445, 135");
        _log.Info("  SQL Server:       TCP 1433");
        _log.Info("  PXE/OSD:          UDP 67, 68, 69, 4011");
    }

    // ========================================================================
    // PowerShell via powershell.exe (Windows PS 5.1)
    // Hosted PS7 Runspace does not find AD modules on DC!
    // ========================================================================

    private async Task<string?> RunPsCommand(string script)
    {
        try
        {
            var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(script));

            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand {encoded}",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            var sw = Stopwatch.StartNew();

            using var process = Process.Start(psi);
            if (process == null)
            {
                _log.Error("powershell.exe could not be started");
                return null;
            }

            // Read stdout and stderr in parallel (prevents deadlock with full buffers)
            var outputTask = process.StandardOutput.ReadToEndAsync();
            var errorTask = process.StandardError.ReadToEndAsync();

            // Timeout: 60 seconds
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(60));
            try
            {
                await process.WaitForExitAsync(cts.Token);
            }
            catch (OperationCanceledException)
            {
                _log.Error("PowerShell timeout after 60 seconds - process will be terminated");
                try { process.Kill(true); } catch { }
                return null;
            }

            var output = await outputTask;
            var error = await errorTask;
            sw.Stop();

            if (!string.IsNullOrWhiteSpace(error))
            {
                // Filter out CLIXML progress messages (PowerShell progress bars written to stderr)
                foreach (var line in error.Split('\n', StringSplitOptions.RemoveEmptyEntries))
                {
                    var trimmed = line.Trim();
                    if (string.IsNullOrEmpty(trimmed)) continue;
                    if (trimmed.StartsWith("#< CLIXML", StringComparison.Ordinal)) continue;
                    if (trimmed.StartsWith("<Objs", StringComparison.Ordinal) && trimmed.Contains("S=\"progress\"")) continue;
                    _log.Error($"PS: {trimmed}");
                }
            }

            if (process.ExitCode != 0 && string.IsNullOrWhiteSpace(output))
            {
                _log.Warn($"PowerShell ExitCode: {process.ExitCode} (duration: {sw.ElapsedMilliseconds}ms)");
            }

            return output?.Trim();
        }
        catch (Exception ex)
        {
            _log.Error($"PowerShell error: {ex.Message}");
            return null;
        }
    }

    // ========================================================================
    // Helper methods (System.DirectoryServices - no PS needed)
    // ========================================================================

    private DirectoryEntry? FindComputer(string name)
    {
        if (_domain == null) return null;
        using var root = new DirectoryEntry($"LDAP://{_domain.DistinguishedName}");
        using var searcher = new DirectorySearcher(root)
        { Filter = $"(&(objectClass=computer)(cn={name}))" };
        var result = searcher.FindOne();
        return result?.GetDirectoryEntry();
    }

    private bool OuExists(string dn)
    {
        try { using var e = new DirectoryEntry($"LDAP://{dn}"); _ = e.Guid; return true; }
        catch { return false; }
    }

    private bool AdObjectExists(string dn) => OuExists(dn);

    private bool GroupExists(string name)
    {
        if (_domain == null) return false;
        using var root = new DirectoryEntry($"LDAP://{_domain.DistinguishedName}");
        using var s = new DirectorySearcher(root) { Filter = $"(&(objectClass=group)(cn={name}))" };
        return s.FindOne() != null;
    }

    private bool UserExists(string samAccount)
    {
        if (_domain == null) return false;
        using var root = new DirectoryEntry($"LDAP://{_domain.DistinguishedName}");
        using var s = new DirectorySearcher(root) { Filter = $"(&(objectClass=user)(sAMAccountName={samAccount}))" };
        return s.FindOne() != null;
    }

    private static string GeneratePassword(int length = 24)
    {
        const string chars = "abcdefghkmnpqrstuvwxyzABCDEFGHKMNPQRSTUVWXYZ23456789!@#$%&*";
        var rng = new Random();
        return new string(Enumerable.Range(0, length).Select(_ => chars[rng.Next(chars.Length)]).ToArray());
    }

    private void SavePasswordFile(List<(string Name, string Password)> passwords)
    {
        try
        {
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (string.IsNullOrEmpty(desktop))
                desktop = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

            var ts = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            var path = Path.Combine(desktop, $"SCCM-Passwords_{ts}.txt");

            var lines = new List<string>
            {
                "============================================================",
                "  SCCM Service Account Passwords",
                $"  Created: {DateTime.Now:dd.MM.yyyy HH:mm:ss}",
                $"  Domain:   {_domain?.DnsRoot}",
                "============================================================",
                ""
            };

            foreach (var (name, pw) in passwords)
            {
                lines.Add($"Account:   {name}");
                lines.Add($"Password:  {pw}");
                lines.Add("");
            }

            lines.Add("============================================================");
            lines.Add("  IMPORTANT: Delete this file after transferring passwords!");
            lines.Add("============================================================");

            File.WriteAllLines(path, lines);
            _log.Ok($"Passwords saved: {path}");
            _log.Warn("Delete file after transferring to password manager!");
        }
        catch (Exception ex)
        {
            _log.Error($"Password file error: {ex.Message}");
        }
    }
}
