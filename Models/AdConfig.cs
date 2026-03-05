using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;

namespace SCCMAdPrep.Models;

/// <summary>
/// Configuration for an OU
/// </summary>
public partial class OuEntry : ObservableObject
{
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    [ObservableProperty] private bool _isSelected = true;
}

/// <summary>
/// Configuration for a security group
/// </summary>
public partial class GroupEntry : ObservableObject
{
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    [ObservableProperty] private string _scope = "Global";
    public string Category { get; set; } = "Security";
    [ObservableProperty] private bool _isSelected = true;
}

/// <summary>
/// Configuration for a service account
/// </summary>
public partial class ServiceAccountEntry : ObservableObject
{
    public string Name { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    [ObservableProperty] private bool _isSelected = true;
    public string? GeneratedPassword { get; set; }
}

/// <summary>
/// Log entry for output
/// </summary>
public class LogEntry
{
    public DateTime Timestamp { get; set; } = DateTime.Now;
    public LogLevel Level { get; set; } = LogLevel.Info;
    public string Message { get; set; } = string.Empty;

    public string FormattedTime => Timestamp.ToString("HH:mm:ss");

    public string Prefix => Level switch
    {
        LogLevel.Ok => "[OK]",
        LogLevel.Info => "[i] ",
        LogLevel.Warn => "[!] ",
        LogLevel.Error => "[X] ",
        LogLevel.DryRun => "[~] ",
        LogLevel.Section => "====",
        _ => "    "
    };
}

public enum LogLevel
{
    Ok,
    Info,
    Warn,
    Error,
    DryRun,
    Section,
    Text
}

/// <summary>
/// Represents an OU node in the AD tree browser
/// </summary>
public class OuTreeItem
{
    public string Name { get; set; } = string.Empty;
    public string DistinguishedName { get; set; } = string.Empty;
    public ObservableCollection<OuTreeItem> Children { get; set; } = new();
    public bool IsExpanded { get; set; }
}

/// <summary>
/// Detected domain information
/// </summary>
public class DomainInfo
{
    public string DnsRoot { get; set; } = string.Empty;
    public string NetBiosName { get; set; } = string.Empty;
    public string DistinguishedName { get; set; } = string.Empty;
    public bool IsDetected { get; set; }
}
