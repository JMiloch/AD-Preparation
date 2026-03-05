using System.Windows;
using System.Windows.Controls;
using SCCMAdPrep.Models;

namespace SCCMAdPrep.Views;

/// <summary>
/// OU Browser dialog - displays AD OU tree for selection.
/// Supports two modes:
/// 1. Select existing OU (leave "New OU Name" empty)
/// 2. Create new OU at selected location (enter a name)
/// </summary>
public partial class OuBrowserDialog : Window
{
    private OuTreeItem? _selectedItem;

    /// <summary>
    /// The name of the selected/new OU (e.g. "Management")
    /// </summary>
    public string? SelectedOuName { get; private set; }

    /// <summary>
    /// The full DN of the selected OU
    /// </summary>
    public string? SelectedOuDn { get; private set; }

    /// <summary>
    /// The DN of the parent where the root OU lives or should be created.
    /// Null when the domain root node itself was selected.
    /// </summary>
    public string? ParentDn { get; private set; }

    /// <summary>
    /// Whether a new OU name was entered (create mode)
    /// </summary>
    public bool IsCreateNew { get; private set; }

    public OuBrowserDialog()
    {
        InitializeComponent();
    }

    /// <summary>
    /// Loads the OU tree into the TreeView
    /// </summary>
    public void LoadTree(OuTreeItem rootItem)
    {
        OuTree.Items.Clear();
        OuTree.Items.Add(rootItem);
    }

    private void OuTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        if (e.NewValue is OuTreeItem item)
        {
            _selectedItem = item;
            UpdatePreview();
        }
    }

    private void NewOuName_TextChanged(object sender, TextChangedEventArgs e)
    {
        UpdatePreview();
    }

    private void UpdatePreview()
    {
        if (_selectedItem == null)
        {
            SelectedPathText.Text = "Click an OU in the tree above";
            SelectedPathText.FontStyle = FontStyles.Italic;
            SelectedPathText.Foreground = new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#9090A4"));
            SelectButton.IsEnabled = false;
            return;
        }

        var newName = NewOuNameBox.Text?.Trim();

        if (!string.IsNullOrEmpty(newName))
        {
            // Create new mode: show new OU path under selected parent
            var previewDn = $"OU={newName},{_selectedItem.DistinguishedName}";
            PreviewLabel.Text = "NEW OU WILL BE CREATED";
            SelectedPathText.Text = previewDn;
            SelectedPathText.FontStyle = FontStyles.Normal;
            SelectedPathText.Foreground = new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#006ABB"));
            SelectButton.IsEnabled = true;
        }
        else
        {
            // Use existing mode: show selected OU
            PreviewLabel.Text = "USING EXISTING OU";
            SelectedPathText.Text = _selectedItem.DistinguishedName;
            SelectedPathText.FontStyle = FontStyles.Normal;
            SelectedPathText.Foreground = new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#1A1A2E"));
            SelectButton.IsEnabled = true;
        }
    }

    private void Select_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedItem == null) return;

        var newName = NewOuNameBox.Text?.Trim();

        if (!string.IsNullOrEmpty(newName))
        {
            // Create new OU mode
            IsCreateNew = true;
            SelectedOuName = newName;
            ParentDn = _selectedItem.DistinguishedName;
            SelectedOuDn = $"OU={newName},{_selectedItem.DistinguishedName}";
        }
        else
        {
            // Use existing OU mode
            IsCreateNew = false;
            SelectedOuName = _selectedItem.Name;
            SelectedOuDn = _selectedItem.DistinguishedName;

            // Extract parent DN from selected OU's DN
            // e.g. "OU=SCCM,OU=IT,DC=contoso,DC=local" -> parent = "OU=IT,DC=contoso,DC=local"
            var dn = _selectedItem.DistinguishedName;
            var commaIdx = dn.IndexOf(',');
            if (commaIdx > 0)
                ParentDn = dn.Substring(commaIdx + 1);
            else
                ParentDn = null; // domain root selected
        }

        DialogResult = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
