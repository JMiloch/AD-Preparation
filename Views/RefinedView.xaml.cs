using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using Wpf.Ui.Controls;
using SCCMAdPrep.ViewModels;

namespace SCCMAdPrep.Views;

public partial class RefinedView : FluentWindow
{
    private readonly FrameworkElement[] _pages;
    private readonly ListBox[] _navLists;

    public RefinedView()
    {
        // Set ViewModel as DataContext
        DataContext = new MainViewModel();

        InitializeComponent();

        _pages = new FrameworkElement[]
        {
            PagePrerequisites,
            PageOUs,
            PageContainer,
            PageGroups,
            PageAccounts,
            PageGMSA,
            PageExtras,
            PageLog,
            PageAbout
        };

        _navLists = new[] { NavList, NavList2, NavList3, NavList4, NavList5 };

        // Select first item on startup
        NavList.SelectedIndex = 0;

        // Auto-scroll for log + auto-navigate to log
        if (DataContext is MainViewModel vm)
        {
            vm.PropertyChanged += OnViewModelPropertyChanged;
        }
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(MainViewModel.LogText))
        {
            // Auto-scroll to bottom
            LogScrollViewer.ScrollToBottom();
        }
        else if (e.PropertyName == nameof(MainViewModel.IsRunning))
        {
            // Auto-navigate to Log page when execution starts
            if (DataContext is MainViewModel vm && vm.IsRunning)
            {
                NavigateToLog();
            }
        }
    }

    /// <summary>
    /// Navigates to the log page and sets the navigation correctly
    /// </summary>
    private void NavigateToLog()
    {
        // Deselect all other NavLists
        foreach (var list in _navLists)
            list.SelectedIndex = -1;

        // Select log item in NavList4
        NavList4.SelectedIndex = 0;

        ShowPage("Log");
    }

    private void ShowPage(string tag)
    {
        foreach (var page in _pages)
            page.Visibility = Visibility.Collapsed;

        switch (tag)
        {
            case "Prerequisites": PagePrerequisites.Visibility = Visibility.Visible; break;
            case "OUs": PageOUs.Visibility = Visibility.Visible; break;
            case "Container": PageContainer.Visibility = Visibility.Visible; break;
            case "Groups": PageGroups.Visibility = Visibility.Visible; break;
            case "Accounts": PageAccounts.Visibility = Visibility.Visible; break;
            case "GMSA": PageGMSA.Visibility = Visibility.Visible; break;
            case "Extras": PageExtras.Visibility = Visibility.Visible; break;
            case "Log": PageLog.Visibility = Visibility.Visible; break;
            case "About": PageAbout.Visibility = Visibility.Visible; break;
        }
    }

    private void ClearOtherSelections(ListBox current)
    {
        foreach (var list in _navLists)
        {
            if (list != current)
                list.SelectedIndex = -1;
        }
    }

    private void NavList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (NavList.SelectedItem is not System.Windows.Controls.ListBoxItem item) return;
        ClearOtherSelections(NavList);
        ShowPage(item.Tag?.ToString() ?? "");
    }

    private void NavList2_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (NavList2.SelectedItem is not System.Windows.Controls.ListBoxItem item) return;
        ClearOtherSelections(NavList2);
        ShowPage(item.Tag?.ToString() ?? "");
    }

    private void NavList3_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (NavList3.SelectedItem is not System.Windows.Controls.ListBoxItem item) return;
        ClearOtherSelections(NavList3);
        ShowPage(item.Tag?.ToString() ?? "");
    }

    private void NavList4_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (NavList4.SelectedItem is not System.Windows.Controls.ListBoxItem item) return;
        ClearOtherSelections(NavList4);
        ShowPage(item.Tag?.ToString() ?? "");
    }

    private void NavList5_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (NavList5.SelectedItem is not System.Windows.Controls.ListBoxItem item) return;
        ClearOtherSelections(NavList5);
        ShowPage(item.Tag?.ToString() ?? "");
    }
}
