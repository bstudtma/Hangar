using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace Hangar
{
    public partial class EventMappingWindow : Window
    {
        public string SimVarName { get; }

        // Use the nested type from MainWindow
        public ObservableCollection<MainWindow.EventMapping> EditedMappings { get; }

        public EventMappingWindow(string simVarName, System.Collections.Generic.List<MainWindow.EventMapping> initial)
        {
            InitializeComponent();
            SimVarName = simVarName;
            TitleText.Text = $"Event mappings for: {simVarName}";
            EditedMappings = new ObservableCollection<MainWindow.EventMapping>(initial.Select(m => new MainWindow.EventMapping
            {
                MatchValue = m.MatchValue,
                EventName = m.EventName,
                Parameter = m.Parameter
            }));
            DataContext = this;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            // Validate: drop empty rows (no event name)
            var cleaned = EditedMappings.Where(m => !string.IsNullOrWhiteSpace(m.EventName)).ToList();
            EditedMappings.Clear();
            foreach (var m in cleaned) EditedMappings.Add(m);
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        // Expose as List for consumer convenience
        public System.Collections.Generic.List<MainWindow.EventMapping> EditedMappingsList => EditedMappings.ToList();
    }
}

