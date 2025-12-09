using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.VisualBasic; // For Interaction.InputBox prompt
using SimConnect.NET;
using SimConnect.NET.InputEvents;
using SimConnect.NET.SimVar;

namespace Hangar;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    // Cache of mapped standard SimConnect events (name -> client event ID)
    private readonly Dictionary<string, uint> _standardEventIds = new(StringComparer.OrdinalIgnoreCase);
    private uint _nextStandardEventId = 1000; // starting ID for client events we map at runtime
    // Base directory for all plane configs
    private readonly string _configDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "Hangar");

    // Returns the config file path for the currently selected plane
    private string ConfigPath => Path.Combine(_configDirectory, $"simvars-{SanitizePlaneName(SelectedPlane)}.json");

    // Settings file for persisting app preferences (e.g., last-used plane)
    private string SettingsPath => Path.Combine(_configDirectory, "settings.json");

    private readonly JsonSerializerOptions _serializerOptions = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() }
    };

    public ObservableCollection<SimVarEntry> SimVars { get; } = new();
    public ObservableCollection<string> PlaneNames { get; } = new();

    private const string CreateNewPlaneOption = "<Create new…>";

    private string _selectedPlane = string.Empty;
    public string SelectedPlane
    {
        get => _selectedPlane;
        set
        {
            if (string.Equals(_selectedPlane, value, StringComparison.Ordinal))
                return;
            _selectedPlane = value ?? string.Empty;
            OnSelectedPlaneChanged();
        }
    }

    // SimVars required to build a complete SimConnectDataInitPosition
    private static readonly string[] RequiredInitPositionSimVars = new[]
    {
        "PLANE LATITUDE",
        "PLANE LONGITUDE",
        "PLANE ALTITUDE",
        "PLANE PITCH DEGREES",
        "PLANE BANK DEGREES",
        "PLANE HEADING DEGREES TRUE",
        "SIM ON GROUND",
        "AIRSPEED TRUE"
    };

    public MainWindow()
    {
        InitializeComponent();
        DataContext = this;

        InitializePlaneList();

        try
        {
            if (File.Exists(ConfigPath))
            {
                var json = File.ReadAllText(ConfigPath);
                var configItems = JsonSerializer.Deserialize<List<SimVarConfigItem>>(json, _serializerOptions);

                if (configItems != null)
                {
                    ApplyConfigItemsToCollection(configItems);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to load initial config:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        // Attach row context menu programmatically to avoid XAML event binding conflicts
        AttachRowContextMenu();

        // Ensure required init-position simvars are present in the collection and prevent their removal
        EnsureRequiredSimVarsPresent();
        SimVars.CollectionChanged += SimVars_CollectionChanged;
    }

    private void InitializePlaneList()
    {
        try
        {
            Directory.CreateDirectory(_configDirectory);
            PlaneNames.Clear();

            var files = Directory.GetFiles(_configDirectory, "simvars-*.json", SearchOption.TopDirectoryOnly);
            foreach (var file in files)
            {
                var name = Path.GetFileNameWithoutExtension(file);
                if (name.StartsWith("simvars-", StringComparison.OrdinalIgnoreCase))
                {
                    var plane = name.Substring("simvars-".Length);
                    if (!string.IsNullOrWhiteSpace(plane))
                    {
                        PlaneNames.Add(plane);
                    }
                }
            }

            if (PlaneNames.Count == 0)
            {
                PlaneNames.Add("Default");
            }

            // Add the special option at the end
            PlaneNames.Add(CreateNewPlaneOption);

            // Load last used plane from settings, fall back to first real plane
            var desired = LoadLastPlaneFromSettings();
            if (string.IsNullOrWhiteSpace(desired) || !PlaneNames.Contains(desired))
            {
                desired = PlaneNames.FirstOrDefault(p => !string.Equals(p, CreateNewPlaneOption, StringComparison.Ordinal)) ?? "Default";
            }
            SelectedPlane = desired;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to initialize plane list:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            // Fallback to default
            if (!PlaneNames.Contains("Default")) PlaneNames.Add("Default");
            if (!PlaneNames.Contains(CreateNewPlaneOption)) PlaneNames.Add(CreateNewPlaneOption);
            SelectedPlane = "Default";
        }
    }

    private static string SanitizePlaneName(string? name)
    {
        name ??= string.Empty;
        name = name.Trim();
        // Allow only A-Z, a-z, 0-9, space, dash
        var valid = new System.Text.RegularExpressions.Regex("^[A-Za-z0-9 \\ -]+$");
        if (!valid.IsMatch(name))
        {
            // Remove invalid characters
            name = new string(name.Where(ch => char.IsLetterOrDigit(ch) || ch == ' ' || ch == '-').ToArray());
            name = name.Trim();
        }
        if (string.IsNullOrWhiteSpace(name))
        {
            return "Default";
        }
        return name;
    }

    private void OnSelectedPlaneChanged()
    {
        if (string.Equals(_selectedPlane, CreateNewPlaneOption, StringComparison.Ordinal))
        {
            // Prompt for new plane name
            var input = Interaction.InputBox("Enter new plane name (letters, numbers, spaces, dashes):", "Create Plane", "");
            var sanitized = SanitizePlaneName(input);
            if (string.IsNullOrWhiteSpace(input))
            {
                // Revert to first available plane if canceled
                _selectedPlane = PlaneNames.FirstOrDefault(p => !string.Equals(p, CreateNewPlaneOption, StringComparison.Ordinal)) ?? "Default";
                SetPlaneSelectorSelection(_selectedPlane);
                return;
            }
            if (!string.Equals(input.Trim(), sanitized, StringComparison.Ordinal))
            {
                MessageBox.Show("Invalid plane name. Use letters, numbers, spaces, and dashes only.", "Invalid Name", MessageBoxButton.OK, MessageBoxImage.Warning);
                // Revert
                _selectedPlane = PlaneNames.FirstOrDefault(p => !string.Equals(p, CreateNewPlaneOption, StringComparison.Ordinal)) ?? "Default";
                SetPlaneSelectorSelection(_selectedPlane);
                return;
            }

            if (!PlaneNames.Contains(sanitized, StringComparer.Ordinal))
            {
                // Insert before the create option so it stays at end
                PlaneNames.Insert(Math.Max(0, PlaneNames.Count - 1), sanitized);
            }
            _selectedPlane = sanitized;
            SetPlaneSelectorSelection(_selectedPlane);
            EnsureConfigFileExists();
            LoadSelectedPlaneConfigIntoCollection();
            SaveLastPlaneToSettings(_selectedPlane);
            return;
        }

        // Normal selection change: ensure file exists and load it
        EnsureConfigFileExists();
        LoadSelectedPlaneConfigIntoCollection();
        if (!string.Equals(_selectedPlane, CreateNewPlaneOption, StringComparison.Ordinal))
        {
            SaveLastPlaneToSettings(_selectedPlane);
        }
    }

    private void SetPlaneSelectorSelection(string plane)
    {
        if (this.FindName("PlaneSelector") is ComboBox combo)
        {
            combo.SelectedItem = plane;
        }
    }

    private string LoadLastPlaneFromSettings()
    {
        try
        {
            if (!File.Exists(SettingsPath)) return string.Empty;
            var json = File.ReadAllText(SettingsPath);
            var settings = JsonSerializer.Deserialize<AppSettings>(json, _serializerOptions);
            return settings?.LastPlane ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private void SaveLastPlaneToSettings(string plane)
    {
        try
        {
            Directory.CreateDirectory(_configDirectory);
            var settings = new AppSettings { LastPlane = SanitizePlaneName(plane) };
            var json = JsonSerializer.Serialize(settings, _serializerOptions);
            File.WriteAllText(SettingsPath, json);
        }
        catch
        {
            // ignore settings persistence failures
        }
    }

    private class AppSettings
    {
        public string LastPlane { get; set; } = "Default";
    }

    private void LoadSelectedPlaneConfigIntoCollection()
    {
        try
        {
            if (!File.Exists(ConfigPath))
            {
                SimVars.Clear();
                return;
            }

            var json = File.ReadAllText(ConfigPath);
            if (string.IsNullOrWhiteSpace(json))
            {
                SimVars.Clear();
                return;
            }

            List<SimVarConfigItem>? configItems = null;
            try
            {
                configItems = JsonSerializer.Deserialize<List<SimVarConfigItem>>(json, _serializerOptions);
            }
            catch (JsonException)
            {
                // ignore
            }

            if (configItems is null)
            {
                // Attempt legacy
                List<LegacySimVarConfig>? legacyItems = null;
                try
                {
                    legacyItems = JsonSerializer.Deserialize<List<LegacySimVarConfig>>(json, _serializerOptions);
                }
                catch (JsonException)
                {
                }

                if (legacyItems != null)
                {
                    configItems = legacyItems
                        .Where(item => !string.IsNullOrWhiteSpace(item.Name))
                        .Select(item =>
                        {
                            var definition = SimVarRegistry.Get(item.Name.Trim());
                            if (definition != null)
                            {
                                return new SimVarConfigItem
                                {
                                    Name = definition.Name,
                                    Unit = definition.Unit,
                                    DataType = definition.DataType,
                                    IsSettable = definition.IsSettable,
                                    Value = string.Empty
                                };
                            }

                            return new SimVarConfigItem
                            {
                                Name = item.Name.Trim(),
                                Unit = (item.Unit ?? string.Empty).Trim(),
                                DataType = SimConnectDataType.FloatDouble,
                                IsSettable = true,
                                Value = string.Empty
                            };
                        })
                        .ToList();
                }
            }

            if (configItems != null)
            {
                ApplyConfigItemsToCollection(configItems);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to load plane config:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void AttachRowContextMenu()
    {
        // Find the DataGrid by name (defined in XAML as x:Name="SimVarsGrid")
        if (this.FindName("SimVarsGrid") is not DataGrid grid)
            return;

        var rowStyle = new Style(typeof(DataGridRow), grid.RowStyle);
        rowStyle.Setters.Add(new Setter(DataGridRow.ContextMenuProperty, CreateRowContextMenu()));
        grid.RowStyle = rowStyle;
    }

    private ContextMenu CreateRowContextMenu()
    {
        var ctx = new ContextMenu();
        var insertItem = new MenuItem { Header = "Insert row above" };
        insertItem.Click += InsertRowAbove_Click;
        var deleteItem = new MenuItem { Header = "Delete row" };
        deleteItem.Click += DeleteRow_Click;
        var applyRowItem = new MenuItem { Header = "Apply this row" };
        applyRowItem.Click += ApplyRow_Click;
        var mapEventItem = new MenuItem { Header = "Event Mapping" };
        mapEventItem.Click += EventMapping_Click;
        ctx.Items.Add(insertItem);
        ctx.Items.Add(deleteItem);
        ctx.Items.Add(applyRowItem);
        ctx.Items.Add(new Separator());
        ctx.Items.Add(mapEventItem);
        return ctx;
    }

    // Guard to avoid re-entrancy when programmatically modifying the collection
    private bool _suppressCollectionChanged;

    private void EnsureRequiredSimVarsPresent()
    {
        if (_suppressCollectionChanged) return;

        _suppressCollectionChanged = true;
        try
        {
            // Add any missing required simvars to the collection
            foreach (var name in RequiredInitPositionSimVars)
            {
                var existing = SimVars.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
                if (existing is null)
                {
                    var definition = SimVarRegistry.Get(name);
                    var entry = new SimVarEntry
                    {
                        Name = definition?.Name ?? name,
                        Unit = definition?.Unit ?? string.Empty,
                        DataType = definition?.DataType ?? SimConnectDataType.FloatDouble,
                        IsSettable = definition?.IsSettable ?? true,
                        Value = string.Empty
                    };

                    entry.PropertyChanged += SimVarEntry_PropertyChanged;
                    SimVars.Add(entry);
                }
            }
        }
        finally
        {
            _suppressCollectionChanged = false;
        }
    }

    private async void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (!TryBuildConfig(out var configItems, out var validationMessage))
            {
                MessageBox.Show(validationMessage!, "Validation Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (configItems.Count == 0)
            {
                var result = MessageBox.Show(
                    "No SimVars were provided. Save an empty config file?",
                    "Confirm Save",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result != MessageBoxResult.Yes)
                {
                    return;
                }
            }

            var warnings = new List<string>();
            SimConnectClient? client = null;

            try
            {
                client = new SimConnectClient("Hangar Config");
                await client.ConnectAsync(cancellationToken: CancellationToken.None).ConfigureAwait(true);

                foreach (var item in configItems)
                {
                    var definition = SimVarRegistry.Get(item.Name);
                    var simVarName = (definition?.Name ?? item.Name).Trim();
                    var simVarUnit = (definition?.Unit ?? item.Unit ?? string.Empty).Trim();
                    var simVarDataType = definition?.DataType ?? SimConnectDataType.FloatDouble;
                    var isSettable = definition?.IsSettable ?? true;

                    item.Name = simVarName;
                    item.Unit = simVarUnit;
                    item.DataType = simVarDataType;
                    item.IsSettable = isSettable;

                    try
                    {
                        var value = await GetSimVarValueAsync(client, simVarName, simVarUnit, simVarDataType, CancellationToken.None).ConfigureAwait(true);
                        var formattedValue = FormatValue(value, simVarDataType);
                        item.Value = formattedValue;

                        var entry = SimVars.FirstOrDefault(s => string.Equals(s.Name, simVarName, StringComparison.OrdinalIgnoreCase));
                        if (entry != null)
                        {
                            entry.Unit = simVarUnit;
                            entry.DataType = simVarDataType;
                            entry.IsSettable = isSettable;
                            entry.Value = formattedValue;
                        }
                    }
                    catch (NotSupportedException ex)
                    {
                        warnings.Add(ex.Message);
                        item.Value = string.Empty;
                    }
                }
            }
            catch (SimConnectException ex)
            {
                MessageBox.Show($"Failed to read SimVars from SimConnect:\n{ex.Message}", "SimConnect Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unexpected error while communicating with SimConnect:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            finally
            {
                if (client != null)
                {
                    try
                    {
                        if (client.IsConnected)
                        {
                            await client.DisconnectAsync().ConfigureAwait(true);
                        }
                    }
                    catch
                    {
                        // ignore disconnect errors
                    }
                    client.Dispose();
                }
            }

            Directory.CreateDirectory(_configDirectory);

            var json = JsonSerializer.Serialize(configItems, _serializerOptions);
            File.WriteAllText(ConfigPath, json);

            
            if (warnings.Count > 0)
            {
                var message = $"Config saved to:\n{ConfigPath}";
                message += "\n\nWarnings:\n" + string.Join('\n', warnings);
                MessageBox.Show(message, "Save Complete", MessageBoxButton.OK, warnings.Count == 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);
            }

            // Clear the warnings for the next operation
            warnings.Clear();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to save config:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void ApplyButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            // UI/perf: disable Apply while running and show busy cursor
            var __applyButton = this.FindName("ApplyButton") as Button;
            object? __applyOriginalContent = __applyButton?.Content;
            var __applyWasEnabled = __applyButton?.IsEnabled ?? true;
            var __sw = Stopwatch.StartNew();
            if (__applyButton != null)
            {
                __applyButton.IsEnabled = false;
                __applyButton.Content = "Applying…";
            }
            Mouse.OverrideCursor = Cursors.Wait;

            if (!File.Exists(ConfigPath))
            {
                MessageBox.Show("No config file found. Save one first.", "Missing Config", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var json = File.ReadAllText(ConfigPath);
            if (string.IsNullOrWhiteSpace(json))
            {
                SimVars.Clear();
                MessageBox.Show("Config file was empty. Nothing to load.", "Load Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            List<SimVarConfigItem>? configItems = null;
            try
            {
                configItems = JsonSerializer.Deserialize<List<SimVarConfigItem>>(json, _serializerOptions);
            }
            catch (JsonException)
            {
                // Fall back to legacy format below
            }

            if (configItems is null)
            {
                List<LegacySimVarConfig>? legacyItems = null;
                try
                {
                    legacyItems = JsonSerializer.Deserialize<List<LegacySimVarConfig>>(json, _serializerOptions);
                }
                catch (JsonException)
                {
                    // Ignore and treat as invalid
                }

                if (legacyItems is null)
                {
                    MessageBox.Show("The config file format is not recognized.", "Load Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                configItems = legacyItems
                    .Where(item => !string.IsNullOrWhiteSpace(item.Name))
                    .Select(item =>
                    {
                        var definition = SimVarRegistry.Get(item.Name.Trim());
                        if (definition != null)
                        {
                            return new SimVarConfigItem
                            {
                                Name = definition.Name,
                                Unit = definition.Unit,
                                DataType = definition.DataType,
                                IsSettable = definition.IsSettable,
                                Value = string.Empty
                            };
                        }

                        return new SimVarConfigItem
                        {
                            Name = item.Name.Trim(),
                            Unit = (item.Unit ?? string.Empty).Trim(),
                            DataType = SimConnectDataType.FloatDouble,
                            IsSettable = true,
                            Value = string.Empty
                        };
                    })
                    .ToList();
            }

            var warnings = new List<string>();
            var appliedCount = 0;
            SimConnectClient? client = null;

            try
            {
                client = new SimConnectClient("Hangar Config");
                await client.ConnectAsync(cancellationToken: CancellationToken.None).ConfigureAwait(true);
                Debug.WriteLine($"[Apply] Connected in {__sw.Elapsed.TotalMilliseconds:F0} ms");

                var position = new SimConnect.NET.SimConnectDataInitPosition();
                // Special handling bucket for VELOCITY WORLD X/Y/Z to send together
                var velocityWorld = new SimConnectDataVelocityWorld();
                bool hasVX = false, hasVY = false, hasVZ = false;

                foreach (var item in configItems)
                {
                    var definition = SimVarRegistry.Get(item.Name);
                    var simVarName = (definition?.Name ?? item.Name).Trim();
                    var simVarUnit = (definition?.Unit ?? item.Unit ?? string.Empty).Trim();
                    var simVarDataType = definition?.DataType ?? SimConnectDataType.FloatDouble;
                    var isSettable = definition?.IsSettable ?? true;

                    item.Name = simVarName;
                    item.Unit = simVarUnit;
                    item.DataType = simVarDataType;
                    item.IsSettable = isSettable;

                    if (string.IsNullOrWhiteSpace(item.Value))
                    {
                        warnings.Add($"SimVar '{simVarName}' has no stored value and was skipped.");
                        continue;
                    }

                    switch (simVarName.ToUpperInvariant())
                    {
                        case "PLANE LATITUDE":
                            {
                                if (double.TryParse(item.Value.Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var v))
                                {
                                    position.Latitude = v;
                                }
                                else
                                {
                                    warnings.Add($"Invalid value for '{simVarName}': '{item.Value}'.");
                                }
                                isSettable = false;
                                continue;
                            }
                        case "PLANE LONGITUDE":
                            {
                                if (double.TryParse(item.Value.Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var v))
                                {
                                    position.Longitude = v;
                                }
                                else
                                {
                                    warnings.Add($"Invalid value for '{simVarName}': '{item.Value}'.");
                                }
                                isSettable = false;
                                continue;
                            }
                        case "PLANE ALTITUDE":
                            {
                                if (double.TryParse(item.Value.Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var v))
                                {
                                    position.Altitude = v;
                                }
                                else
                                {
                                    warnings.Add($"Invalid value for '{simVarName}': '{item.Value}'.");
                                }
                                isSettable = false;
                                continue;
                            }
                        case "PLANE PITCH DEGREES":
                            {
                                if (double.TryParse(item.Value.Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var v))
                                {
                                    position.Pitch = v;
                                }
                                else
                                {
                                    warnings.Add($"Invalid value for '{simVarName}': '{item.Value}'.");
                                }
                                isSettable = false;
                                continue;
                            }
                        case "PLANE BANK DEGREES":
                            {
                                if (double.TryParse(item.Value.Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var v))
                                {
                                    position.Bank = v;
                                }
                                else
                                {
                                    warnings.Add($"Invalid value for '{simVarName}': '{item.Value}'.");
                                }
                                isSettable = false;
                                continue;
                            }
                        case "PLANE HEADING DEGREES TRUE":
                            {
                                if (double.TryParse(item.Value.Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var v))
                                {
                                    position.Heading = v;
                                }
                                else
                                {
                                    warnings.Add($"Invalid value for '{simVarName}': '{item.Value}'.");
                                }
                                isSettable = false;
                                continue;
                            }
                        case "SIM ON GROUND":
                            {
                                var t = item.Value.Trim();
                                if (bool.TryParse(t, out var b))
                                {
                                    position.OnGround = b ? 1u : 0u;
                                }
                                else if (uint.TryParse(t, NumberStyles.Integer, CultureInfo.InvariantCulture, out var u))
                                {
                                    position.OnGround = u;
                                }
                                else if (int.TryParse(t, NumberStyles.Integer, CultureInfo.InvariantCulture, out var i))
                                {
                                    position.OnGround = i != 0 ? 1u : 0u;
                                }
                                else
                                {
                                    warnings.Add($"Invalid value for '{simVarName}': '{item.Value}'.");
                                }
                                isSettable = false;
                                continue;
                            }
                        case "AIRSPEED TRUE":
                            {
                                var t = item.Value.Trim();
                                if (uint.TryParse(t, NumberStyles.Integer, CultureInfo.InvariantCulture, out var u))
                                {
                                    position.Airspeed = u;
                                }
                                else if (int.TryParse(t, NumberStyles.Integer, CultureInfo.InvariantCulture, out var i))
                                {
                                    // Accept negative sentinel values by clamping to 0 if they occur; the SDK may handle special cases differently.
                                    position.Airspeed = (uint)Math.Max(0, i);
                                }
                                else if (double.TryParse(t, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var d))
                                {
                                    position.Airspeed = (uint)Math.Max(0, (int)Math.Round(d));
                                }
                                else
                                {
                                    warnings.Add($"Invalid value for '{simVarName}': '{item.Value}'.");
                                }
                                isSettable = false;
                                continue;
                            }
                        case "VELOCITY WORLD X":
                            {
                                if (double.TryParse(item.Value.Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var v))
                                {
                                    velocityWorld.X = v;
                                    hasVX = true;
                                }
                                else
                                {
                                    warnings.Add($"Invalid value for '{simVarName}': '{item.Value}'.");
                                }
                                isSettable = false;
                                continue;
                            }
                        case "VELOCITY WORLD Y":
                            {
                                if (double.TryParse(item.Value.Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var v))
                                {
                                    velocityWorld.Y = v;
                                    hasVY = true;
                                }
                                else
                                {
                                    warnings.Add($"Invalid value for '{simVarName}': '{item.Value}'.");
                                }
                                isSettable = false;
                                continue;
                            }
                        case "VELOCITY WORLD Z":
                            {
                                if (double.TryParse(item.Value.Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var v))
                                {
                                    velocityWorld.Z = v;
                                    hasVZ = true;
                                }
                                else
                                {
                                    warnings.Add($"Invalid value for '{simVarName}': '{item.Value}'.");
                                }
                                isSettable = false;
                                continue;
                            }
                    }
                }

                // If on the ground, set airspeed to zero
                if (position.OnGround != 0)
                {
                    position.Airspeed = 0;
                }

                // Apply position last
                if (!position.Equals(default(SimConnectDataInitPosition)))
                {
                    await client.SimVars.SetAsync("Initial Position", "", position).ConfigureAwait(true);
                    //await client.SimVars.SetAsync(position).ConfigureAwait(true);
                    Debug.WriteLine($"[Apply] Initial Position set at {__sw.Elapsed.TotalMilliseconds:F0} ms");
                }

                // Give the simulator a short moment to process the initial position before applying other SimVars
                await Task.Delay(500).ConfigureAwait(true);
                Debug.WriteLine($"[Apply] Post-position delay reached at {__sw.Elapsed.TotalMilliseconds:F0} ms");
                
                // Apply VELOCITY WORLD (X/Y/Z) together if all components provided
                if (hasVX && hasVY && hasVZ)
                {
                    // If the aircraft is on the ground, zero any configured world velocity components
                    if (position.OnGround != 0)
                    {
                        velocityWorld.X = 0.0;
                        velocityWorld.Y = 0.0;
                        velocityWorld.Z = 0.0;
                    }
                    await client.SimVars.SetAsync(velocityWorld).ConfigureAwait(true);
                    Debug.WriteLine($"[Apply] VelocityWorld set at {__sw.Elapsed.TotalMilliseconds:F0} ms");
                }

                // Prepare input event map only if any mapping is present
                Dictionary<string, InputEventDescriptor>? eventByName2 = null;
                if (configItems.Any(ci => ci.EventMappings != null && ci.EventMappings.Count > 0))
                {
                    var descriptors2 = await client.InputEvents.EnumerateInputEventsAsync(CancellationToken.None).ConfigureAwait(true);
                    // descriptors2 = descriptors2
                    //     .Where(d => !string.IsNullOrWhiteSpace(d.Name))
                    //     .GroupBy(d => d.Name.Trim(), StringComparer.OrdinalIgnoreCase)
                    //     .Select(g => g.First())
                    //     .ToArray();
                    eventByName2 = new Dictionary<string, InputEventDescriptor>(StringComparer.OrdinalIgnoreCase);
                    foreach (var descriptor in descriptors2)
                    {
                        var key = (descriptor.Name ?? string.Empty).Trim();
                        if (key.Length == 0)
                        {
                            continue;
                        }

                        // Ignore duplicates by keeping the first descriptor encountered for each key.
                        if (!eventByName2.ContainsKey(key))
                        {
                            eventByName2[key] = descriptor;
                        }
                    }
                    Debug.WriteLine($"[Apply] Enumerated {eventByName2.Count} input events at {__sw.Elapsed.TotalMilliseconds:F0} ms");
                }

                foreach (var item in configItems)
                {
                    var definition = SimVarRegistry.Get(item.Name);
                    var simVarName = (definition?.Name ?? item.Name).Trim();
                    var simVarUnit = (definition?.Unit ?? item.Unit ?? string.Empty).Trim();
                    var simVarDataType = definition?.DataType ?? SimConnectDataType.FloatDouble;
                    var isSettable = definition?.IsSettable ?? true;

                    item.Name = simVarName;
                    item.Unit = simVarUnit;
                    item.DataType = simVarDataType;
                    item.IsSettable = isSettable;

                    if (string.IsNullOrWhiteSpace(item.Value))
                    {
                        warnings.Add($"SimVar '{simVarName}' has no stored value and was skipped.");
                        continue;
                    }

                    switch (simVarName.ToUpperInvariant())
                    {
                        case "PLANE LATITUDE":
                        case "PLANE LONGITUDE":
                        case "PLANE ALTITUDE":
                        case "PLANE PITCH DEGREES":
                        case "PLANE BANK DEGREES":
                        case "PLANE HEADING DEGREES TRUE":
                        case "SIM ON GROUND":
                        case "AIRSPEED TRUE":
                            {
                                // These are applied together via SimConnectDataInitPosition above
                                continue;
                            }
                        case "VELOCITY WORLD X":
                        case "VELOCITY WORLD Y":
                        case "VELOCITY WORLD Z":
                            {
                                // These are applied together via SimConnectDataVelocityWorld above
                                continue;
                            }
                    }

                    // If there is an event mapping matching the configured value, use it
                    if (item.EventMappings != null && item.EventMappings.Count > 0)
                    {
                        var match = item.EventMappings.FirstOrDefault(m => string.Equals((m.MatchValue ?? string.Empty).Trim(), item.Value.Trim(), StringComparison.OrdinalIgnoreCase));
                        if (match != null)
                        {
                            if (string.IsNullOrWhiteSpace(match.EventName))
                            {
                                warnings.Add($"Event mapping for '{simVarName}' has an empty event name and was skipped.");
                            }
                            else if (eventByName2 == null || !eventByName2.TryGetValue(match.EventName.Trim(), out var desc))
                            {
                                // Fallback: treat as a standard SimConnect event and transmit it to the user aircraft
                                try
                                {
                                    var sent = await TryTransmitStandardEventAsync(client, match.EventName.Trim(), match.Parameter, warnings).ConfigureAwait(true);
                                    if (sent)
                                    {
                                        appliedCount++;
                                        continue; // Event sent via TransmitClientEventAsync, skip SimVar setting
                                    }
                                    else
                                    {
                                        // If we failed to send as a standard event, record a warning and skip further processing for this item
                                        warnings.Add($"Standard event '{match.EventName}' could not be transmitted.");
                                        continue;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    warnings.Add($"Failed to transmit standard event '{match.EventName}': {ex.Message}");
                                    continue;
                                }
                            }
                            else
                            {
                                try
                                {
                                    await client.InputEvents.SetInputEventAsync(desc.Hash, match.Parameter, CancellationToken.None).ConfigureAwait(true);
                                    appliedCount++;
                                    continue; // Do not attempt to set the SimVar directly when a mapping was applied
                                }
                                catch (Exception ex)
                                {
                                    warnings.Add($"Failed to send input event '{match.EventName}': {ex.Message}");
                                    continue;
                                }
                            }
                        }

                        // Percent-range interpolation: when unit is percent and there are exactly two mappings, interpolate parameter
                        bool IsPercentUnit(string u) => !string.IsNullOrWhiteSpace(u) && u.IndexOf("percent", StringComparison.OrdinalIgnoreCase) >= 0;
                        static bool TryParsePercentNumber(string text, out double val)
                        {
                            val = 0;
                            if (string.IsNullOrWhiteSpace(text)) return false;
                            var t = text.Trim();
                            if (t.EndsWith("%", StringComparison.Ordinal)) t = t.Substring(0, t.Length - 1);
                            return double.TryParse(t, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, System.Globalization.CultureInfo.InvariantCulture, out val);
                        }

                        if (eventByName2 != null && item.EventMappings.Count == 2 && IsPercentUnit(simVarUnit))
                        {
                            var m0 = item.EventMappings[0];
                            var m1 = item.EventMappings[1];

                            // Require the same input event for interpolation
                            var ev0 = (m0.EventName ?? string.Empty).Trim();
                            var ev1 = (m1.EventName ?? string.Empty).Trim();
                            if (!string.IsNullOrWhiteSpace(ev0) && string.Equals(ev0, ev1, StringComparison.OrdinalIgnoreCase) &&
                                eventByName2!.TryGetValue(ev0, out var interpDesc) &&
                                TryParsePercentNumber(m0.MatchValue, out var v0) &&
                                TryParsePercentNumber(m1.MatchValue, out var v1) &&
                                TryParsePercentNumber(item.Value, out var v))
                            {
                                // Determine min/max by simvar values
                                double simMin = v0, simMax = v1;
                                double pMin = m0.Parameter, pMax = m1.Parameter;
                                if (simMin > simMax)
                                {
                                    (simMin, simMax) = (simMax, simMin);
                                    (pMin, pMax) = (pMax, pMin);
                                }

                                // Clamp v into [simMin, simMax]
                                var clamped = Math.Max(simMin, Math.Min(simMax, v));
                                var span = simMax - simMin;
                                double ratio = span == 0 ? 0.0 : (clamped - simMin) / span;
                                var parameter = pMin + (pMax - pMin) * ratio;

                                try
                                {
                                    await client.InputEvents.SetInputEventAsync(interpDesc.Hash, parameter, CancellationToken.None).ConfigureAwait(true);
                                    appliedCount++;
                                    continue; // handled by interpolated event
                                }
                                catch (Exception ex)
                                {
                                    warnings.Add($"Failed to send interpolated input event '{ev0}': {ex.Message}");
                                    // fall through to normal set
                                }
                            }
                        }
                    }

                    if (!isSettable)
                    {
                        warnings.Add($"SimVar '{simVarName}' is read-only and no event mapping applied; skipped.");
                        continue;
                    }

                    if (!TryParseValue(item.Value, simVarDataType, out var parsedValue, out var parseError))
                    {
                        warnings.Add(parseError ?? $"Failed to parse value for '{simVarName}'.");
                        continue;
                    }

                    try
                    {
                        await SetSimVarValueAsync(client, simVarName, simVarUnit, simVarDataType, parsedValue!, CancellationToken.None).ConfigureAwait(true);
                        appliedCount++;
                    }
                    catch (NotSupportedException ex)
                    {
                        warnings.Add(ex.Message);
                    }


                }

                // Apply PARKING_BRAKE_SET based on stored BRAKE PARKING POSITION, if provided
                // try
                // {
                //     var parkingBrakeItem = configItems.FirstOrDefault(ci => string.Equals(ci.Name, "BRAKE PARKING POSITION", StringComparison.OrdinalIgnoreCase));
                //     if (parkingBrakeItem is not null && !string.IsNullOrWhiteSpace(parkingBrakeItem.Value))
                //     {
                //         // Determine desired state: 0 = released, 1 = set
                //         bool desiredSet = false;
                //         var t = parkingBrakeItem.Value.Trim();
                //         if (bool.TryParse(t, out var b))
                //         {
                //             desiredSet = b;
                //         }
                //         else if (int.TryParse(t, NumberStyles.Integer, CultureInfo.InvariantCulture, out var i))
                //         {
                //             desiredSet = i != 0;
                //         }
                //         else if (double.TryParse(t, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var d))
                //         {
                //             desiredSet = d >= 0.5;
                //         }

                //         // Enumerate input events and look for PARKING_BRAKE_SET
                //         var inputEvents = client.InputEvents;
                //         var descriptors = await inputEvents.EnumerateInputEventsAsync(CancellationToken.None).ConfigureAwait(true);
                //         var pbDescriptor = descriptors.FirstOrDefault(desc => string.Equals(desc.Name, "LANDING_GEAR_PARKINGBRAKE", StringComparison.OrdinalIgnoreCase));

                //         if (pbDescriptor != null)
                //         {
                //             // Most input events use double payloads; send 0.0/1.0
                //             await inputEvents.SetInputEventAsync(pbDescriptor.Hash, desiredSet ? 1.0 : 0.0, CancellationToken.None).ConfigureAwait(true);
                //             appliedCount++;
                //         }
                //         else
                //         {
                //             warnings.Add("Input event 'PARKING_BRAKE_SET' is not available for this aircraft.");
                //         }
                //     }
                // }
                // catch (Exception ex)
                // {
                //     warnings.Add($"Failed to set PARKING_BRAKE_SET: {ex.Message}");
                // }

            }
            catch (SimConnectException ex)
            {
                MessageBox.Show($"Failed to write SimVars via SimConnect:\n{ex.Message}", "SimConnect Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unexpected error while communicating with SimConnect:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            finally
            {
                if (client != null)
                {
                    try
                    {
                        if (client.IsConnected)
                        {
                            await client.DisconnectAsync().ConfigureAwait(true);
                        }
                    }
                    catch
                    {
                        // ignore disconnect errors
                    }
                    client.Dispose();
                }
                // Restore UI state
                Mouse.OverrideCursor = null;
                if (__applyButton != null)
                {
                    __applyButton.Content = __applyOriginalContent;
                    __applyButton.IsEnabled = __applyWasEnabled;
                }
                Debug.WriteLine($"[Apply] Completed in {__sw.Elapsed.TotalMilliseconds:F0} ms (applied {appliedCount} items, warnings={warnings.Count})");
            }

            ApplyConfigItemsToCollection(configItems);

            
            if (warnings.Count > 0)
            {
                var message = $"Config applied to simulator ({appliedCount} SimVar(s) updated).";
                message += "\n\nWarnings:\n" + string.Join('\n', warnings);
                MessageBox.Show(message, "Load Complete", MessageBoxButton.OK, warnings.Count == 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);
            }

            
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to load config:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // Attempt to send a standard (named) SimConnect event by mapping a client event ID and transmitting it.
    // Returns true if the event was successfully transmitted, otherwise false and appends to warnings.
    private async Task<bool> TryTransmitStandardEventAsync(SimConnectClient client, string eventName, double parameter, List<string> warnings)
    {
        if (string.IsNullOrWhiteSpace(eventName))
            return false;

        var name = eventName.Trim();

        // Get or allocate a client event ID for this standard event name
        if (!_standardEventIds.TryGetValue(name, out var eventId))
        {
            eventId = _nextStandardEventId++;

            // Map the client event ID to the simulator's named event via native SimConnect
            int hr = NativeSimConnect.SimConnect_MapClientEventToSimEvent(client.Handle, eventId, name);
            if (hr != 0)
            {
                warnings.Add($"Failed to map standard event '{name}': HRESULT=0x{hr:X8}");
                return false;
            }

            _standardEventIds[name] = eventId;
        }

        // Transmit to the user aircraft (objectId = 0). Many standard events don't require a data payload.
        await client.InputEvents.TransmitClientEventAsync(0, eventId).ConfigureAwait(true);
        return true;
    }

    private void OpenConfigButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            EnsureConfigFileExists();

            var startInfo = new ProcessStartInfo
            {
                FileName = ConfigPath,
                UseShellExecute = true
            };

            Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to open config:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OpenSimVarsDocsButton_Click(object sender, RoutedEventArgs e)
    {
        const string url = "https://docs.flightsimulator.com/msfs2024/html/6_Programming_APIs/SimVars/Simulation_Variables.htm";
        try
        {
            var psi = new ProcessStartInfo(url)
            {
                UseShellExecute = true
            };
            Process.Start(psi);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to open documentation:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // Context menu: insert a new row above the clicked row
    private void InsertRowAbove_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement fe)
            return;

        if (fe.DataContext is not SimVarEntry rowItem)
            return;

        var index = SimVars.IndexOf(rowItem);
        if (index < 0) index = SimVars.Count; // fallback

        var newEntry = new SimVarEntry
        {
            Name = string.Empty,
            Unit = string.Empty,
            Value = string.Empty,
            DataType = SimConnectDataType.FloatDouble,
            IsSettable = true
        };

        _suppressCollectionChanged = true;
        try
        {
            SimVars.Insert(index, newEntry);
        }
        finally
        {
            _suppressCollectionChanged = false;
        }
    }

    // Context menu: delete the clicked row (unless it's a required init-position simvar)
    private void DeleteRow_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement fe)
            return;

        if (fe.DataContext is not SimVarEntry rowItem)
            return;

        if (RequiredInitPositionSimVars.Any(r => string.Equals(r, rowItem.Name, StringComparison.OrdinalIgnoreCase)))
        {
            MessageBox.Show($"'{rowItem.Name}' is required and cannot be removed.", "Required SimVar", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        _suppressCollectionChanged = true;
        try
        {
            SimVars.Remove(rowItem);
        }
        finally
        {
            _suppressCollectionChanged = false;
        }

        SaveSimVarsToFile();
    }

    // Right-click: apply only the clicked row to the simulator
    private async void ApplyRow_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement fe)
            return;

        if (fe.DataContext is not SimVarEntry rowItem)
            return;

        // Disallow applying grouped/compound SimVars individually to avoid inconsistent or dangerous states
        var nameUpper = (rowItem.Name ?? string.Empty).Trim().ToUpperInvariant();
        static bool IsInitPositionVar(string n) => n is "PLANE LATITUDE" or "PLANE LONGITUDE" or "PLANE ALTITUDE" or "PLANE PITCH DEGREES" or "PLANE BANK DEGREES" or "PLANE HEADING DEGREES TRUE" or "SIM ON GROUND" or "AIRSPEED TRUE";
        static bool IsVelocityWorld(string n) => n is "VELOCITY WORLD X" or "VELOCITY WORLD Y" or "VELOCITY WORLD Z";
        if (IsInitPositionVar(nameUpper))
        {
            MessageBox.Show("This SimVar is part of the aircraft's initial position group and can't be applied individually.", "Grouped SimVar", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        if (IsVelocityWorld(nameUpper))
        {
            MessageBox.Show("World velocity components must be applied together (X, Y, Z). Apply all rows instead.", "Grouped SimVar", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        // Build a normalized config item from the row item (similar to TryBuildConfig for a single entry)
        var trimmedName = (rowItem.Name ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(trimmedName))
        {
            MessageBox.Show("SimVar name is empty.", "Cannot Apply", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var definition = SimVarRegistry.Get(trimmedName);
        var configItem = new SimVarConfigItem
        {
            Name = definition?.Name ?? trimmedName,
            Unit = definition?.Unit ?? (rowItem.Unit ?? string.Empty).Trim(),
            DataType = definition?.DataType ?? SimConnectDataType.FloatDouble,
            IsSettable = definition?.IsSettable ?? true,
            Value = rowItem.Value ?? string.Empty,
            EventMappings = rowItem.EventMappings?.Select(m => new EventMapping
            {
                MatchValue = m.MatchValue,
                EventName = m.EventName,
                Parameter = m.Parameter
            }).ToList()
        };

        // Validate value presence
        if (string.IsNullOrWhiteSpace(configItem.Value))
        {
            MessageBox.Show($"'{configItem.Name}' has no value to apply.", "Cannot Apply", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var warnings = new List<string>();
        SimConnectClient? client = null;
        var applied = false;

        try
        {
            client = new SimConnectClient("Hangar Config");
            await client.ConnectAsync(cancellationToken: CancellationToken.None).ConfigureAwait(true);

            // If there's an event mapping that matches the current value, prefer it
            if (configItem.EventMappings != null && configItem.EventMappings.Count > 0)
            {
                var match = configItem.EventMappings.FirstOrDefault(m => string.Equals((m.MatchValue ?? string.Empty).Trim(), configItem.Value.Trim(), StringComparison.OrdinalIgnoreCase));
                if (match != null)
                {
                    if (string.IsNullOrWhiteSpace(match.EventName))
                    {
                        warnings.Add($"Event mapping for '{configItem.Name}' has an empty event name and was skipped.");
                    }
                    else
                    {
                        // Try exact input event name first
                        try
                        {
                            var inputEvents = client.InputEvents;
                            var descriptors = await inputEvents.EnumerateInputEventsAsync(CancellationToken.None).ConfigureAwait(true);
                            var eventByName = descriptors.ToDictionary(d => d.Name.Trim(), d => d, StringComparer.OrdinalIgnoreCase);

                            if (eventByName.TryGetValue(match.EventName.Trim(), out var desc))
                            {
                                await inputEvents.SetInputEventAsync(desc.Hash, match.Parameter, CancellationToken.None).ConfigureAwait(true);
                                applied = true;
                            }
                            else
                            {
                                // Fallback to standard event transmit
                                var sent = await TryTransmitStandardEventAsync(client, match.EventName.Trim(), match.Parameter, warnings).ConfigureAwait(true);
                                applied = sent;
                            }
                        }
                        catch (Exception ex)
                        {
                            warnings.Add($"Failed to apply event mapping for '{configItem.Name}': {ex.Message}");
                        }
                    }
                }

                // Percent-range interpolation for single-row apply
                bool IsPercentUnitLocal(string u) => !string.IsNullOrWhiteSpace(u) && u.IndexOf("percent", StringComparison.OrdinalIgnoreCase) >= 0;
                static bool TryParsePercentNumberLocal(string text, out double val)
                {
                    val = 0;
                    if (string.IsNullOrWhiteSpace(text)) return false;
                    var t = text.Trim();
                    if (t.EndsWith("%", StringComparison.Ordinal)) t = t.Substring(0, t.Length - 1);
                    return double.TryParse(t, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, System.Globalization.CultureInfo.InvariantCulture, out val);
                }

                if (!applied && configItem.EventMappings.Count == 2 && IsPercentUnitLocal(configItem.Unit))
                {
                    var m0 = configItem.EventMappings[0];
                    var m1 = configItem.EventMappings[1];
                    var ev0 = (m0.EventName ?? string.Empty).Trim();
                    var ev1 = (m1.EventName ?? string.Empty).Trim();

                    try
                    {
                        var inputEvents = client.InputEvents;
                        var descriptors = await inputEvents.EnumerateInputEventsAsync(CancellationToken.None).ConfigureAwait(true);
                        var eventByName = descriptors.ToDictionary(d => d.Name.Trim(), d => d, StringComparer.OrdinalIgnoreCase);

                        if (!string.IsNullOrWhiteSpace(ev0) && string.Equals(ev0, ev1, StringComparison.OrdinalIgnoreCase) &&
                            eventByName.TryGetValue(ev0, out var interpDesc) &&
                            TryParsePercentNumberLocal(m0.MatchValue, out var v0) &&
                            TryParsePercentNumberLocal(m1.MatchValue, out var v1) &&
                            TryParsePercentNumberLocal(configItem.Value, out var v))
                        {
                            double simMin = v0, simMax = v1;
                            double pMin = m0.Parameter, pMax = m1.Parameter;
                            if (simMin > simMax)
                            {
                                (simMin, simMax) = (simMax, simMin);
                                (pMin, pMax) = (pMax, pMin);
                            }

                            var clamped = Math.Max(simMin, Math.Min(simMax, v));
                            var span = simMax - simMin;
                            double ratio = span == 0 ? 0.0 : (clamped - simMin) / span;
                            var parameter = pMin + (pMax - pMin) * ratio;

                            await inputEvents.SetInputEventAsync(interpDesc.Hash, parameter, CancellationToken.None).ConfigureAwait(true);
                            applied = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        warnings.Add($"Failed to send interpolated input event '{ev0}': {ex.Message}");
                    }
                }
            }

            // If not applied via event mapping, try direct SimVar set
            if (!applied)
            {
                if (!configItem.IsSettable)
                {
                    warnings.Add($"'{configItem.Name}' is read-only. No applicable event mapping found.");
                }
                else if (TryParseValue(configItem.Value, configItem.DataType, out var parsedValue, out var parseError))
                {
                    try
                    {
                        await SetSimVarValueAsync(client, configItem.Name, configItem.Unit, configItem.DataType, parsedValue!, CancellationToken.None).ConfigureAwait(true);
                        applied = true;
                    }
                    catch (Exception ex)
                    {
                        warnings.Add($"Failed to set '{configItem.Name}': {ex.Message}");
                    }
                }
                else
                {
                    warnings.Add(parseError ?? $"Failed to parse value for '{configItem.Name}'.");
                }
            }
        }
        catch (SimConnectException ex)
        {
            MessageBox.Show($"Failed to write SimVar via SimConnect:\n{ex.Message}", "SimConnect Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Unexpected error while communicating with SimConnect:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }
        finally
        {
            if (client != null)
            {
                try
                {
                    if (client.IsConnected)
                    {
                        await client.DisconnectAsync().ConfigureAwait(true);
                    }
                }
                catch
                {
                    // ignore disconnect errors
                }
                client.Dispose();
            }
        }

        if (applied)
        {
            if (warnings.Count > 0)
            {
                MessageBox.Show($"Applied '{configItem.Name}' with warnings:\n" + string.Join('\n', warnings), "Row Applied", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                //MessageBox.Show($"Applied '{configItem.Name}'.", "Row Applied", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        else
        {
            var msg = warnings.Count > 0 ? string.Join('\n', warnings) : "Nothing was applied.";
            MessageBox.Show(msg, "Row Not Applied", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private void EnsureConfigFileExists()
    {
        Directory.CreateDirectory(_configDirectory);

        if (!File.Exists(ConfigPath))
        {
            // Create a config file that includes required init-position simvars plus an empty set for others
            var defaultItems = RequiredInitPositionSimVars
                .Select(name =>
                {
                    var def = SimVarRegistry.Get(name);
                    return new SimVarConfigItem
                    {
                        Name = def?.Name ?? name,
                        Unit = def?.Unit ?? string.Empty,
                        DataType = def?.DataType ?? SimConnectDataType.FloatDouble,
                        IsSettable = def?.IsSettable ?? true,
                        Value = string.Empty
                    };
                })
                .ToArray();

            var json = JsonSerializer.Serialize(defaultItems, _serializerOptions);
            File.WriteAllText(ConfigPath, json);
        }
    }

    private bool TryBuildConfig(out List<SimVarConfigItem> configItems, out string? validationMessage)
    {
        configItems = new List<SimVarConfigItem>();
        validationMessage = null;

        foreach (var entry in SimVars)
        {
            if (string.IsNullOrWhiteSpace(entry.Name))
            {
                continue;
            }

            var trimmedName = entry.Name.Trim();
            var definition = SimVarRegistry.Get(trimmedName);

            if (definition != null)
            {
                entry.Name = definition.Name;
                entry.Unit = definition.Unit;
                entry.DataType = definition.DataType;
                entry.IsSettable = definition.IsSettable;

                configItems.Add(new SimVarConfigItem
                {
                    Name = definition.Name,
                    Unit = definition.Unit,
                    DataType = definition.DataType,
                    IsSettable = definition.IsSettable,
                    Value = entry.Value,
                    EventMappings = entry.EventMappings?.Select(m => new EventMapping
                    {
                        MatchValue = m.MatchValue,
                        EventName = m.EventName,
                        Parameter = m.Parameter
                    }).ToList()
                });

                continue;
            }

            var trimmedUnit = (entry.Unit ?? string.Empty).Trim();
            entry.Name = trimmedName;
            entry.Unit = trimmedUnit;
            entry.DataType = SimConnectDataType.FloatDouble;
            entry.IsSettable = true;

            configItems.Add(new SimVarConfigItem
            {
                Name = entry.Name,
                Unit = trimmedUnit,
                DataType = SimConnectDataType.FloatDouble,
                IsSettable = true,
                Value = entry.Value,
                EventMappings = entry.EventMappings?.Select(m => new EventMapping
                {
                    MatchValue = m.MatchValue,
                    EventName = m.EventName,
                    Parameter = m.Parameter
                }).ToList()
            });
        }

        return true;
    }

    private void ApplyConfigItemsToCollection(IEnumerable<SimVarConfigItem> configItems)
    {
        SimVars.Clear();

        foreach (var item in configItems)
        {
            var definition = SimVarRegistry.Get(item.Name);
            var name = (definition?.Name ?? item.Name).Trim();
            var unit = (definition?.Unit ?? item.Unit ?? string.Empty).Trim();
            var dataType = definition?.DataType ?? SimConnectDataType.FloatDouble;
            var isSettable = definition?.IsSettable ?? true;

            var entry = new SimVarEntry
            {
                Name = name,
                Unit = unit,
                DataType = dataType,
                IsSettable = isSettable,
                Value = item.Value
            };

            // Load persisted event mappings
            if (item.EventMappings != null)
            {
                entry.EventMappings = new ObservableCollection<EventMapping>(item.EventMappings.Select(m => new EventMapping
                {
                    MatchValue = m.MatchValue,
                    EventName = m.EventName,
                    Parameter = m.Parameter
                }));
            }

            entry.PropertyChanged += SimVarEntry_PropertyChanged;
            SimVars.Add(entry);
        }

        // After loading/applying config, ensure the required init-position simvars are present
        EnsureRequiredSimVarsPresent();
    }

    private void SimVarEntry_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (sender is SimVarEntry entry && e.PropertyName == nameof(SimVarEntry.Value))
        {
            SaveSimVarsToFile();
        }
    }

    private static async Task<object?> GetSimVarValueAsync(
        SimConnectClient client,
        string name,
        string unit,
        SimConnectDataType dataType,
        CancellationToken cancellationToken)
    {
        var simVars = client.SimVars;
        unit ??= string.Empty;

        return dataType switch
        {
            SimConnectDataType.Integer32 => await simVars.GetAsync<int>(name, unit, cancellationToken: cancellationToken).ConfigureAwait(false),
            SimConnectDataType.Integer64 => await simVars.GetAsync<long>(name, unit, cancellationToken: cancellationToken).ConfigureAwait(false),
            SimConnectDataType.FloatSingle => await simVars.GetAsync<float>(name, unit, cancellationToken: cancellationToken).ConfigureAwait(false),
            SimConnectDataType.FloatDouble => await simVars.GetAsync<double>(name, unit, cancellationToken: cancellationToken).ConfigureAwait(false),
            SimConnectDataType.String8 or
            SimConnectDataType.String32 or
            SimConnectDataType.String64 or
            SimConnectDataType.String128 or
            SimConnectDataType.String256 or
            SimConnectDataType.String260 or
            SimConnectDataType.StringV => await simVars.GetAsync<string>(name, unit, cancellationToken: cancellationToken).ConfigureAwait(false),
            SimConnectDataType.LatLonAlt => await simVars.GetAsync<SimConnectDataLatLonAlt>(name, unit, cancellationToken: cancellationToken).ConfigureAwait(false),
            SimConnectDataType.Xyz => await simVars.GetAsync<SimConnectDataXyz>(name, unit, cancellationToken: cancellationToken).ConfigureAwait(false),
            _ => throw new NotSupportedException($"SimVar '{name}' with data type '{dataType}' is not supported."),
        };
    }

    private static async Task SetSimVarValueAsync(
        SimConnectClient client,
        string name,
        string unit,
        SimConnectDataType dataType,
        object value,
        CancellationToken cancellationToken)
    {
        var simVars = client.SimVars;
        unit ??= string.Empty;

        switch (dataType)
        {
            case SimConnectDataType.Integer32:
                await simVars.SetAsync(name, unit, Convert.ToInt32(value, CultureInfo.InvariantCulture), cancellationToken: cancellationToken).ConfigureAwait(false);
                break;
            case SimConnectDataType.Integer64:
                await simVars.SetAsync(name, unit, Convert.ToInt64(value, CultureInfo.InvariantCulture), cancellationToken: cancellationToken).ConfigureAwait(false);
                break;
            case SimConnectDataType.FloatSingle:
                await simVars.SetAsync(name, unit, Convert.ToSingle(value, CultureInfo.InvariantCulture), cancellationToken: cancellationToken).ConfigureAwait(false);
                break;
            case SimConnectDataType.FloatDouble:
                await simVars.SetAsync(name, unit, Convert.ToDouble(value, CultureInfo.InvariantCulture), cancellationToken: cancellationToken).ConfigureAwait(false);
                break;
            case SimConnectDataType.String8:
            case SimConnectDataType.String32:
            case SimConnectDataType.String64:
            case SimConnectDataType.String128:
            case SimConnectDataType.String256:
            case SimConnectDataType.String260:
            case SimConnectDataType.StringV:
                await simVars.SetAsync(name, unit, Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty, cancellationToken: cancellationToken).ConfigureAwait(false);
                break;
            case SimConnectDataType.LatLonAlt:
                await simVars.SetAsync(name, unit, (SimConnectDataLatLonAlt)value, cancellationToken: cancellationToken).ConfigureAwait(false);
                break;
            case SimConnectDataType.Xyz:
                await simVars.SetAsync(name, unit, (SimConnectDataXyz)value, cancellationToken: cancellationToken).ConfigureAwait(false);
                break;
            case SimConnectDataType.InitPosition:
                await simVars.SetAsync(name, unit, (SimConnectDataInitPosition)value, cancellationToken: cancellationToken).ConfigureAwait(false);
                break;
            default:
                throw new NotSupportedException($"SimVar '{name}' with data type '{dataType}' is not supported for setting.");
        }
    }

    private static bool TryParseValue(string? text, SimConnectDataType dataType, out object? value, out string? error)
    {
        value = null;
        error = null;

        if (string.IsNullOrWhiteSpace(text))
        {
            error = "Value must not be empty.";
            return false;
        }

        var trimmed = text.Trim();

        switch (dataType)
        {
            case SimConnectDataType.Integer32:
                if (int.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out var intValue))
                {
                    value = intValue;
                    return true;
                }

                if (bool.TryParse(trimmed, out var boolValue))
                {
                    value = boolValue ? 1 : 0;
                    return true;
                }

                error = "Expected an integer or boolean value.";
                return false;

            case SimConnectDataType.Integer64:
                if (long.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out var longValue))
                {
                    value = longValue;
                    return true;
                }

                error = "Expected a 64-bit integer value.";
                return false;

            case SimConnectDataType.FloatSingle:
                if (double.TryParse(trimmed, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var singleDouble))
                {
                    value = (float)singleDouble;
                    return true;
                }

                error = "Expected a floating point value.";
                return false;

            case SimConnectDataType.FloatDouble:
                if (double.TryParse(trimmed, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var doubleValue))
                {
                    value = doubleValue;
                    return true;
                }

                error = "Expected a floating point value.";
                return false;

            case SimConnectDataType.String8:
            case SimConnectDataType.String32:
            case SimConnectDataType.String64:
            case SimConnectDataType.String128:
            case SimConnectDataType.String256:
            case SimConnectDataType.String260:
            case SimConnectDataType.StringV:
                value = trimmed;
                return true;

            case SimConnectDataType.LatLonAlt:
                {
                    var segments = trimmed.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
                    if (segments.Length != 3)
                    {
                        error = "Expected format 'latitude,longitude,altitude'.";
                        return false;
                    }

                    if (!double.TryParse(segments[0], NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var lat) ||
                        !double.TryParse(segments[1], NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var lon) ||
                        !double.TryParse(segments[2], NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var alt))
                    {
                        error = "Could not parse latitude, longitude, or altitude.";
                        return false;
                    }

                    value = new SimConnectDataLatLonAlt
                    {
                        Latitude = lat,
                        Longitude = lon,
                        Altitude = alt
                    };

                    return true;
                }

            case SimConnectDataType.Xyz:
                {
                    var segments = trimmed.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
                    if (segments.Length != 3)
                    {
                        error = "Expected format 'x,y,z'.";
                        return false;
                    }

                    if (!double.TryParse(segments[0], NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var x) ||
                        !double.TryParse(segments[1], NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var y) ||
                        !double.TryParse(segments[2], NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var z))
                    {
                        error = "Could not parse x, y, or z.";
                        return false;
                    }

                    value = new SimConnectDataXyz
                    {
                        X = x,
                        Y = y,
                        Z = z
                    };

                    return true;
                }

            default:
                error = $"Data type '{dataType}' is not supported.";
                return false;
        }
    }

    private static string FormatValue(object? value, SimConnectDataType dataType)
    {
        if (value is null)
        {
            return string.Empty;
        }

        return dataType switch
        {
            SimConnectDataType.LatLonAlt when value is SimConnectDataLatLonAlt latLonAlt => string.Join(',',
                latLonAlt.Latitude.ToString("G", CultureInfo.InvariantCulture),
                latLonAlt.Longitude.ToString("G", CultureInfo.InvariantCulture),
                latLonAlt.Altitude.ToString("G", CultureInfo.InvariantCulture)),
            SimConnectDataType.Xyz when value is SimConnectDataXyz xyz => string.Join(',',
                xyz.X.ToString("G", CultureInfo.InvariantCulture),
                xyz.Y.ToString("G", CultureInfo.InvariantCulture),
                xyz.Z.ToString("G", CultureInfo.InvariantCulture)),
            _ => value is IFormattable formattable
                ? formattable.ToString(null, CultureInfo.InvariantCulture)
                : Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty,
        };
    }

    public class SimVarEntry : INotifyPropertyChanged
    {
        private string _name = string.Empty;
        private string _unit = string.Empty;
        private string _value = string.Empty;
        private SimConnectDataType _dataType = SimConnectDataType.String256;
        private bool _isSettable;
        private ObservableCollection<EventMapping> _eventMappings = new();

        public string Name
        {
            get => _name;
            set => SetProperty(ref _name, value ?? string.Empty);
        }

        public string Unit
        {
            get => _unit;
            set => SetProperty(ref _unit, value ?? string.Empty);
        }

        public string Value
        {
            get => _value;
            set => SetProperty(ref _value, value ?? string.Empty);
        }

        public SimConnectDataType DataType
        {
            get => _dataType;
            set => SetProperty(ref _dataType, value);
        }

        public bool IsSettable
        {
            get => _isSettable;
            set => SetProperty(ref _isSettable, value);
        }

        // Optional: map specific SimVar values to input events
        public ObservableCollection<EventMapping> EventMappings
        {
            get => _eventMappings;
            set => SetProperty(ref _eventMappings, value ?? new ObservableCollection<EventMapping>());
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        private bool SetProperty<T>(ref T storage, T value, [CallerMemberName] string? propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(storage, value))
            {
                return false;
            }

            storage = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        private void OnPropertyChanged(string? propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private class SimVarConfigItem
    {
        private string _name = string.Empty;
        private string _unit = string.Empty;
        private string _value = string.Empty;

        public string Name
        {
            get => _name;
            set => _name = value ?? string.Empty;
        }

        public string Unit
        {
            get => _unit;
            set => _unit = value ?? string.Empty;
        }

        public SimConnectDataType DataType { get; set; } = SimConnectDataType.String256;

        public bool IsSettable { get; set; }

        public string Value
        {
            get => _value;
            set => _value = value ?? string.Empty;
        }

        // Persisted event mappings for this SimVar (optional)
        public List<EventMapping>? EventMappings { get; set; }
    }

    // Represents a mapping from a SimVar value to an Input Event and its parameter
    public class EventMapping : INotifyPropertyChanged
    {
        private string _matchValue = string.Empty;
        private string _eventName = string.Empty;
        private double _parameter;

        public string MatchValue
        {
            get => _matchValue;
            set
            {
                if (_matchValue != value)
                {
                    _matchValue = value ?? string.Empty;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(MatchValue)));
                }
            }
        }

        public string EventName
        {
            get => _eventName;
            set
            {
                if (_eventName != value)
                {
                    _eventName = value ?? string.Empty;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(EventName)));
                }
            }
        }

        public double Parameter
        {
            get => _parameter;
            set
            {
                if (!EqualityComparer<double>.Default.Equals(_parameter, value))
                {
                    _parameter = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Parameter)));
                }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
    }

    private class LegacySimVarConfig
    {
        public string Name { get; set; } = string.Empty;

        public string? Unit { get; set; }
    }

    private void SimVars_CollectionChanged(object? sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
    {
        if (_suppressCollectionChanged) return;

        // If items were replaced, copy new values into the existing entries
        if (e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Replace && e.NewItems != null)
        {
            foreach (SimVarEntry entry in e.NewItems)
            {
                var configItem = SimVars.FirstOrDefault(s => s.Name == entry.Name);
                if (configItem != null)
                {
                    configItem.Value = entry.Value;
                }
            }

            SaveSimVarsToFile();
            return;
        }

        // If items were removed, ensure required entries are not permanently removed
        if (e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Remove && e.OldItems != null)
        {
            var removedRequired = new List<string>();
            foreach (SimVarEntry old in e.OldItems)
            {
                if (RequiredInitPositionSimVars.Any(r => string.Equals(r, old.Name, StringComparison.OrdinalIgnoreCase)))
                {
                    removedRequired.Add(old.Name);
                }
            }

            if (removedRequired.Count > 0)
            {
                _suppressCollectionChanged = true;
                try
                {
                    foreach (var name in removedRequired)
                    {
                        var def = SimVarRegistry.Get(name);
                        var entry = new SimVarEntry
                        {
                            Name = def?.Name ?? name,
                            Unit = def?.Unit ?? string.Empty,
                            DataType = def?.DataType ?? SimConnectDataType.FloatDouble,
                            IsSettable = def?.IsSettable ?? true,
                            Value = string.Empty
                        };

                        entry.PropertyChanged += SimVarEntry_PropertyChanged;
                        SimVars.Add(entry);
                    }
                }
                finally
                {
                    _suppressCollectionChanged = false;
                }

                MessageBox.Show($"The following SimVars are required and cannot be removed:\n{string.Join('\n', removedRequired)}", "Required SimVars", MessageBoxButton.OK, MessageBoxImage.Information);
                SaveSimVarsToFile();
            }
        }
    }

    private void SaveSimVarsToFile()
    {
        try
        {
            var configItems = SimVars.Select(s => new SimVarConfigItem
            {
                Name = s.Name,
                Unit = s.Unit,
                DataType = s.DataType,
                IsSettable = s.IsSettable,
                Value = s.Value,
                EventMappings = s.EventMappings?.Select(m => new EventMapping
                {
                    MatchValue = m.MatchValue,
                    EventName = m.EventName,
                    Parameter = m.Parameter
                }).ToList()
            }).ToList();

            var json = JsonSerializer.Serialize(configItems, _serializerOptions);
            File.WriteAllText(ConfigPath, json);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to save SimVars to file:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // Open a window to edit event mappings for a specific row
    private void EventMapping_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement fe)
            return;

        if (fe.DataContext is not SimVarEntry rowItem)
            return;

        var initial = rowItem.EventMappings?.Select(m => new EventMapping
        {
            MatchValue = m.MatchValue,
            EventName = m.EventName,
            Parameter = m.Parameter
        }).ToList() ?? new List<EventMapping>();

        var dlg = new EventMappingWindow(rowItem.Name, initial)
        {
            Owner = this
        };

        var result = dlg.ShowDialog();
        if (result == true)
        {
            rowItem.EventMappings = new ObservableCollection<EventMapping>(dlg.EditedMappings);
            SaveSimVarsToFile();
        }
    }
}