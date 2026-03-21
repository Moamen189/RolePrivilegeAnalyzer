using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Xrm.Sdk;
using RolePrivilegeAnalyzer.Models;
using RolePrivilegeAnalyzer.Services;
using XrmToolBox.Extensibility;
using Label = System.Windows.Forms.Label;
namespace RolePrivilegeAnalyzer.UI
{
    /// <summary>
    /// Main XrmToolBox plugin control.
    /// Implements VirtualMode DataGridView for high-performance rendering of 20k+ users.
    /// </summary>
    public partial class RolePrivilegeAnalyzerControl : PluginControlBase
    {
        // ── Services ──
        private DataService _dataService;
        private RiskAnalysisService _riskService;
        private ExportService _exportService;

        // ── Data ──
        private List<UserRoleModel> _allUsers = new List<UserRoleModel>();
        private List<UserRoleModel> _filteredUsers = new List<UserRoleModel>();
        private AnalyticsSummary _analytics;

        // ── UI Components ──
        private DataGridView _grid;
        private TextBox _searchBox;
        private ComboBox _riskFilterCombo;
        private Panel _topPanel;
        private Panel _detailPanel;
        private Panel _analyticsPanel;
        private StatusStrip _statusStrip;
        private ToolStripStatusLabel _statusLabel;
        private ToolStripProgressBar _progressBar;
        private Timer _searchDebounceTimer;

        // ── Comparison State ──
        private UserRoleModel _compareUserA;
        private UserRoleModel _compareUserB;

        // ── Colors ──
        private static readonly Color ColorPrimary = Color.FromArgb(27, 58, 92);
        private static readonly Color ColorSecondary = Color.FromArgb(44, 95, 138);
        private static readonly Color ColorAccent = Color.FromArgb(52, 152, 219);
        private static readonly Color ColorBackground = Color.FromArgb(245, 247, 250);
        private static readonly Color ColorRowAlt = Color.FromArgb(236, 242, 248);
        private static readonly Color ColorHighRisk = Color.FromArgb(255, 215, 215);
        private static readonly Color ColorWarning = Color.FromArgb(255, 243, 205);
        private static readonly Color ColorNormal = Color.FromArgb(212, 237, 218);
        private static readonly Color ColorNoRole = Color.FromArgb(240, 240, 240);

        public RolePrivilegeAnalyzerControl()
        {
            InitializeUI();
        }

        #region UI Initialization
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);

        private const int EM_SETCUEBANNER = 0x1501;
        private void InitializeUI()
        {
            this.BackColor = ColorBackground;
            this.Dock = DockStyle.Fill;

            // ── Debounce Timer for Search ──
            _searchDebounceTimer = new Timer { Interval = 300 };
            _searchDebounceTimer.Tick += (s, e) =>
            {
                _searchDebounceTimer.Stop();
                ApplyFilters();
            };

            // ── Top Toolbar Panel ──
            _topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 56,
                BackColor = ColorPrimary,
                Padding = new Padding(10, 8, 10, 8)
            };

            var btnLoad = CreateToolButton("🔄 Load Data", ColorSecondary, async (s, e) => await LoadDataAsync());
            btnLoad.Location = new Point(10, 10);

            var btnExport = CreateToolButton("📊 Export Excel", Color.FromArgb(39, 174, 96), async (s, e) => await ExportAsync());
            btnExport.Location = new Point(135, 10);

            var btnAnalytics = CreateToolButton("📈 Analytics", Color.FromArgb(142, 68, 173), (s, e) => ToggleAnalyticsPanel());
            btnAnalytics.Location = new Point(280, 10);

            var btnCompare = CreateToolButton("⚖️ Compare", Color.FromArgb(230, 126, 34), (s, e) => ShowComparisonDialog());
            btnCompare.Location = new Point(405, 10);

            // Search Box
            _searchBox = new TextBox
            {
                Location = new Point(540, 12),
                Size = new Size(250, 30),
                Font = new Font("Segoe UI", 10F),
                //Text = "🔍  Search user or role...",
                BorderStyle = BorderStyle.FixedSingle
            };
            _searchBox.HandleCreated += (s, e) =>
            {
                SendMessage(_searchBox.Handle, EM_SETCUEBANNER, IntPtr.Zero, "Search user or role...");
            };
            _searchBox.TextChanged += (s, e) =>
            {
                _searchDebounceTimer.Stop();
                _searchDebounceTimer.Start();
            };
            // Risk filter dropdown
            _riskFilterCombo = new ComboBox
            {
                Location = new Point(800, 12),
                Size = new Size(165, 30),
                Font = new Font("Segoe UI", 9.5F),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _riskFilterCombo.Items.AddRange(new object[]
            {
                "All Statuses",
                "🔴 High Risk",
                "⚠️ Over Privileged",
                "✅ Normal",
                "❌ No Roles"
            });
            _riskFilterCombo.SelectedIndex = 0;
            _riskFilterCombo.SelectedIndexChanged += (s, e) => ApplyFilters();

            _topPanel.Controls.AddRange(new Control[] { btnLoad, btnExport, btnAnalytics, btnCompare, _searchBox, _riskFilterCombo });

            // ── Status Bar ──
            _statusStrip = new StatusStrip { BackColor = ColorPrimary };
            _statusLabel = new ToolStripStatusLabel("Ready")
            {
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9F)
            };
            _progressBar = new ToolStripProgressBar
            {
                Size = new Size(200, 16),
                Visible = false
            };
            _statusStrip.Items.AddRange(new ToolStripItem[] { _statusLabel, _progressBar });

            // ── Detail Panel (right side) ──
            _detailPanel = new Panel
            {
                Dock = DockStyle.Right,
                Width = 360,
                BackColor = Color.White,
                Visible = false,
                Padding = new Padding(12),
                BorderStyle = BorderStyle.None
            };
            // Add a left border line
            var detailBorder = new Panel
            {
                Dock = DockStyle.Left,
                Width = 2,
                BackColor = ColorAccent
            };
            _detailPanel.Controls.Add(detailBorder);

            // ── Analytics Panel (right side, toggled) ──
            _analyticsPanel = new Panel
            {
                Dock = DockStyle.Right,
                Width = 360,
                BackColor = Color.White,
                Visible = false,
                Padding = new Padding(12),
                AutoScroll = true
            };

            // ── Main Grid ──
            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                VirtualMode = true,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                GridColor = Color.FromArgb(230, 230, 230),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 42,
                RowTemplate = { Height = 36 },
                Font = new Font("Segoe UI", 9.5F),
                EnableHeadersVisualStyles = false
            };

            // Column definitions
            _grid.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn { Name = "colBU", HeaderText = "Business Unit", FillWeight = 15 },
                new DataGridViewTextBoxColumn { Name = "colUser", HeaderText = "User", FillWeight = 18 },
                new DataGridViewTextBoxColumn { Name = "colAD", HeaderText = "AD Name / Email", FillWeight = 20 },
                new DataGridViewTextBoxColumn { Name = "colRoles", HeaderText = "Roles", FillWeight = 22 },
                new DataGridViewTextBoxColumn { Name = "colRisk", HeaderText = "Privilege Risk Level", FillWeight = 12 },
                new DataGridViewTextBoxColumn { Name = "colStatus", HeaderText = "Risk Status", FillWeight = 13 }
            });

            // Header styling
            _grid.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = ColorPrimary,
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 10F),
                Alignment = DataGridViewContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0)
            };

            _grid.DefaultCellStyle = new DataGridViewCellStyle
            {
                Padding = new Padding(8, 4, 4, 4),
                SelectionBackColor = ColorAccent,
                SelectionForeColor = Color.White
            };

            // ── Wire Events ──
            _grid.CellValueNeeded += Grid_CellValueNeeded;
            _grid.CellFormatting += Grid_CellFormatting;
            _grid.CellDoubleClick += Grid_CellDoubleClick;
            _grid.SelectionChanged += Grid_SelectionChanged;
            _grid.ColumnHeaderMouseClick += Grid_ColumnHeaderMouseClick;
            _grid.CellToolTipTextNeeded += Grid_CellToolTipTextNeeded;

            // ── Add controls to form ──
            // Order matters for Dock layout
            this.Controls.Add(_grid);
            this.Controls.Add(_detailPanel);
            this.Controls.Add(_analyticsPanel);
            this.Controls.Add(_topPanel);
            this.Controls.Add(_statusStrip);
        }

        private Button CreateToolButton(string text, Color bgColor, EventHandler onClick)
        {
            var btn = new Button
            {
                Text = text,
                Size = new Size(120, 34),
                FlatStyle = FlatStyle.Flat,
                BackColor = bgColor,
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 9F),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.Click += onClick;
            return btn;
        }

        private Button CreateToolButton(string text, Color bgColor, Func<object, EventArgs, Task> asyncOnClick)
        {
            var btn = CreateToolButton(text, bgColor, (EventHandler)null);
            btn.Click += async (s, e) => await asyncOnClick(s, e);
            return btn;
        }

        #endregion

        #region Data Loading

        private async Task LoadDataAsync()
        {
            if (Service == null)
            {
                MessageBox.Show("Please connect to a Dataverse environment first.",
                    "Not Connected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                SetBusy(true);
                _progressBar.Visible = true;
                _progressBar.Value = 0;

                _dataService = new DataService(Service);
                _riskService = new RiskAnalysisService(_dataService);
                _exportService = new ExportService();

                // Determine org URL from connection
                string orgUrl = string.Empty;
                try
                {
                    orgUrl = ConnectionDetail?.WebApplicationUrl ?? string.Empty;
                }
                catch { /* Ignore if not available */ }

                _allUsers = await _dataService.LoadAllDataAsync(
                    orgUrl,
                    (message, progress) =>
                    {
                        this.Invoke((Action)(() =>
                        {
                            _statusLabel.Text = message;
                            _progressBar.Value = Math.Min(progress, 100);
                        }));
                    });

                // Run risk analysis
                _riskService.AnalyzeAll(_allUsers);

                // Generate analytics
                _analytics = _riskService.GenerateAnalytics(_allUsers);

                // Apply initial filter (show all)
                ApplyFilters();

                _statusLabel.Text = $"Loaded {_allUsers.Count:N0} users | {_analytics.TotalRoles:N0} roles";
                _progressBar.Visible = false;
            }
            catch (Exception ex)
            {
                _progressBar.Visible = false;
                _statusLabel.Text = "Error loading data";
                MessageBox.Show($"Failed to load data:\n\n{ex.Message}\n\n{ex.InnerException?.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private void SetBusy(bool busy)
        {
            _topPanel.Enabled = !busy;
            _grid.Enabled = !busy;
            Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
        }

        #endregion

        #region VirtualMode Grid Implementation

        private void Grid_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _filteredUsers.Count)
                return;

            var user = _filteredUsers[e.RowIndex];

            switch (e.ColumnIndex)
            {
                case 0: e.Value = user.BusinessUnit; break;
                case 1: e.Value = user.FullName; break;
                case 2: e.Value = user.DomainName; break;
                case 3: e.Value = user.RolesDisplay; break;
                case 4: e.Value = user.RiskLevel; break;
                case 5: e.Value = user.RiskStatus; break;
            }
        }

        private void Grid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _filteredUsers.Count)
                return;

            var user = _filteredUsers[e.RowIndex];

            // Alternating row colors
            if (e.RowIndex % 2 == 1)
            {
                e.CellStyle.BackColor = ColorRowAlt;
            }

            // Risk Status column color coding
            if (e.ColumnIndex == 5)
            {
                if (user.RiskStatus.Contains("High Risk"))
                {
                    e.CellStyle.BackColor = ColorHighRisk;
                    e.CellStyle.ForeColor = Color.DarkRed;
                    e.CellStyle.Font = new Font("Segoe UI Semibold", 9.5F);
                }
                else if (user.RiskStatus.Contains("Over Privileged"))
                {
                    e.CellStyle.BackColor = ColorWarning;
                    e.CellStyle.ForeColor = Color.FromArgb(133, 100, 4);
                    e.CellStyle.Font = new Font("Segoe UI Semibold", 9.5F);
                }
                else if (user.RiskStatus.Contains("Normal"))
                {
                    e.CellStyle.BackColor = ColorNormal;
                    e.CellStyle.ForeColor = Color.FromArgb(21, 87, 36);
                }
                else if (user.RiskStatus.Contains("No Roles"))
                {
                    e.CellStyle.BackColor = ColorNoRole;
                    e.CellStyle.ForeColor = Color.Gray;
                }
            }

            // Risk Level column color coding
            if (e.ColumnIndex == 4)
            {
                if (user.RiskLevel.Contains("Global"))
                    e.CellStyle.ForeColor = Color.Red;
                else if (user.RiskLevel.Contains("Deep"))
                    e.CellStyle.ForeColor = Color.OrangeRed;
                else if (user.RiskLevel.Contains("Local"))
                    e.CellStyle.ForeColor = Color.FromArgb(180, 140, 0);
                else if (user.RiskLevel.Contains("Basic"))
                    e.CellStyle.ForeColor = Color.Green;

                e.CellStyle.Font = new Font("Segoe UI Semibold", 9.5F);
            }

            // User name column — make it look clickable
            if (e.ColumnIndex == 1)
            {
                e.CellStyle.ForeColor = ColorAccent;
                e.CellStyle.Font = new Font("Segoe UI", 9.5F, FontStyle.Underline);
            }
        }

        private void Grid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _filteredUsers.Count)
                return;

            var user = _filteredUsers[e.RowIndex];

            // Double-click on user column → open in CRM
            if (e.ColumnIndex == 1 && !string.IsNullOrEmpty(user.CrmUrl))
            {
                try
                {
                    Process.Start(new ProcessStartInfo(user.CrmUrl) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Could not open CRM record:\n{ex.Message}",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void Grid_SelectionChanged(object sender, EventArgs e)
        {
            if (_grid.CurrentRow == null || _grid.CurrentRow.Index < 0
                || _grid.CurrentRow.Index >= _filteredUsers.Count)
            {
                _detailPanel.Visible = false;
                return;
            }

            var user = _filteredUsers[_grid.CurrentRow.Index];
            ShowUserDetail(user);
        }

        private void Grid_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _filteredUsers.Count)
                return;

            if (e.ColumnIndex == 1)
            {
                e.ToolTipText = "Double-click to open in CRM";
            }
            else if (e.ColumnIndex == 3)
            {
                var user = _filteredUsers[e.RowIndex];
                e.ToolTipText = string.Join("\n", user.Roles);
            }
        }

        #endregion

        #region Search & Filtering

        private void ApplyFilters()
        {
            string searchTerm = _searchBox.Text?.Trim() ?? string.Empty;
            string riskFilter = _riskFilterCombo.SelectedItem?.ToString() ?? "All Statuses";

            _filteredUsers = _allUsers.Where(u =>
            {
                // Text search filter
                if (!string.IsNullOrEmpty(searchTerm))
                {
                    bool matchesText =
                        u.FullName.IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        u.DomainName.IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        u.BusinessUnit.IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        u.Roles.Any(r => r.IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0);

                    if (!matchesText) return false;
                }

                // Risk status filter
                if (riskFilter != "All Statuses")
                {
                    if (!u.RiskStatus.Contains(riskFilter.Substring(riskFilter.IndexOf(' ') + 1)))
                        return false;
                }

                return true;
            }).ToList();

            // Update grid row count for VirtualMode
            _grid.RowCount = 0; // Reset first to avoid index issues
            _grid.RowCount = _filteredUsers.Count;
            _grid.Invalidate();

            _statusLabel.Text = $"Showing {_filteredUsers.Count:N0} of {_allUsers.Count:N0} users";
        }

        #endregion

        #region Column Sorting

        private string _currentSortColumn = null;
        private bool _sortAscending = true;

        private void Grid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (_filteredUsers.Count == 0) return;

            string colName = _grid.Columns[e.ColumnIndex].Name;

            if (_currentSortColumn == colName)
                _sortAscending = !_sortAscending;
            else
            {
                _currentSortColumn = colName;
                _sortAscending = true;
            }

            Func<UserRoleModel, object> keySelector;
            switch (e.ColumnIndex)
            {
                case 0: keySelector = u => u.BusinessUnit; break;
                case 1: keySelector = u => u.FullName; break;
                case 2: keySelector = u => u.DomainName; break;
                case 3: keySelector = u => u.Roles.Count; break;
                case 4: keySelector = u => u.RiskLevel; break;
                case 5: keySelector = u => u.RiskStatus; break;
                default: return;
            }

            _filteredUsers = _sortAscending
                ? _filteredUsers.OrderBy(keySelector).ToList()
                : _filteredUsers.OrderByDescending(keySelector).ToList();

            _grid.Invalidate();
        }

        #endregion

        #region Detail Panel (Privilege Drill-down)

        private void ShowUserDetail(UserRoleModel user)
        {
            _detailPanel.Controls.Clear();

            var borderPanel = new Panel
            {
                Dock = DockStyle.Left,
                Width = 2,
                BackColor = ColorAccent
            };
            _detailPanel.Controls.Add(borderPanel);

            var scrollPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(14, 10, 10, 10)
            };

            int y = 10;

            // Close button
            var btnClose = new Label
            {
                Text = "✕",
                Font = new Font("Segoe UI", 14F),
                ForeColor = Color.Gray,
                AutoSize = true,
                Location = new Point(310, 2),
                Cursor = Cursors.Hand
            };
            btnClose.Click += (s, e) => _detailPanel.Visible = false;
            scrollPanel.Controls.Add(btnClose);

            // User name header
            var lblName = new Label
            {
                Text = user.FullName,
                Font = new Font("Segoe UI Semibold", 14F),
                ForeColor = ColorPrimary,
                AutoSize = true,
                Location = new Point(5, y),
                MaximumSize = new Size(300, 0)
            };
            scrollPanel.Controls.Add(lblName);
            y += lblName.Height + 4;

            // Domain
            AddDetailLabel(scrollPanel, $"📧 {user.DomainName}", Color.Gray, ref y);
            AddDetailLabel(scrollPanel, $"🏢 {user.BusinessUnit}", Color.Gray, ref y);
            AddDetailLabel(scrollPanel, $"🛡️ {user.RiskStatus}  |  {user.RiskLevel}", ColorPrimary, ref y);
            y += 8;

            // Copy roles button
            var btnCopy = new Button
            {
                Text = "📋 Copy Roles",
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAccent,
                ForeColor = Color.White,
                Size = new Size(110, 28),
                Font = new Font("Segoe UI", 8.5F),
                Location = new Point(5, y),
                Cursor = Cursors.Hand
            };
            btnCopy.FlatAppearance.BorderSize = 0;
            btnCopy.Click += (s, e) =>
            {
                var text = $"User: {user.FullName}\nRoles:\n{string.Join("\n", user.Roles.Select((r, i) => $"  {i + 1}. {r}"))}";
                Clipboard.SetText(text);
                _statusLabel.Text = "Roles copied to clipboard!";
            };
            scrollPanel.Controls.Add(btnCopy);

            // Set as Compare A / B buttons
            var btnSetA = new Button
            {
                Text = "Set as A",
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(230, 126, 34),
                ForeColor = Color.White,
                Size = new Size(75, 28),
                Font = new Font("Segoe UI", 8.5F),
                Location = new Point(125, y),
                Cursor = Cursors.Hand
            };
            btnSetA.FlatAppearance.BorderSize = 0;
            btnSetA.Click += (s, e) =>
            {
                _compareUserA = user;
                _statusLabel.Text = $"Compare User A: {user.FullName}";
            };
            scrollPanel.Controls.Add(btnSetA);

            var btnSetB = new Button
            {
                Text = "Set as B",
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(142, 68, 173),
                ForeColor = Color.White,
                Size = new Size(75, 28),
                Font = new Font("Segoe UI", 8.5F),
                Location = new Point(208, y),
                Cursor = Cursors.Hand
            };
            btnSetB.FlatAppearance.BorderSize = 0;
            btnSetB.Click += (s, e) =>
            {
                _compareUserB = user;
                _statusLabel.Text = $"Compare User B: {user.FullName}";
            };
            scrollPanel.Controls.Add(btnSetB);

            y += 40;

            // ── Privilege Drill-down ──
            AddDetailLabel(scrollPanel, "── Privileges by Role ──", ColorPrimary, ref y, bold: true);
            y += 4;

            if (_dataService != null)
            {
                var privsByRole = _dataService.GetUserPrivilegesByRole(user.UserId);

                if (privsByRole.Count == 0)
                {
                    AddDetailLabel(scrollPanel, "No privileges found.", Color.Gray, ref y);
                }
                else
                {
                    foreach (var kvp in privsByRole.OrderBy(k => k.Key))
                    {
                        AddDetailLabel(scrollPanel, $"🔑 {kvp.Key}", ColorSecondary, ref y, bold: true);

                        // Show top 20 privileges per role (to avoid massive lists)
                        var topPrivs = kvp.Value
                            .OrderByDescending(p => p.AccessDepth)
                            .ThenBy(p => p.Name)
                            .Take(20);

                        foreach (var priv in topPrivs)
                        {
                            string icon = GetAccessLevelIcon(priv.AccessLevel);
                            AddDetailLabel(scrollPanel, $"  {icon} {priv.Name} → {priv.AccessLevel}", Color.DimGray, ref y, fontSize: 8.5F);
                        }

                        if (kvp.Value.Count > 20)
                        {
                            AddDetailLabel(scrollPanel, $"  ... and {kvp.Value.Count - 20} more", Color.LightGray, ref y, fontSize: 8F);
                        }

                        y += 6;
                    }
                }
            }

            _detailPanel.Controls.Add(scrollPanel);
            _detailPanel.Visible = true;
        }

        private void AddDetailLabel(Panel parent, string text, Color foreColor, ref int y,
            bool bold = false, float fontSize = 9F)
        {
            var lbl = new Label
            {
                Text = text,
                Font = new Font("Segoe UI", fontSize, bold ? FontStyle.Bold : FontStyle.Regular),
                ForeColor = foreColor,
                AutoSize = true,
                MaximumSize = new Size(320, 0),
                Location = new Point(5, y)
            };
            parent.Controls.Add(lbl);
            y += lbl.Height + 2;
        }

        private string GetAccessLevelIcon(string level)
        {
            switch (level)
            {
                case "Global": return "🔴";
                case "Deep": return "🟠";
                case "Local": return "🟡";
                case "Basic": return "🟢";
                default: return "⚪";
            }
        }

        #endregion

        #region Analytics Panel

        private void ToggleAnalyticsPanel()
        {
            if (_analytics == null)
            {
                MessageBox.Show("Please load data first.", "No Data",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (_analyticsPanel.Visible)
            {
                _analyticsPanel.Visible = false;
                return;
            }

            _analyticsPanel.Controls.Clear();

            var scrollPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(14, 10, 10, 10)
            };

            int y = 10;

            // Close button
            var btnClose = new Label
            {
                Text = "✕",
                Font = new Font("Segoe UI", 14F),
                ForeColor = Color.Gray,
                AutoSize = true,
                Location = new Point(320, 2),
                Cursor = Cursors.Hand
            };
            btnClose.Click += (s, e) => _analyticsPanel.Visible = false;
            scrollPanel.Controls.Add(btnClose);

            AddDetailLabel(scrollPanel, "📊 Analytics Dashboard", ColorPrimary, ref y, bold: true, fontSize: 13F);
            y += 10;

            // Summary cards
            AddStatCard(scrollPanel, "Total Users", _analytics.TotalUsers.ToString("N0"), ColorSecondary, ref y);
            AddStatCard(scrollPanel, "Total Roles", _analytics.TotalRoles.ToString("N0"), ColorAccent, ref y);
            AddStatCard(scrollPanel, "🔴 High Risk", _analytics.HighRiskUsers.ToString("N0"), Color.Red, ref y);
            AddStatCard(scrollPanel, "⚠️ Over Privileged", _analytics.OverPrivilegedUsers.ToString("N0"), Color.FromArgb(230, 126, 34), ref y);
            AddStatCard(scrollPanel, "❌ No Roles", _analytics.NoRoleUsers.ToString("N0"), Color.Gray, ref y);
            AddStatCard(scrollPanel, "✅ Normal", _analytics.NormalUsers.ToString("N0"), Color.FromArgb(39, 174, 96), ref y);

            y += 16;
            AddDetailLabel(scrollPanel, "── Most Assigned Roles ──", ColorPrimary, ref y, bold: true);
            y += 6;

            foreach (var role in _analytics.MostAssignedRoles)
            {
                // Simple bar indicator
                int barWidth = Math.Max(10, (int)(300.0 * role.Count / Math.Max(1, _analytics.TotalUsers)));
                var barPanel = new Panel
                {
                    Location = new Point(5, y),
                    Size = new Size(barWidth, 20),
                    BackColor = ColorAccent
                };
                scrollPanel.Controls.Add(barPanel);

                var lblRole = new Label
                {
                    Text = $"{role.RoleName} ({role.Count})",
                    Font = new Font("Segoe UI", 8.5F),
                    ForeColor = Color.DimGray,
                    AutoSize = true,
                    Location = new Point(5, y + 22)
                };
                scrollPanel.Controls.Add(lblRole);

                y += 46;
            }

            _analyticsPanel.Controls.Add(scrollPanel);
            _analyticsPanel.Visible = true;
            _detailPanel.Visible = false; // Hide detail when analytics is shown
        }

        private void AddStatCard(Panel parent, string title, string value, Color accentColor, ref int y)
        {
            var card = new Panel
            {
                Location = new Point(5, y),
                Size = new Size(320, 50),
                BackColor = Color.FromArgb(248, 249, 252),
                BorderStyle = BorderStyle.None
            };

            var accent = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(4, 50),
                BackColor = accentColor
            };
            card.Controls.Add(accent);

            var lblTitle = new Label
            {
                Text = title,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.Gray,
                Location = new Point(14, 6),
                AutoSize = true
            };
            card.Controls.Add(lblTitle);

            var lblValue = new Label
            {
                Text = value,
                Font = new Font("Segoe UI Semibold", 14F),
                ForeColor = accentColor,
                Location = new Point(14, 24),
                AutoSize = true
            };
            card.Controls.Add(lblValue);

            parent.Controls.Add(card);
            y += 58;
        }

        #endregion

        #region Role Comparison

        private void ShowComparisonDialog()
        {
            if (_compareUserA == null || _compareUserB == null)
            {
                MessageBox.Show(
                    "Please select two users for comparison.\n\n" +
                    "Click a user row, then use 'Set as A' and 'Set as B' buttons in the detail panel.",
                    "Select Users", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (_riskService == null)
            {
                MessageBox.Show("Please load data first.", "No Data",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var result = _riskService.CompareUsers(_compareUserA, _compareUserB);

            // Build comparison form
            var form = new Form
            {
                Text = $"Role Comparison: {_compareUserA.FullName} vs {_compareUserB.FullName}",
                Size = new Size(800, 600),
                StartPosition = FormStartPosition.CenterParent,
                Font = new Font("Segoe UI", 9.5F),
                BackColor = ColorBackground
            };

            var rtb = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                Font = new Font("Consolas", 10F),
                BackColor = Color.White,
                Padding = new Padding(12)
            };

            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"═══════════════════════════════════════════════════");
            sb.AppendLine($"  ROLE COMPARISON");
            sb.AppendLine($"═══════════════════════════════════════════════════");
            sb.AppendLine();
            sb.AppendLine($"  User A: {result.UserA.FullName}  ({result.UserA.RiskStatus})");
            sb.AppendLine($"  User B: {result.UserB.FullName}  ({result.UserB.RiskStatus})");
            sb.AppendLine();
            sb.AppendLine($"── Common Roles ({result.CommonRoles.Count}) ──");
            foreach (var role in result.CommonRoles)
                sb.AppendLine($"   ✔ {role}");
            sb.AppendLine();

            sb.AppendLine($"── Only in {result.UserA.FullName} ({result.OnlyInA.Count}) ──");
            foreach (var role in result.OnlyInA)
                sb.AppendLine($"   🔹 {role}");
            sb.AppendLine();

            sb.AppendLine($"── Only in {result.UserB.FullName} ({result.OnlyInB.Count}) ──");
            foreach (var role in result.OnlyInB)
                sb.AppendLine($"   🔸 {role}");
            sb.AppendLine();

            if (result.PrivilegeDifferences.Count > 0)
            {
                sb.AppendLine($"── Privilege Differences (top 50) ──");
                sb.AppendLine($"  {"Privilege",-40} {"User A",-12} {"User B",-12}");
                sb.AppendLine($"  {new string('─', 64)}");
                foreach (var diff in result.PrivilegeDifferences.Take(50))
                {
                    sb.AppendLine($"  {diff.PrivilegeName,-40} {diff.UserALevel,-12} {diff.UserBLevel,-12}");
                }

                if (result.PrivilegeDifferences.Count > 50)
                {
                    sb.AppendLine($"\n  ... and {result.PrivilegeDifferences.Count - 50} more differences");
                }
            }

            rtb.Text = sb.ToString();
            form.Controls.Add(rtb);
            form.ShowDialog(this);
        }

        #endregion

        #region Export

        private async Task ExportAsync()
        {
            if (_allUsers.Count == 0)
            {
                MessageBox.Show("No data to export. Please load data first.",
                    "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var sfd = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"RolePrivilegeAnalysis_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
                Title = "Export to Excel"
            })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        SetBusy(true);
                        _statusLabel.Text = "Exporting...";

                        await _exportService.ExportToExcelAsync(_filteredUsers, _analytics, sfd.FileName);

                        _statusLabel.Text = $"Exported {_filteredUsers.Count:N0} users to Excel";
                        MessageBox.Show($"Export complete!\n\n{sfd.FileName}",
                            "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Export failed:\n{ex.Message}",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        SetBusy(false);
                    }
                }
            }
        }

        #endregion
    }
}
