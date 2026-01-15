using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows.Forms;

namespace TaskManager;

public sealed class TaskItem : INotifyPropertyChanged
{
    private string _title = string.Empty;
    private string _memo = string.Empty;
    private DateTime _dueDate = DateTime.Today;
    private int _progress;
    private bool _isCompleted;
    private long _workSeconds;
    private long _dailyWorkSeconds;
    private DateTime _dailyWorkDate = DateTime.Today;
    private bool _isTracking;
    private DateTime? _trackingStartUtc;

    public string Title
    {
        get => _title;
        set => SetField(ref _title, value?.Trim() ?? string.Empty, nameof(Title));
    }

    public DateTime DueDate
    {
        get => _dueDate;
        set => SetField(ref _dueDate, value.Date, nameof(DueDate));
    }

    public string Memo
    {
        get => _memo;
        set => SetField(ref _memo, value?.Trim() ?? string.Empty, nameof(Memo));
    }

    public int Progress
    {
        get => _progress;
        set
        {
            var clamped = Math.Clamp(value, 0, 100);
            if (!SetField(ref _progress, clamped, nameof(Progress)))
            {
                return;
            }

            if (_progress >= 100 && !_isCompleted)
            {
                _isCompleted = true;
                OnPropertyChanged(nameof(IsCompleted));
                StopTracking();
            }
            else if (_progress < 100 && _isCompleted)
            {
                _isCompleted = false;
                OnPropertyChanged(nameof(IsCompleted));
            }
        }
    }

    public bool IsCompleted
    {
        get => _isCompleted;
        set
        {
            if (!SetField(ref _isCompleted, value, nameof(IsCompleted)))
            {
                return;
            }

            if (_isCompleted && _progress < 100)
            {
                _progress = 100;
                OnPropertyChanged(nameof(Progress));
            }

            if (_isCompleted)
            {
                StopTracking();
            }
        }
    }

    public long WorkSeconds
    {
        get => _workSeconds;
        set
        {
            var clamped = Math.Max(0, value);
            if (!SetField(ref _workSeconds, clamped, nameof(WorkSeconds)))
            {
                return;
            }

            OnPropertyChanged(nameof(WorkDisplay));
        }
    }

    public long DailyWorkSeconds
    {
        get => _dailyWorkSeconds;
        set
        {
            var clamped = Math.Max(0, value);
            if (!SetField(ref _dailyWorkSeconds, clamped, nameof(DailyWorkSeconds)))
            {
                return;
            }

            OnPropertyChanged(nameof(WorkDisplay));
        }
    }

    public DateTime DailyWorkDate
    {
        get => _dailyWorkDate;
        set
        {
            var date = value.Date;
            if (!SetField(ref _dailyWorkDate, date, nameof(DailyWorkDate)))
            {
                return;
            }

            OnPropertyChanged(nameof(WorkDisplay));
        }
    }

    [JsonIgnore]
    public bool IsTracking => _isTracking;

    [JsonIgnore]
    public string WorkDisplay => $"{FormatDuration(GetDailySeconds())} ({FormatDuration(GetTotalSeconds())})";

    public void StartTracking()
    {
        if (_isTracking)
        {
            return;
        }

        EnsureDailyDate();
        _isTracking = true;
        _trackingStartUtc = DateTime.UtcNow;
        OnPropertyChanged(nameof(IsTracking));
        OnPropertyChanged(nameof(WorkDisplay));
    }

    public void StopTracking()
    {
        if (!_isTracking)
        {
            return;
        }

        EnsureDailyDate();
        _workSeconds = GetTotalSeconds();
        _dailyWorkSeconds = GetDailySeconds();
        _trackingStartUtc = null;
        _isTracking = false;
        OnPropertyChanged(nameof(WorkSeconds));
        OnPropertyChanged(nameof(DailyWorkSeconds));
        OnPropertyChanged(nameof(IsTracking));
        OnPropertyChanged(nameof(WorkDisplay));
    }

    public void Normalize()
    {
        _title = string.IsNullOrWhiteSpace(_title) ? "無題" : _title.Trim();
        _memo = string.IsNullOrWhiteSpace(_memo) ? string.Empty : _memo.Trim();
        _dueDate = _dueDate.Date;
        _progress = Math.Clamp(_progress, 0, 100);
        _workSeconds = Math.Max(0, _workSeconds);
        _dailyWorkSeconds = Math.Max(0, _dailyWorkSeconds);
        _dailyWorkDate = _dailyWorkDate == default ? DateTime.Today : _dailyWorkDate.Date;
        if (_dailyWorkDate != DateTime.Today)
        {
            _dailyWorkDate = DateTime.Today;
            _dailyWorkSeconds = 0;
        }
        _isTracking = false;
        _trackingStartUtc = null;

        if (_progress >= 100 && !_isCompleted)
        {
            _isCompleted = true;
        }
        else if (_progress < 100 && _isCompleted)
        {
            _isCompleted = false;
        }
    }

    private long GetTotalSeconds()
    {
        if (!_isTracking || _trackingStartUtc is null)
        {
            return _workSeconds;
        }

        var elapsed = (long)Math.Max(0, (DateTime.UtcNow - _trackingStartUtc.Value).TotalSeconds);
        return _workSeconds + elapsed;
    }

    private long GetDailySeconds()
    {
        var today = DateTime.Today;
        var baseDailySeconds = _dailyWorkDate == today ? _dailyWorkSeconds : 0;
        if (!_isTracking || _trackingStartUtc is null)
        {
            return baseDailySeconds;
        }

        var nowLocal = DateTime.Now;
        var startLocal = _trackingStartUtc.Value.ToLocalTime();
        var dayStart = today;
        var trackingStart = startLocal < dayStart ? dayStart : startLocal;
        var elapsedToday = (long)Math.Max(0, (nowLocal - trackingStart).TotalSeconds);
        return baseDailySeconds + elapsedToday;
    }

    private void EnsureDailyDate()
    {
        var today = DateTime.Today;
        if (_dailyWorkDate == today)
        {
            return;
        }

        _dailyWorkDate = today;
        _dailyWorkSeconds = 0;
        OnPropertyChanged(nameof(DailyWorkDate));
        OnPropertyChanged(nameof(DailyWorkSeconds));
        OnPropertyChanged(nameof(WorkDisplay));
    }

    private static string FormatDuration(long totalSeconds)
    {
        if (totalSeconds < 0)
        {
            totalSeconds = 0;
        }

        var time = TimeSpan.FromSeconds(totalSeconds);
        if (time.TotalHours >= 1)
        {
            return $"{(int)time.TotalHours:D2}:{time.Minutes:D2}:{time.Seconds:D2}";
        }

        return $"{time.Minutes:D2}:{time.Seconds:D2}";
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private bool SetField<T>(ref T field, T value, string propertyName)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    private void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

public sealed class NoteItem : INotifyPropertyChanged
{
    private string _title = string.Empty;
    private string _content = string.Empty;
    private string _folder = string.Empty;

    public string Title
    {
        get => _title;
        set => SetField(ref _title, value?.Trim() ?? string.Empty, nameof(Title));
    }

    public string Content
    {
        get => _content;
        set => SetField(ref _content, value ?? string.Empty, nameof(Content));
    }

    public string Folder
    {
        get => _folder;
        set => SetField(ref _folder, NormalizeFolderPath(value), nameof(Folder));
    }

    internal static string NormalizeFolderPath(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var normalized = value.Trim().Replace('\\', '/').Trim('/');
        while (normalized.Contains("//", StringComparison.Ordinal))
        {
            normalized = normalized.Replace("//", "/", StringComparison.Ordinal);
        }

        return normalized;
    }

    public void Normalize()
    {
        _title = string.IsNullOrWhiteSpace(_title) ? "無題" : _title.Trim();
        _content = _content ?? string.Empty;
        _folder = NormalizeFolderPath(_folder);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private bool SetField<T>(ref T field, T value, string propertyName)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        return true;
    }
}

public sealed class MainForm : Form
{
    private const string AppName = "枯乃葉のタスク管理";
    private const int ActionMenuSwitchWidth = 760;
    private const int NotesListMinWidth = 200;
    private const int NotesEditorMinWidth = 260;
    private const float NoteZoomStep = 0.1f;
    private const float NoteZoomMin = 0.6f;
    private const float NoteZoomMax = 2.0f;
    private static readonly Color BackgroundColor = Color.FromArgb(246, 244, 241);
    private static readonly Color CardColor = Color.FromArgb(255, 255, 255);
    private static readonly Color BorderColor = Color.FromArgb(233, 227, 222);
    private static readonly Color HeaderBackColor = Color.FromArgb(252, 247, 244);
    private static readonly Color AccentColor = Color.FromArgb(255, 153, 181);
    private static readonly Color AccentDarkColor = Color.FromArgb(255, 126, 163);
    private static readonly Color TextColor = Color.FromArgb(45, 45, 54);
    private static readonly Color MutedTextColor = Color.FromArgb(120, 120, 132);
    private static readonly Color SelectionColor = Color.FromArgb(255, 226, 236);
    private static readonly Color ProgressTrackColor = Color.FromArgb(244, 238, 234);
    private static readonly Color SuccessColor = Color.FromArgb(145, 214, 180);
    private readonly Icon _appIcon;
    private readonly bool _ownsIcon;
    private readonly Font _baseFont;
    private readonly Font _headerFont;
    private readonly Font _sectionFont;
    private readonly Font _hintFont;
    private readonly Font _buttonFont;
    private readonly BindingList<TaskItem> _activeTasks;
    private readonly BindingList<TaskItem> _completedTasks;
    private readonly BindingSource _activeBindingSource;
    private readonly BindingSource _completedBindingSource;
    private readonly BindingList<NoteItem> _notes;
    private readonly TextBox _titleInput;
    private readonly TextBox _memoInput;
    private readonly DateTimePicker _duePicker;
    private readonly NumericUpDown _progressInput;
    private readonly Button _addButton;
    private ComboBox _noteFolderInput = null!;
    private TextBox _noteTitleInput = null!;
    private RichTextBox _noteBodyInput = null!;
    private Button _noteSaveButton = null!;
    private Button _noteNewButton = null!;
    private Button _noteDeleteButton = null!;
    private TreeView _notesTreeView = null!;
    private Button _noteBoldButton = null!;
    private Button _noteUnderlineButton = null!;
    private Button _noteStrikeButton = null!;
    private Button _noteZoomOutButton = null!;
    private Button _noteZoomInButton = null!;
    private Button _noteZoomResetButton = null!;
    private Label _noteZoomLabel = null!;
    private Label _noteToastLabel = null!;
    private DataGridView _activeGrid = null!;
    private DataGridView _completedGrid = null!;
    private DataGridViewTextBoxColumn _activeProgressColumn = null!;
    private DataGridViewTextBoxColumn _activeDueColumn = null!;
    private DataGridViewTextBoxColumn _activeWorkTimeColumn = null!;
    private DataGridViewButtonColumn _activeTrackButtonColumn = null!;
    private DataGridViewButtonColumn _activeMenuColumn = null!;
    private DataGridViewButtonColumn _activeCompleteColumn = null!;
    private DataGridViewButtonColumn _activeDeleteColumn = null!;
    private DataGridViewTextBoxColumn _completedProgressColumn = null!;
    private DataGridViewTextBoxColumn _completedDueColumn = null!;
    private DataGridViewTextBoxColumn _completedWorkTimeColumn = null!;
    private DataGridViewButtonColumn _completedMenuColumn = null!;
    private DataGridViewButtonColumn _completedRestoreColumn = null!;
    private DataGridViewButtonColumn _completedDeleteColumn = null!;
    private TabPage _activeTabPage = null!;
    private TabPage _completedTabPage = null!;
    private TabControl _mainTabControl = null!;
    private TabPage _tasksMainTabPage = null!;
    private TabPage _notesMainTabPage = null!;
    private TabPage? _lastMainTabPage;
    private ProgressBar _overallProgress = null!;
    private Label _summaryLabel = null!;
    private readonly NotifyIcon _trayIcon;
    private readonly ContextMenuStrip _trayMenu;
    private readonly ContextMenuStrip _rowMenu;
    private readonly ToolStripMenuItem _rowMenuTimerItem;
    private readonly ToolStripSeparator _rowMenuTimerSeparator;
    private readonly ToolStripMenuItem _rowMenuToggleItem;
    private readonly ToolStripMenuItem _rowMenuDeleteItem;
    private readonly System.Windows.Forms.Timer _trackingTimer;
    private readonly System.Windows.Forms.Timer _noteToastTimer;
    private readonly string _dataPath;
    private readonly string _notesPath;
    private readonly JsonSerializerOptions _jsonOptions;
    private bool _allowExit;
    private bool _suppressSave;
    private bool _suppressNoteSave;
    private bool _suppressNoteTreeSelection;
    private bool _notesSplitterInitialized;
    private bool _trayTipShown;
    private bool _suppressMove;
    private NoteItem? _selectedNote;
    private sealed class RowMenuContext
    {
        public RowMenuContext(TaskItem task, bool isCompleted)
        {
            Task = task;
            IsCompleted = isCompleted;
        }

        public TaskItem Task { get; }
        public bool IsCompleted { get; }
    }

    public MainForm()
    {
        Text = AppName;
        _baseFont = CreateUiFont("Zen Maru Gothic", 10f, FontStyle.Regular);
        _headerFont = CreateUiFont("Zen Maru Gothic", 20f, FontStyle.Bold);
        _sectionFont = CreateUiFont("Zen Maru Gothic", 12f, FontStyle.Bold);
        _hintFont = CreateUiFont("Zen Maru Gothic", 9f, FontStyle.Regular);
        _buttonFont = CreateUiFont("Zen Maru Gothic", 10f, FontStyle.Bold);
        Font = _baseFont;
        ForeColor = TextColor;
        (_appIcon, _ownsIcon) = LoadAppIcon();
        MinimumSize = new Size(520, 650);
        Size = new Size(560, 720);
        StartPosition = FormStartPosition.CenterScreen;
        Icon = _appIcon;
        BackColor = BackgroundColor;
        DoubleBuffered = true;

        _dataPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "TaskManager",
            "tasks.json");
        _notesPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "TaskManager",
            "notes.json");

        _jsonOptions = new JsonSerializerOptions { WriteIndented = true };

        _activeTasks = new BindingList<TaskItem>();
        _completedTasks = new BindingList<TaskItem>();
        _activeTasks.ListChanged += Tasks_ListChanged;
        _completedTasks.ListChanged += Tasks_ListChanged;

        _activeBindingSource = new BindingSource { DataSource = _activeTasks };
        _completedBindingSource = new BindingSource { DataSource = _completedTasks };
        _notes = new BindingList<NoteItem>();
        _notes.ListChanged += Notes_ListChanged;

        _titleInput = new TextBox
        {
            Dock = DockStyle.Fill,
            PlaceholderText = "タイトルを入力",
            BackColor = Color.White,
            ForeColor = TextColor
        };
        _memoInput = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            PlaceholderText = "メモ（任意）",
            BackColor = Color.White,
            ForeColor = TextColor
        };
        _memoInput.KeyDown += (_, e) =>
        {
            if (e.Control && e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                e.Handled = true;
                AddTask();
            }
        };
        _memoInput.Enter += (_, _) => AcceptButton = null;
        _memoInput.Leave += (_, _) => RestoreAcceptButtonForCurrentTab();
        _duePicker = new DateTimePicker
        {
            Format = DateTimePickerFormat.Short,
            Width = 160,
            Value = DateTime.Today
        };
        _progressInput = new NumericUpDown
        {
            Minimum = 0,
            Maximum = 100,
            Width = 90,
            BackColor = Color.White,
            ForeColor = TextColor
        };
        _addButton = new Button
        {
            Text = "タスクを追加",
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = AccentColor,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = _buttonFont,
            Padding = new Padding(14, 6, 14, 6),
            Cursor = Cursors.Hand
        };
        _addButton.FlatAppearance.BorderSize = 0;
        _addButton.Click += (_, _) => AddTask();
        AcceptButton = _addButton;

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Padding = new Padding(14, 14, 14, 14)
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        var headerPanel = BuildHeaderPanel();
        var mainTabs = BuildMainTabs();

        root.Controls.Add(headerPanel, 0, 0);
        root.Controls.Add(mainTabs, 0, 1);

        Controls.Add(root);

        _noteToastTimer = new System.Windows.Forms.Timer { Interval = 2000 };
        _noteToastTimer.Tick += (_, _) =>
        {
            _noteToastTimer.Stop();
            if (_noteToastLabel != null)
            {
                _noteToastLabel.Visible = false;
            }
        };

        _trayMenu = new ContextMenuStrip();
        _trayMenu.Items.Add("開く", null, (_, _) => ShowFromTray());
        _trayMenu.Items.Add(new ToolStripSeparator());
        _trayMenu.Items.Add("終了", null, (_, _) => ExitFromTray());

        _rowMenu = new ContextMenuStrip();
        _rowMenuTimerItem = new ToolStripMenuItem("タイマー開始");
        _rowMenuTimerSeparator = new ToolStripSeparator();
        _rowMenuToggleItem = new ToolStripMenuItem("完了にする");
        _rowMenuDeleteItem = new ToolStripMenuItem("削除");
        _rowMenuTimerItem.Click += HandleRowMenuTimer;
        _rowMenuToggleItem.Click += HandleRowMenuToggle;
        _rowMenuDeleteItem.Click += HandleRowMenuDelete;
        _rowMenu.Items.Add(_rowMenuTimerItem);
        _rowMenu.Items.Add(_rowMenuTimerSeparator);
        _rowMenu.Items.Add(_rowMenuToggleItem);
        _rowMenu.Items.Add(new ToolStripSeparator());
        _rowMenu.Items.Add(_rowMenuDeleteItem);

        _trayIcon = new NotifyIcon
        {
            Text = AppName,
            Icon = _appIcon,
            ContextMenuStrip = _trayMenu,
            Visible = true
        };
        _trayIcon.DoubleClick += (_, _) => ShowFromTray();

        FormClosing += MainForm_FormClosing;
        Resize += MainForm_Resize;
        Shown += (_, _) =>
        {
            UpdateGridLayout(_activeGrid, isCompleted: false);
            UpdateGridLayout(_completedGrid, isCompleted: true);
        };

        LoadTasks();
        UpdateSummary();
        LoadNotes();

        _trackingTimer = new System.Windows.Forms.Timer { Interval = 1000 };
        _trackingTimer.Tick += (_, _) => UpdateTrackingDisplay();
        _trackingTimer.Start();
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _trayIcon.Visible = false;
        _trayIcon.Dispose();
        _rowMenu.Dispose();
        _trayMenu.Dispose();
        _trackingTimer.Stop();
        _trackingTimer.Dispose();
        _noteToastTimer.Stop();
        _noteToastTimer.Dispose();
        if (_ownsIcon)
        {
            _appIcon.Dispose();
        }
        _buttonFont.Dispose();
        _hintFont.Dispose();
        _sectionFont.Dispose();
        _headerFont.Dispose();
        _baseFont.Dispose();
        base.OnFormClosed(e);
    }

    private static Font CreateUiFont(string family, float size, FontStyle style)
    {
        try
        {
            return new Font(family, size, style);
        }
        catch
        {
            var fallback = SystemFonts.MessageBoxFont ?? SystemFonts.DefaultFont;
            var fontFamily = fallback?.FontFamily ?? FontFamily.GenericSansSerif;
            return new Font(fontFamily, size, style);
        }
    }

    private (Icon Icon, bool OwnsIcon) LoadAppIcon()
    {
        try
        {
            var iconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "app.ico");
            if (File.Exists(iconPath))
            {
                return (new Icon(iconPath), true);
            }
        }
        catch
        {
        }

        return (SystemIcons.Application, false);
    }

    private Control BuildHeaderPanel()
    {
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 1,
            Margin = new Padding(0, 0, 0, 4)
        };
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var titleLabel = new Label
        {
            Text = AppName,
            AutoSize = true,
            Font = _headerFont,
            ForeColor = TextColor
        };
        var subtitleLabel = new Label
        {
            Text = "気軽にタスク管理、進捗を見える化。",
            AutoSize = true,
            Font = _hintFont,
            ForeColor = MutedTextColor,
            Margin = new Padding(2, 2, 0, 0)
        };

        panel.Controls.Add(titleLabel, 0, 0);
        panel.Controls.Add(subtitleLabel, 0, 1);

        return panel;
    }

    private Control BuildMainTabs()
    {
        _mainTabControl = new TabControl
        {
            Dock = DockStyle.Fill
        };

        _tasksMainTabPage = new TabPage("タスク") { Padding = new Padding(6) };
        _notesMainTabPage = new TabPage("メモ") { Padding = new Padding(6) };

        _tasksMainTabPage.Controls.Add(BuildTasksTabPanel());
        _notesMainTabPage.Controls.Add(BuildNotesPanel());

        _mainTabControl.TabPages.Add(_tasksMainTabPage);
        _mainTabControl.TabPages.Add(_notesMainTabPage);
        _lastMainTabPage = _mainTabControl.SelectedTab;
        _mainTabControl.SelectedIndexChanged += (_, _) => HandleMainTabChanged();

        return _mainTabControl;
    }

    private Control BuildTasksTabPanel()
    {
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2
        };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        var inputPanel = BuildInputPanel();
        var tasksPanel = BuildTasksPanel();

        layout.Controls.Add(inputPanel, 0, 0);
        layout.Controls.Add(tasksPanel, 0, 1);

        return layout;
    }

    private Control BuildNotesPanel()
    {
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2
        };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        var headerPanel = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = false,
            FlowDirection = FlowDirection.LeftToRight,
            Dock = DockStyle.Top
        };
        var headerLabel = new Label
        {
            Text = "メモ",
            AutoSize = true,
            Font = _sectionFont,
            ForeColor = TextColor
        };
        var hintLabel = new Label
        {
            Text = "タイトルを付けて保存できます",
            AutoSize = true,
            Font = _hintFont,
            ForeColor = MutedTextColor,
            Margin = new Padding(12, 6, 0, 0)
        };
        headerPanel.Controls.Add(headerLabel);
        headerPanel.Controls.Add(hintLabel);

        var contentLayout = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Vertical,
            SplitterWidth = 6,
            BackColor = BackgroundColor,
            BorderStyle = BorderStyle.None
        };
        contentLayout.Panel1.Padding = new Padding(0, 8, 8, 0);
        contentLayout.Panel2.Padding = new Padding(8, 8, 0, 0);
        contentLayout.HandleCreated += (_, _) => EnsureNotesSplitterDistance(contentLayout, preferInitialSplit: true);
        contentLayout.SizeChanged += (_, _) => EnsureNotesSplitterDistance(contentLayout, preferInitialSplit: false);

        var listCard = CreateCardPanel();
        listCard.Dock = DockStyle.Fill;
        listCard.Padding = new Padding(12);
        listCard.Margin = new Padding(0);

        var listLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2
        };
        listLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        listLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        var listLabel = new Label
        {
            Text = "メモ一覧",
            AutoSize = true,
            Font = _sectionFont,
            ForeColor = TextColor
        };

        _notesTreeView = new TreeView
        {
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.None,
            BackColor = Color.White,
            ForeColor = TextColor,
            Font = _baseFont,
            HideSelection = false,
            FullRowSelect = true,
            ShowLines = true,
            ShowPlusMinus = true,
            ShowRootLines = true
        };
        _notesTreeView.AfterSelect += NotesTreeView_AfterSelect;

        listLayout.Controls.Add(listLabel, 0, 0);
        listLayout.Controls.Add(_notesTreeView, 0, 1);
        listCard.Controls.Add(listLayout);

        var editorCard = CreateCardPanel();
        editorCard.Dock = DockStyle.Fill;
        editorCard.Padding = new Padding(12);
        editorCard.Margin = new Padding(0);

        var editorLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 6,
            Padding = new Padding(4, 2, 4, 2)
        };
        editorLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        editorLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        editorLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        editorLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        editorLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        editorLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        editorLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        editorLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var noteFolderLabel = new Label
        {
            Text = "フォルダ",
            AutoSize = true,
            Anchor = AnchorStyles.Left
        };

        _noteFolderInput = new ComboBox
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            ForeColor = TextColor,
            DropDownStyle = ComboBoxStyle.DropDown,
            AutoCompleteMode = AutoCompleteMode.SuggestAppend,
            AutoCompleteSource = AutoCompleteSource.ListItems,
            IntegralHeight = false
        };

        var noteTitleLabel = new Label
        {
            Text = "タイトル",
            AutoSize = true,
            Anchor = AnchorStyles.Left
        };

        _noteTitleInput = new TextBox
        {
            Dock = DockStyle.Fill,
            PlaceholderText = "メモのタイトル",
            BackColor = Color.White,
            ForeColor = TextColor
        };

        var noteBodyLabel = new Label
        {
            Text = "メモ",
            AutoSize = true,
            Anchor = AnchorStyles.Left | AnchorStyles.Top,
            Margin = new Padding(0, 6, 0, 0)
        };

        var formatPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            WrapContents = false,
            FlowDirection = FlowDirection.LeftToRight,
            Margin = new Padding(0, 2, 0, 2)
        };

        _noteBoldButton = new Button
        {
            Text = "太字",
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = Color.White,
            ForeColor = TextColor,
            FlatStyle = FlatStyle.Flat,
            Font = _buttonFont,
            Padding = new Padding(4, 1, 4, 1),
            Cursor = Cursors.Hand
        };
        _noteBoldButton.FlatAppearance.BorderColor = BorderColor;
        _noteBoldButton.FlatAppearance.BorderSize = 1;
        _noteBoldButton.Click += (_, _) => ApplyNoteStyle(FontStyle.Bold);

        _noteUnderlineButton = new Button
        {
            Text = "下線",
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = Color.White,
            ForeColor = TextColor,
            FlatStyle = FlatStyle.Flat,
            Font = _buttonFont,
            Padding = new Padding(4, 1, 4, 1),
            Cursor = Cursors.Hand
        };
        _noteUnderlineButton.FlatAppearance.BorderColor = BorderColor;
        _noteUnderlineButton.FlatAppearance.BorderSize = 1;
        _noteUnderlineButton.Click += (_, _) => ApplyNoteStyle(FontStyle.Underline);

        _noteStrikeButton = new Button
        {
            Text = "取消",
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = Color.White,
            ForeColor = TextColor,
            FlatStyle = FlatStyle.Flat,
            Font = _buttonFont,
            Padding = new Padding(4, 1, 4, 1),
            Cursor = Cursors.Hand
        };
        _noteStrikeButton.FlatAppearance.BorderColor = BorderColor;
        _noteStrikeButton.FlatAppearance.BorderSize = 1;
        _noteStrikeButton.Click += (_, _) => ApplyNoteStyle(FontStyle.Strikeout);

        var zoomLabel = new Label
        {
            Text = "ズーム",
            AutoSize = true,
            ForeColor = MutedTextColor,
            Font = _hintFont,
            Margin = new Padding(12, 8, 4, 0)
        };

        _noteZoomOutButton = new Button
        {
            Text = "－",
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = Color.White,
            ForeColor = TextColor,
            FlatStyle = FlatStyle.Flat,
            Font = _buttonFont,
            Padding = new Padding(4, 1, 4, 1),
            Cursor = Cursors.Hand
        };
        _noteZoomOutButton.FlatAppearance.BorderColor = BorderColor;
        _noteZoomOutButton.FlatAppearance.BorderSize = 1;
        _noteZoomOutButton.Click += (_, _) => UpdateNoteZoom(_noteBodyInput.ZoomFactor - NoteZoomStep);

        _noteZoomInButton = new Button
        {
            Text = "＋",
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = Color.White,
            ForeColor = TextColor,
            FlatStyle = FlatStyle.Flat,
            Font = _buttonFont,
            Padding = new Padding(4, 1, 4, 1),
            Cursor = Cursors.Hand
        };
        _noteZoomInButton.FlatAppearance.BorderColor = BorderColor;
        _noteZoomInButton.FlatAppearance.BorderSize = 1;
        _noteZoomInButton.Click += (_, _) => UpdateNoteZoom(_noteBodyInput.ZoomFactor + NoteZoomStep);

        _noteZoomResetButton = new Button
        {
            Text = "100%",
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = Color.White,
            ForeColor = TextColor,
            FlatStyle = FlatStyle.Flat,
            Font = _buttonFont,
            Padding = new Padding(4, 1, 4, 1),
            Cursor = Cursors.Hand
        };
        _noteZoomResetButton.FlatAppearance.BorderColor = BorderColor;
        _noteZoomResetButton.FlatAppearance.BorderSize = 1;
        _noteZoomResetButton.Click += (_, _) => UpdateNoteZoom(1.0f);

        _noteZoomLabel = new Label
        {
            Text = "100%",
            AutoSize = true,
            ForeColor = MutedTextColor,
            Font = _hintFont,
            Margin = new Padding(6, 8, 0, 0)
        };

        formatPanel.Controls.Add(_noteBoldButton);
        formatPanel.Controls.Add(_noteUnderlineButton);
        formatPanel.Controls.Add(_noteStrikeButton);
        formatPanel.Controls.Add(zoomLabel);
        formatPanel.Controls.Add(_noteZoomOutButton);
        formatPanel.Controls.Add(_noteZoomInButton);
        formatPanel.Controls.Add(_noteZoomResetButton);
        formatPanel.Controls.Add(_noteZoomLabel);

        _noteBodyInput = new RichTextBox
        {
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.None,
            BackColor = Color.White,
            ForeColor = TextColor,
            Font = _baseFont,
            DetectUrls = false
        };
        _noteBodyInput.Enter += (_, _) => AcceptButton = null;
        _noteBodyInput.Leave += (_, _) => RestoreAcceptButtonForCurrentTab();
        _noteBodyInput.KeyDown += (_, e) =>
        {
            if (e.Control && e.KeyCode == Keys.B)
            {
                e.SuppressKeyPress = true;
                e.Handled = true;
                ApplyNoteStyle(FontStyle.Bold);
                return;
            }

            if (e.Control && e.Shift && e.KeyCode == Keys.OemMinus)
            {
                e.SuppressKeyPress = true;
                e.Handled = true;
                ApplyNoteStyle(FontStyle.Underline);
                return;
            }

            if (e.Control && (e.KeyCode == Keys.OemMinus || e.KeyCode == Keys.Subtract))
            {
                e.SuppressKeyPress = true;
                e.Handled = true;
                ApplyNoteStyle(FontStyle.Strikeout);
                return;
            }

            if (e.Control && e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                e.Handled = true;
                SaveNote();
            }
        };
        _noteBodyInput.SelectionChanged += (_, _) => UpdateNoteFormatButtons();
        _noteBodyInput.MouseWheel += (_, e) =>
        {
            if ((ModifierKeys & Keys.Control) == Keys.Control)
            {
                var delta = e.Delta > 0 ? NoteZoomStep : -NoteZoomStep;
                UpdateNoteZoom(_noteBodyInput.ZoomFactor + delta);
            }
        };
        UpdateNoteZoom(1.0f);
        UpdateNoteFormatButtons();

        _noteSaveButton = new Button
        {
            Text = "保存",
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = AccentColor,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = _buttonFont,
            Padding = new Padding(14, 6, 14, 6),
            Cursor = Cursors.Hand
        };
        _noteSaveButton.FlatAppearance.BorderSize = 0;
        _noteSaveButton.Click += (_, _) => SaveNote();

        _noteNewButton = new Button
        {
            Text = "新規",
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = Color.White,
            ForeColor = TextColor,
            FlatStyle = FlatStyle.Flat,
            Font = _buttonFont,
            Padding = new Padding(14, 6, 14, 6),
            Cursor = Cursors.Hand
        };
        _noteNewButton.FlatAppearance.BorderColor = BorderColor;
        _noteNewButton.FlatAppearance.BorderSize = 1;
        _noteNewButton.Click += (_, _) => CreateNewNote();

        _noteDeleteButton = new Button
        {
            Text = "削除",
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = Color.FromArgb(255, 228, 228),
            ForeColor = TextColor,
            FlatStyle = FlatStyle.Flat,
            Font = _buttonFont,
            Padding = new Padding(14, 6, 14, 6),
            Cursor = Cursors.Hand
        };
        _noteDeleteButton.FlatAppearance.BorderSize = 0;
        _noteDeleteButton.Click += (_, _) => DeleteNote();

        _noteToastLabel = new Label
        {
            Text = "保存しました",
            AutoSize = true,
            ForeColor = SuccessColor,
            Font = _hintFont,
            Visible = false,
            Anchor = AnchorStyles.Left,
            Margin = new Padding(0, 6, 0, 0)
        };

        var buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            WrapContents = false,
            Margin = new Padding(0, 10, 0, 0)
        };
        buttonPanel.Controls.Add(_noteSaveButton);
        buttonPanel.Controls.Add(_noteNewButton);
        buttonPanel.Controls.Add(_noteDeleteButton);

        var footerLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            AutoSize = true,
            Margin = new Padding(0, 10, 0, 0)
        };
        footerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        footerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        footerLayout.Controls.Add(_noteToastLabel, 0, 0);
        footerLayout.Controls.Add(buttonPanel, 1, 0);

        editorLayout.Controls.Add(noteFolderLabel, 0, 0);
        editorLayout.Controls.Add(_noteFolderInput, 1, 0);
        editorLayout.Controls.Add(noteTitleLabel, 0, 1);
        editorLayout.Controls.Add(_noteTitleInput, 1, 1);
        editorLayout.Controls.Add(noteBodyLabel, 0, 2);
        editorLayout.SetColumnSpan(noteBodyLabel, 2);
        editorLayout.Controls.Add(formatPanel, 0, 3);
        editorLayout.SetColumnSpan(formatPanel, 2);
        editorLayout.Controls.Add(_noteBodyInput, 0, 4);
        editorLayout.SetColumnSpan(_noteBodyInput, 2);
        editorLayout.Controls.Add(footerLayout, 0, 5);
        editorLayout.SetColumnSpan(footerLayout, 2);

        editorCard.Controls.Add(editorLayout);

        contentLayout.Panel1.Controls.Add(listCard);
        contentLayout.Panel2.Controls.Add(editorCard);

        layout.Controls.Add(headerPanel, 0, 0);
        layout.Controls.Add(contentLayout, 0, 1);

        return layout;
    }

    private void HandleMainTabChanged()
    {
        if (_mainTabControl is null)
        {
            return;
        }

        if (_lastMainTabPage == _notesMainTabPage && _mainTabControl.SelectedTab != _notesMainTabPage)
        {
            AutoSaveNoteDraft(showToast: true, showTrayNotification: false, keepSelection: true);
        }

        _lastMainTabPage = _mainTabControl.SelectedTab;
        RestoreAcceptButtonForCurrentTab();
    }

    private void RestoreAcceptButtonForCurrentTab()
    {
        if (_mainTabControl is null)
        {
            return;
        }

        AcceptButton = _mainTabControl.SelectedTab == _notesMainTabPage ? _noteSaveButton : _addButton;
    }

    private Control BuildInputPanel()
    {
        var card = CreateCardPanel();
        card.AutoSize = true;
        card.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        card.Dock = DockStyle.Top;
        card.Padding = new Padding(12);
        card.Margin = new Padding(0, 8, 0, 0);

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 2,
            Padding = new Padding(4, 2, 4, 2)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 80f));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var titleLabel = new Label
        {
            Text = "タイトル",
            AutoSize = true,
            Anchor = AnchorStyles.Left
        };
        panel.Controls.Add(titleLabel, 0, 0);
        panel.Controls.Add(_titleInput, 1, 0);

        var memoLabel = new Label
        {
            Text = "メモ",
            AutoSize = true,
            Anchor = AnchorStyles.Left | AnchorStyles.Top
        };
        panel.Controls.Add(memoLabel, 0, 1);
        panel.Controls.Add(_memoInput, 1, 1);

        var dueLabel = new Label
        {
            Text = "期限",
            AutoSize = true,
            Anchor = AnchorStyles.Left
        };
        panel.Controls.Add(dueLabel, 0, 2);
        panel.Controls.Add(_duePicker, 1, 2);

        var progressLabel = new Label
        {
            Text = "進捗",
            AutoSize = true,
            Anchor = AnchorStyles.Left
        };
        panel.Controls.Add(progressLabel, 0, 3);
        panel.Controls.Add(_progressInput, 1, 3);

        var buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            WrapContents = false,
            Margin = new Padding(0, 6, 0, 0)
        };
        buttonPanel.Controls.Add(_addButton);
        panel.Controls.Add(buttonPanel, 0, 4);
        panel.SetColumnSpan(buttonPanel, 2);

        card.Controls.Add(panel);
        return card;
    }

    private Control BuildTasksPanel()
    {
        var card = CreateCardPanel();
        card.Dock = DockStyle.Fill;
        card.Padding = new Padding(12);
        card.Margin = new Padding(0, 12, 0, 0);

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3
        };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var headerPanel = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = false,
            FlowDirection = FlowDirection.LeftToRight,
            Dock = DockStyle.Top
        };
        var headerLabel = new Label
        {
            Text = "タスク",
            AutoSize = true,
            Font = _sectionFont,
            ForeColor = TextColor
        };
        var hintLabel = new Label
        {
            Text = "進捗をクリックして編集 / ・・・で操作",
            AutoSize = true,
            Font = _hintFont,
            ForeColor = MutedTextColor,
            Margin = new Padding(12, 6, 0, 0)
        };
        headerPanel.Controls.Add(headerLabel);
        headerPanel.Controls.Add(hintLabel);

        var tabControl = new TabControl
        {
            Dock = DockStyle.Fill
        };
        _activeTabPage = new TabPage("タスク一覧") { Padding = new Padding(6) };
        _completedTabPage = new TabPage("完了済み") { Padding = new Padding(6) };

        _activeGrid = BuildActiveGrid();
        _completedGrid = BuildCompletedGrid();

        _activeTabPage.Controls.Add(_activeGrid);
        _completedTabPage.Controls.Add(_completedGrid);
        tabControl.TabPages.Add(_activeTabPage);
        tabControl.TabPages.Add(_completedTabPage);

        var summaryPanel = BuildSummaryPanel();
        summaryPanel.Margin = new Padding(0, 8, 0, 0);

        layout.Controls.Add(headerPanel, 0, 0);
        layout.Controls.Add(tabControl, 0, 1);
        layout.Controls.Add(summaryPanel, 0, 2);

        card.Controls.Add(layout);

        return card;
    }

    private DataGridView CreateBaseGrid(BindingSource source)
    {
        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = false,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            MultiSelect = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            DataSource = source,
            RowHeadersVisible = false,
            BackgroundColor = CardColor,
            BorderStyle = BorderStyle.None,
            CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
            GridColor = BorderColor,
            EnableHeadersVisualStyles = false,
            ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None,
            ColumnHeadersHeight = 30,
            RowTemplate = { Height = 32 }
        };
        grid.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = HeaderBackColor,
            ForeColor = TextColor,
            Font = _baseFont,
            Alignment = DataGridViewContentAlignment.MiddleLeft
        };
        grid.DefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = CardColor,
            ForeColor = TextColor,
            SelectionBackColor = SelectionColor,
            SelectionForeColor = TextColor,
            Font = _hintFont
        };
        grid.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Color.FromArgb(252, 249, 247),
            ForeColor = TextColor,
            SelectionBackColor = SelectionColor,
            SelectionForeColor = TextColor,
            Font = _hintFont
        };

        return grid;
    }

    private DataGridView BuildActiveGrid()
    {
        var grid = CreateBaseGrid(_activeBindingSource);

        var titleColumn = new DataGridViewTextBoxColumn
        {
            HeaderText = "タイトル",
            DataPropertyName = nameof(TaskItem.Title),
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 140,
            ReadOnly = true
        };

        var memoColumn = new DataGridViewTextBoxColumn
        {
            HeaderText = "メモ",
            DataPropertyName = nameof(TaskItem.Memo),
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 200
        };

        _activeDueColumn = new DataGridViewTextBoxColumn
        {
            HeaderText = "期限",
            DataPropertyName = nameof(TaskItem.DueDate),
            Width = 110,
            ReadOnly = false,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                Format = "yyyy-MM-dd",
                Alignment = DataGridViewContentAlignment.MiddleCenter
            }
        };

        _activeProgressColumn = new DataGridViewTextBoxColumn
        {
            HeaderText = "進捗",
            DataPropertyName = nameof(TaskItem.Progress),
            Width = 110,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                ForeColor = TextColor
            }
        };

        _activeWorkTimeColumn = new DataGridViewTextBoxColumn
        {
            HeaderText = "作業時間",
            DataPropertyName = nameof(TaskItem.WorkDisplay),
            Width = 120,
            ReadOnly = true,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                ForeColor = TextColor
            }
        };

        _activeTrackButtonColumn = new DataGridViewButtonColumn
        {
            HeaderText = "タイマー",
            Width = 72,
            FlatStyle = FlatStyle.Flat,
            UseColumnTextForButtonValue = false,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                ForeColor = TextColor
            }
        };

        _activeCompleteColumn = new DataGridViewButtonColumn
        {
            HeaderText = "",
            Width = 66,
            FlatStyle = FlatStyle.Flat,
            UseColumnTextForButtonValue = false,
            Visible = false,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                ForeColor = TextColor
            }
        };

        _activeDeleteColumn = new DataGridViewButtonColumn
        {
            HeaderText = "",
            Width = 66,
            FlatStyle = FlatStyle.Flat,
            UseColumnTextForButtonValue = false,
            Visible = false,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                ForeColor = TextColor
            }
        };

        _activeMenuColumn = new DataGridViewButtonColumn
        {
            HeaderText = "",
            Width = 52,
            FlatStyle = FlatStyle.Flat,
            UseColumnTextForButtonValue = false,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                ForeColor = TextColor
            }
        };

        grid.Columns.AddRange(
            titleColumn,
            memoColumn,
            _activeDueColumn,
            _activeProgressColumn,
            _activeWorkTimeColumn,
            _activeTrackButtonColumn,
            _activeCompleteColumn,
            _activeDeleteColumn,
            _activeMenuColumn);

        grid.CurrentCellDirtyStateChanged += TasksGrid_CurrentCellDirtyStateChanged;
        grid.CellClick += ActiveGrid_CellContentClick;
        grid.CellFormatting += ActiveGrid_CellFormatting;
        grid.CellValidating += TasksGrid_CellValidating;
        grid.CellEndEdit += (_, e) => grid.Rows[e.RowIndex].ErrorText = string.Empty;
        grid.CellPainting += TasksGrid_CellPainting;
        grid.DataError += TasksGrid_DataError;
        grid.SizeChanged += (_, _) => UpdateGridLayout(grid, isCompleted: false);

        return grid;
    }

    private DataGridView BuildCompletedGrid()
    {
        var grid = CreateBaseGrid(_completedBindingSource);

        var titleColumn = new DataGridViewTextBoxColumn
        {
            HeaderText = "タイトル",
            DataPropertyName = nameof(TaskItem.Title),
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 140,
            ReadOnly = true
        };

        var memoColumn = new DataGridViewTextBoxColumn
        {
            HeaderText = "メモ",
            DataPropertyName = nameof(TaskItem.Memo),
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 200,
            ReadOnly = true
        };

        _completedDueColumn = new DataGridViewTextBoxColumn
        {
            HeaderText = "期限",
            DataPropertyName = nameof(TaskItem.DueDate),
            Width = 110,
            ReadOnly = true,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                Format = "yyyy-MM-dd",
                Alignment = DataGridViewContentAlignment.MiddleCenter
            }
        };

        _completedProgressColumn = new DataGridViewTextBoxColumn
        {
            HeaderText = "進捗",
            DataPropertyName = nameof(TaskItem.Progress),
            Width = 110,
            ReadOnly = true,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                ForeColor = TextColor
            }
        };

        _completedWorkTimeColumn = new DataGridViewTextBoxColumn
        {
            HeaderText = "作業時間",
            DataPropertyName = nameof(TaskItem.WorkDisplay),
            Width = 120,
            ReadOnly = true,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                ForeColor = TextColor
            }
        };

        _completedRestoreColumn = new DataGridViewButtonColumn
        {
            HeaderText = "",
            Width = 66,
            FlatStyle = FlatStyle.Flat,
            UseColumnTextForButtonValue = false,
            Visible = false,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                ForeColor = TextColor
            }
        };

        _completedDeleteColumn = new DataGridViewButtonColumn
        {
            HeaderText = "",
            Width = 66,
            FlatStyle = FlatStyle.Flat,
            UseColumnTextForButtonValue = false,
            Visible = false,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                ForeColor = TextColor
            }
        };

        _completedMenuColumn = new DataGridViewButtonColumn
        {
            HeaderText = "",
            Width = 52,
            FlatStyle = FlatStyle.Flat,
            UseColumnTextForButtonValue = false,
            DefaultCellStyle = new DataGridViewCellStyle
            {
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                ForeColor = TextColor
            }
        };

        grid.Columns.AddRange(
            titleColumn,
            memoColumn,
            _completedDueColumn,
            _completedProgressColumn,
            _completedWorkTimeColumn,
            _completedRestoreColumn,
            _completedDeleteColumn,
            _completedMenuColumn);

        grid.CellClick += CompletedGrid_CellContentClick;
        grid.CellFormatting += CompletedGrid_CellFormatting;
        grid.CellPainting += TasksGrid_CellPainting;
        grid.DataError += TasksGrid_DataError;
        grid.SizeChanged += (_, _) => UpdateGridLayout(grid, isCompleted: true);

        return grid;
    }

    private Control BuildSummaryPanel()
    {
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            AutoSize = true,
            Padding = new Padding(2, 6, 2, 2)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

        _summaryLabel = new Label
        {
            Text = "タスクなし",
            AutoSize = true,
            ForeColor = MutedTextColor,
            Font = _hintFont
        };

        _overallProgress = new ProgressBar
        {
            Minimum = 0,
            Maximum = 100,
            Style = ProgressBarStyle.Continuous,
            Dock = DockStyle.Fill,
            Height = 12
        };

        panel.Controls.Add(_summaryLabel, 0, 0);
        panel.Controls.Add(_overallProgress, 1, 0);

        return panel;
    }

    private void AddTask()
    {
        var title = _titleInput.Text.Trim();
        if (string.IsNullOrWhiteSpace(title))
        {
            MessageBox.Show(this, "タイトルを入力してください。", AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            _titleInput.Focus();
            return;
        }

        var progress = (int)_progressInput.Value;
        var task = new TaskItem
        {
            Title = title,
            Memo = _memoInput.Text.Trim(),
            DueDate = _duePicker.Value.Date,
            Progress = progress,
            IsCompleted = progress >= 100
        };

        AttachTask(task);
        if (task.IsCompleted)
        {
            _completedTasks.Add(task);
        }
        else
        {
            _activeTasks.Add(task);
        }

        _titleInput.Clear();
        _memoInput.Clear();
        _progressInput.Value = 0;
        _duePicker.Value = DateTime.Today;
        _titleInput.Focus();
    }

    private void LoadTasks()
    {
        _suppressSave = true;
        try
        {
            if (!File.Exists(_dataPath))
            {
                return;
            }

            var json = File.ReadAllText(_dataPath);
            var tasks = JsonSerializer.Deserialize<List<TaskItem>>(json, _jsonOptions) ?? new List<TaskItem>();

            ClearTasks();
            foreach (var task in tasks)
            {
                task.Normalize();
                AttachTask(task);
                if (task.IsCompleted)
                {
                    _completedTasks.Add(task);
                }
                else
                {
                    _activeTasks.Add(task);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"タスクの読み込みに失敗しました: {ex.Message}", AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            _suppressSave = false;
        }
    }

    private void SaveTasks()
    {
        if (_suppressSave)
        {
            return;
        }

        try
        {
            var directory = Path.GetDirectoryName(_dataPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var allTasks = _activeTasks.Concat(_completedTasks).ToList();
            var json = JsonSerializer.Serialize(allTasks, _jsonOptions);
            File.WriteAllText(_dataPath, json);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"タスクの保存に失敗しました: {ex.Message}", AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void LoadNotes()
    {
        _suppressNoteSave = true;
        try
        {
            if (!File.Exists(_notesPath))
            {
                return;
            }

            var json = File.ReadAllText(_notesPath);
            var notes = JsonSerializer.Deserialize<List<NoteItem>>(json, _jsonOptions) ?? new List<NoteItem>();

            _notes.Clear();
            foreach (var note in notes)
            {
                note.Normalize();
                _notes.Add(note);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"メモの読み込みに失敗しました: {ex.Message}", AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            _suppressNoteSave = false;
            RefreshNotesTree();
        }
    }

    private void SaveNotes()
    {
        if (_suppressNoteSave)
        {
            return;
        }

        try
        {
            var directory = Path.GetDirectoryName(_notesPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var json = JsonSerializer.Serialize(_notes.ToList(), _jsonOptions);
            File.WriteAllText(_notesPath, json);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"メモの保存に失敗しました: {ex.Message}", AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void Notes_ListChanged(object? sender, ListChangedEventArgs e)
    {
        if (_suppressNoteSave)
        {
            return;
        }

        SaveNotes();
        RefreshNotesTree();
    }

    private void NotesTreeView_AfterSelect(object? sender, TreeViewEventArgs e)
    {
        if (_suppressNoteTreeSelection)
        {
            return;
        }

        AutoSaveNoteDraft(showToast: false, showTrayNotification: false, keepSelection: true);

        if (e.Node?.Tag is not NoteItem note)
        {
            _selectedNote = null;
            _noteTitleInput.Clear();
            _noteBodyInput.Clear();
            UpdateNoteZoom(1.0f);
            UpdateNoteFormatButtons();
            if (e.Node?.Tag is string folderPath)
            {
                _noteFolderInput.Text = folderPath;
            }
            return;
        }

        _selectedNote = note;
        _noteTitleInput.Text = note.Title;
        _noteFolderInput.Text = note.Folder;
        if (!string.IsNullOrWhiteSpace(note.Content) && note.Content.StartsWith("{\\rtf", StringComparison.Ordinal))
        {
            try
            {
                _noteBodyInput.Rtf = note.Content;
            }
            catch
            {
                _noteBodyInput.Text = note.Content;
            }
        }
        else
        {
            _noteBodyInput.Text = note.Content;
        }

        UpdateNoteFormatButtons();
    }

    private void RefreshNotesTree()
    {
        if (_notesTreeView is null)
        {
            return;
        }

        object? selectionTag = _notesTreeView.SelectedNode?.Tag ?? _selectedNote;
        _suppressNoteTreeSelection = true;
        _notesTreeView.BeginUpdate();
        _notesTreeView.Nodes.Clear();

        var folderNodes = new Dictionary<string, TreeNode>(StringComparer.Ordinal);
        var folderPaths = new SortedSet<string>(StringComparer.CurrentCulture);
        foreach (var _note in _notes)
        {
            var folderPath = NoteItem.NormalizeFolderPath(_note.Folder);
            TreeNode? parentNode = null;
            if (!string.IsNullOrWhiteSpace(folderPath))
            {
                var currentPath = string.Empty;
                foreach (var part in folderPath.Split('/', StringSplitOptions.RemoveEmptyEntries))
                {
                    currentPath = string.IsNullOrEmpty(currentPath) ? part : $"{currentPath}/{part}";
                    folderPaths.Add(currentPath);
                    if (!folderNodes.TryGetValue(currentPath, out var folderNode))
                    {
                        folderNode = new TreeNode(part) { Tag = currentPath };
                        if (parentNode is null)
                        {
                            _notesTreeView.Nodes.Add(folderNode);
                        }
                        else
                        {
                            parentNode.Nodes.Add(folderNode);
                        }

                        folderNodes[currentPath] = folderNode;
                    }

                    parentNode = folderNode;
                }
            }

            var noteNode = new TreeNode(_note.Title) { Tag = _note };
            if (parentNode is null)
            {
                _notesTreeView.Nodes.Add(noteNode);
            }
            else
            {
                parentNode.Nodes.Add(noteNode);
            }
        }

        _notesTreeView.EndUpdate();

        if (_noteFolderInput is not null)
        {
            var currentFolder = _noteFolderInput.Text;
            _noteFolderInput.BeginUpdate();
            _noteFolderInput.Items.Clear();
            foreach (var path in folderPaths)
            {
                _noteFolderInput.Items.Add(path);
            }
            _noteFolderInput.EndUpdate();
            if (!string.IsNullOrWhiteSpace(currentFolder))
            {
                _noteFolderInput.Text = currentFolder;
            }
        }

        if (selectionTag is NoteItem note)
        {
            SelectNoteNode(note);
        }
        else if (selectionTag is string folderPath)
        {
            SelectFolderNode(folderPath);
        }

        _suppressNoteTreeSelection = false;
    }

    private void EnsureNotesSplitterDistance(SplitContainer container, bool preferInitialSplit)
    {
        if (container.Width <= 0)
        {
            return;
        }

        var required = NotesListMinWidth + NotesEditorMinWidth + container.SplitterWidth;
        if (container.Width < required)
        {
            return;
        }

        if (container.Panel1MinSize != NotesListMinWidth)
        {
            container.Panel1MinSize = NotesListMinWidth;
        }

        if (container.Panel2MinSize != NotesEditorMinWidth)
        {
            container.Panel2MinSize = NotesEditorMinWidth;
        }

        var min = container.Panel1MinSize;
        var max = container.Width - container.Panel2MinSize - container.SplitterWidth;
        if (max < min)
        {
            return;
        }

        var desired = container.SplitterDistance;
        if (!_notesSplitterInitialized || preferInitialSplit)
        {
            desired = (int)Math.Round(container.Width * 0.36f);
            _notesSplitterInitialized = true;
        }

        desired = Math.Clamp(desired, min, max);
        if (container.SplitterDistance != desired)
        {
            container.SplitterDistance = desired;
        }
    }

    private void ClearNoteSelection()
    {
        if (_notesTreeView is null)
        {
            return;
        }

        _suppressNoteTreeSelection = true;
        _notesTreeView.SelectedNode = null;
        _suppressNoteTreeSelection = false;
    }

    private void SelectNoteNode(NoteItem note)
    {
        if (_notesTreeView is null)
        {
            return;
        }

        var node = FindNoteNode(_notesTreeView.Nodes, note);
        if (node is null)
        {
            return;
        }

        _notesTreeView.SelectedNode = node;
        node.EnsureVisible();
    }

    private void SelectFolderNode(string folderPath)
    {
        if (_notesTreeView is null)
        {
            return;
        }

        var normalized = NoteItem.NormalizeFolderPath(folderPath);
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return;
        }

        var node = FindFolderNode(_notesTreeView.Nodes, normalized);
        if (node is null)
        {
            return;
        }

        _notesTreeView.SelectedNode = node;
        node.Expand();
        node.EnsureVisible();
    }

    private static TreeNode? FindNoteNode(TreeNodeCollection nodes, NoteItem note)
    {
        foreach (TreeNode node in nodes)
        {
            if (ReferenceEquals(node.Tag, note))
            {
                return node;
            }

            var child = FindNoteNode(node.Nodes, note);
            if (child is not null)
            {
                return child;
            }
        }

        return null;
    }

    private static TreeNode? FindFolderNode(TreeNodeCollection nodes, string folderPath)
    {
        foreach (TreeNode node in nodes)
        {
            if (node.Tag is string tag && string.Equals(tag, folderPath, StringComparison.Ordinal))
            {
                return node;
            }

            var child = FindFolderNode(node.Nodes, folderPath);
            if (child is not null)
            {
                return child;
            }
        }

        return null;
    }

    private void SaveNote()
    {
        var title = _noteTitleInput.Text.Trim();
        if (string.IsNullOrWhiteSpace(title))
        {
            MessageBox.Show(this, "タイトルを入力してください。", AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            _noteTitleInput.Focus();
            return;
        }

        var content = _noteBodyInput.Rtf ?? string.Empty;
        var folder = NoteItem.NormalizeFolderPath(_noteFolderInput.Text);

        if (_selectedNote is null)
        {
            var note = new NoteItem { Title = title, Content = content, Folder = folder };
            note.Normalize();
            _selectedNote = note;
            _notes.Add(note);
            ShowNoteToast("保存しました");
            return;
        }

        _selectedNote.Title = title;
        _selectedNote.Content = content;
        _selectedNote.Folder = folder;
        ShowNoteToast("保存しました");
    }

    private void CreateNewNote()
    {
        AutoSaveNoteDraft(showToast: false, showTrayNotification: false, keepSelection: false);
        ClearNoteSelection();
        _selectedNote = null;
        _noteTitleInput.Clear();
        _noteBodyInput.Clear();
        UpdateNoteZoom(1.0f);
        UpdateNoteFormatButtons();
        _noteTitleInput.Focus();
    }

    private void DeleteNote()
    {
        if (_selectedNote is null)
        {
            MessageBox.Show(this, "削除するメモを選択してください。", AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var result = MessageBox.Show(this, "このメモを削除しますか？", AppName, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
        if (result != DialogResult.Yes)
        {
            return;
        }

        _notes.Remove(_selectedNote);
        _selectedNote = null;
        ClearNoteSelection();
        _noteTitleInput.Clear();
        _noteBodyInput.Clear();
        UpdateNoteZoom(1.0f);
        UpdateNoteFormatButtons();
    }

    private void ApplyNoteStyle(FontStyle style)
    {
        if (_noteBodyInput is null)
        {
            return;
        }

        var selectionFont = _noteBodyInput.SelectionFont ?? _noteBodyInput.Font ?? _baseFont;
        var currentStyle = selectionFont.Style;
        var newStyle = (currentStyle & style) == style ? currentStyle & ~style : currentStyle | style;
        _noteBodyInput.SelectionFont = new Font(selectionFont.FontFamily, selectionFont.Size, newStyle);
        _noteBodyInput.Focus();
        UpdateNoteFormatButtons();
    }

    private void UpdateNoteFormatButtons()
    {
        if (_noteBodyInput is null)
        {
            return;
        }

        var selectionFont = _noteBodyInput.SelectionFont ?? _noteBodyInput.Font ?? _baseFont;
        var style = selectionFont.Style;
        SetNoteFormatButtonState(_noteBoldButton, style.HasFlag(FontStyle.Bold));
        SetNoteFormatButtonState(_noteUnderlineButton, style.HasFlag(FontStyle.Underline));
        SetNoteFormatButtonState(_noteStrikeButton, style.HasFlag(FontStyle.Strikeout));
    }

    private void SetNoteFormatButtonState(Button button, bool isActive)
    {
        if (button is null)
        {
            return;
        }

        if (isActive)
        {
            button.BackColor = SelectionColor;
            button.ForeColor = AccentDarkColor;
            button.FlatAppearance.BorderColor = AccentColor;
            button.FlatAppearance.BorderSize = 1;
            return;
        }

        button.BackColor = Color.White;
        button.ForeColor = TextColor;
        button.FlatAppearance.BorderColor = BorderColor;
        button.FlatAppearance.BorderSize = 1;
    }

    private void UpdateNoteZoom(float newZoom)
    {
        if (_noteBodyInput is null)
        {
            return;
        }

        var clamped = Math.Clamp(newZoom, NoteZoomMin, NoteZoomMax);
        _noteBodyInput.ZoomFactor = clamped;
        UpdateNoteZoomLabel();
    }

    private void UpdateNoteZoomLabel()
    {
        if (_noteZoomLabel is null || _noteBodyInput is null)
        {
            return;
        }

        var percent = (int)Math.Round(_noteBodyInput.ZoomFactor * 100);
        _noteZoomLabel.Text = $"{percent}%";
    }

    private void ShowNoteToast(string message)
    {
        if (_noteToastLabel is null)
        {
            return;
        }

        _noteToastLabel.Text = message;
        _noteToastLabel.Visible = true;
        _noteToastTimer.Stop();
        _noteToastTimer.Start();
    }

    private void AutoSaveNoteOnHide(bool showTrayNotification)
    {
        AutoSaveNoteDraft(showToast: true, showTrayNotification: showTrayNotification, keepSelection: false);
    }

    private void AutoSaveNoteDraft(bool showToast, bool showTrayNotification, bool keepSelection)
    {
        if (_noteTitleInput is null || _noteBodyInput is null || _noteFolderInput is null || _notesTreeView is null)
        {
            return;
        }

        var title = _noteTitleInput.Text.Trim();
        var bodyText = _noteBodyInput.Text;
        if (string.IsNullOrWhiteSpace(title) && string.IsNullOrWhiteSpace(bodyText))
        {
            return;
        }

        var content = _noteBodyInput.Rtf ?? string.Empty;
        var normalizedTitle = string.IsNullOrWhiteSpace(title) ? "無題" : title;
        var normalizedFolder = NoteItem.NormalizeFolderPath(_noteFolderInput.Text);

        var isNew = _selectedNote is null;
        var isChanged = isNew
            || !string.Equals(_selectedNote!.Title, normalizedTitle, StringComparison.Ordinal)
            || !string.Equals(_selectedNote.Content, content, StringComparison.Ordinal)
            || !string.Equals(_selectedNote.Folder, normalizedFolder, StringComparison.Ordinal);
        if (!isChanged)
        {
            return;
        }

        if (isNew)
        {
            var note = new NoteItem { Title = normalizedTitle, Content = content, Folder = normalizedFolder };
            note.Normalize();
            _selectedNote = note;
            _notes.Add(note);
        }
        else
        {
            _selectedNote!.Title = normalizedTitle;
            _selectedNote.Content = content;
            _selectedNote.Folder = normalizedFolder;
        }

        if (showToast)
        {
            ShowNoteToast("自動保存しました");
        }

        if (showTrayNotification)
        {
            _trayIcon.ShowBalloonTip(1500, AppName, "メモを自動保存しました", ToolTipIcon.Info);
        }
    }

    private void ClearTasks()
    {
        foreach (var task in _activeTasks.Concat(_completedTasks))
        {
            DetachTask(task);
        }

        _activeTasks.Clear();
        _completedTasks.Clear();
    }

    private void AttachTask(TaskItem task)
    {
        task.PropertyChanged += Task_PropertyChanged;
    }

    private void DetachTask(TaskItem task)
    {
        task.PropertyChanged -= Task_PropertyChanged;
    }

    private void Task_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (_suppressMove || sender is not TaskItem task)
        {
            return;
        }

        if (e.PropertyName == nameof(TaskItem.IsCompleted))
        {
            if (task.IsCompleted && _activeTasks.Contains(task))
            {
                MoveToCompleted(task);
            }
            else if (!task.IsCompleted && _completedTasks.Contains(task))
            {
                RestoreTask(task);
            }
        }
    }

    private void MoveToCompleted(TaskItem task)
    {
        if (!_activeTasks.Contains(task))
        {
            return;
        }

        _suppressMove = true;
        _suppressSave = true;
        try
        {
            task.StopTracking();
            if (!task.IsCompleted)
            {
                task.IsCompleted = true;
            }

            _activeTasks.Remove(task);
            if (!_completedTasks.Contains(task))
            {
                _completedTasks.Add(task);
            }
        }
        finally
        {
            _suppressSave = false;
            _suppressMove = false;
        }

        UpdateSummary();
        SaveTasks();
    }

    private void RestoreTask(TaskItem task)
    {
        if (!_completedTasks.Contains(task))
        {
            return;
        }

        _suppressMove = true;
        _suppressSave = true;
        try
        {
            if (task.Progress >= 100)
            {
                task.Progress = 99;
            }

            task.IsCompleted = false;

            _completedTasks.Remove(task);
            if (!_activeTasks.Contains(task))
            {
                _activeTasks.Add(task);
            }
        }
        finally
        {
            _suppressSave = false;
            _suppressMove = false;
        }

        UpdateSummary();
        SaveTasks();
    }

    private void DeleteTask(TaskItem task)
    {
        var result = MessageBox.Show(
            this,
            "このタスクを削除しますか？",
            AppName,
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);

        if (result != DialogResult.Yes)
        {
            return;
        }

        _suppressSave = true;
        try
        {
            task.StopTracking();
            if (_activeTasks.Contains(task))
            {
                _activeTasks.Remove(task);
            }
            else if (_completedTasks.Contains(task))
            {
                _completedTasks.Remove(task);
            }

            DetachTask(task);
        }
        finally
        {
            _suppressSave = false;
        }

        UpdateSummary();
        SaveTasks();
    }

    private void ShowRowMenu(DataGridView grid, TaskItem task, bool isCompleted, Rectangle cellBounds)
    {
        if (grid.IsDisposed)
        {
            return;
        }

        var context = new RowMenuContext(task, isCompleted);
        _rowMenu.Tag = context;
        _rowMenuTimerItem.Tag = context;
        _rowMenuToggleItem.Tag = context;
        _rowMenuDeleteItem.Tag = context;
        var showTimer = !isCompleted && !_activeTrackButtonColumn.Visible;
        _rowMenuTimerItem.Visible = showTimer;
        _rowMenuTimerSeparator.Visible = showTimer;
        if (showTimer)
        {
            _rowMenuTimerItem.Text = task.IsTracking ? "タイマー停止" : "タイマー開始";
        }
        _rowMenuToggleItem.Text = isCompleted ? "復元" : "完了にする";

        var location = cellBounds.IsEmpty
            ? grid.PointToClient(Cursor.Position)
            : new Point(cellBounds.Left + cellBounds.Width / 2, cellBounds.Bottom);

        if (_rowMenu.Visible)
        {
            _rowMenu.Close();
        }

        BeginInvoke(new Action(() =>
        {
            if (grid.IsDisposed)
            {
                return;
            }

            _rowMenu.Show(grid, location);
        }));
    }

    private RowMenuContext? GetRowMenuContext(object? sender)
    {
        if (sender is ToolStripItem item && item.Tag is RowMenuContext itemContext)
        {
            return itemContext;
        }

        return _rowMenu.Tag as RowMenuContext;
    }

    private void HandleRowMenuTimer(object? sender, EventArgs e)
    {
        var context = GetRowMenuContext(sender);
        if (context is null || context.IsCompleted)
        {
            return;
        }

        if (context.Task.IsTracking)
        {
            context.Task.StopTracking();
        }
        else
        {
            context.Task.StartTracking();
        }

        if (!_activeGrid.IsDisposed)
        {
            _activeGrid.Invalidate();
        }
    }

    private void HandleRowMenuToggle(object? sender, EventArgs e)
    {
        var context = GetRowMenuContext(sender);
        if (context is null)
        {
            return;
        }

        if (context.IsCompleted)
        {
            RestoreTask(context.Task);
        }
        else
        {
            MoveToCompleted(context.Task);
        }
    }

    private void HandleRowMenuDelete(object? sender, EventArgs e)
    {
        var context = GetRowMenuContext(sender);
        if (context is null)
        {
            return;
        }

        DeleteTask(context.Task);
    }

    private void UpdateActionColumnLayout(DataGridView grid, bool isCompleted)
    {
        if (grid.IsDisposed)
        {
            return;
        }

        var showButtons = grid.ClientSize.Width >= ActionMenuSwitchWidth;
        if (isCompleted)
        {
            _completedMenuColumn.Visible = !showButtons;
            _completedRestoreColumn.Visible = showButtons;
            _completedDeleteColumn.Visible = showButtons;
        }
        else
        {
            _activeMenuColumn.Visible = !showButtons;
            _activeCompleteColumn.Visible = showButtons;
            _activeDeleteColumn.Visible = showButtons;
        }
    }

    private void UpdateGridLayout(DataGridView grid, bool isCompleted)
    {
        if (grid.IsDisposed)
        {
            return;
        }

        UpdateActionColumnLayout(grid, isCompleted);

        var width = grid.ClientSize.Width;
        var showDue = width >= 760;
        var showWorkTime = width >= 700;
        var progressWidth = width >= 860 ? 110 : width >= 740 ? 96 : 84;
        var dueWidth = width >= 860 ? 110 : 96;
        var workWidth = width >= 860 ? 120 : 106;

        if (isCompleted)
        {
            _completedDueColumn.Visible = showDue;
            _completedWorkTimeColumn.Visible = showWorkTime;
            _completedProgressColumn.Width = progressWidth;
            if (showDue)
            {
                _completedDueColumn.Width = dueWidth;
            }

            if (showWorkTime)
            {
                _completedWorkTimeColumn.Width = workWidth;
            }

            return;
        }

        var showTimer = width >= 660;
        var timerWidth = width >= 860 ? 72 : 60;

        _activeDueColumn.Visible = showDue;
        _activeWorkTimeColumn.Visible = showWorkTime;
        _activeTrackButtonColumn.Visible = showTimer;
        _activeProgressColumn.Width = progressWidth;
        if (showDue)
        {
            _activeDueColumn.Width = dueWidth;
        }

        if (showWorkTime)
        {
            _activeWorkTimeColumn.Width = workWidth;
        }

        if (showTimer)
        {
            _activeTrackButtonColumn.Width = timerWidth;
        }
    }

    private void UpdateSummary()
    {
        var allTasks = _activeTasks.Concat(_completedTasks).ToList();
        var total = allTasks.Count;
        if (total == 0)
        {
            _summaryLabel.Text = "タスクなし";
            _overallProgress.Value = 0;
            UpdateTabTitles();
            return;
        }

        var completed = _completedTasks.Count;
        var average = (int)Math.Round(allTasks.Average(task => task.Progress));
        var active = _activeTasks.Count;

        _summaryLabel.Text = $"全体 {average}%（{completed}/{total} 完了 / 残り {active}）";
        _overallProgress.Value = Math.Clamp(average, 0, 100);
        UpdateTabTitles();
    }

    private void UpdateTabTitles()
    {
        if (_activeTabPage != null)
        {
            _activeTabPage.Text = $"タスク一覧 ({_activeTasks.Count})";
        }

        if (_completedTabPage != null)
        {
            _completedTabPage.Text = $"完了済み ({_completedTasks.Count})";
        }
    }

    private void Tasks_ListChanged(object? sender, ListChangedEventArgs e)
    {
        if (_suppressSave)
        {
            return;
        }

        UpdateSummary();
        SaveTasks();
    }

    private void TasksGrid_CurrentCellDirtyStateChanged(object? sender, EventArgs e)
    {
        if (sender is not DataGridView grid || !grid.IsCurrentCellDirty)
        {
            return;
        }

        if (grid.CurrentCell is DataGridViewCheckBoxCell || grid.CurrentCell is DataGridViewComboBoxCell)
        {
            grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }
    }

    private void ActiveGrid_CellContentClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (sender is not DataGridView grid || e.RowIndex < 0)
        {
            return;
        }

        if (grid.Rows[e.RowIndex].DataBoundItem is not TaskItem task)
        {
            return;
        }

        if (e.ColumnIndex == _activeTrackButtonColumn.Index)
        {
            if (task.IsTracking)
            {
                task.StopTracking();
            }
            else
            {
                task.StartTracking();
            }

            grid.InvalidateRow(e.RowIndex);
            return;
        }

        if (e.ColumnIndex == _activeCompleteColumn.Index)
        {
            MoveToCompleted(task);
            return;
        }

        if (e.ColumnIndex == _activeDeleteColumn.Index)
        {
            DeleteTask(task);
            return;
        }

        if (e.ColumnIndex == _activeMenuColumn.Index)
        {
            var cellBounds = grid.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
            ShowRowMenu(grid, task, isCompleted: false, cellBounds);
        }
    }

    private void ActiveGrid_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (sender is not DataGridView grid || e.RowIndex < 0)
        {
            return;
        }

        if (grid.Rows[e.RowIndex].DataBoundItem is not TaskItem task)
        {
            return;
        }

        if (e.ColumnIndex == _activeTrackButtonColumn.Index)
        {
            e.Value = task.IsTracking ? "停止" : "開始";
            e.FormattingApplied = true;
            return;
        }

        if (e.ColumnIndex == _activeCompleteColumn.Index)
        {
            e.Value = "完了";
            e.FormattingApplied = true;
            return;
        }

        if (e.ColumnIndex == _activeDeleteColumn.Index)
        {
            e.Value = "削除";
            e.FormattingApplied = true;
            return;
        }

        if (e.ColumnIndex == _activeMenuColumn.Index)
        {
            e.Value = "・・・";
            e.FormattingApplied = true;
        }
    }

    private void CompletedGrid_CellContentClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (sender is not DataGridView grid || e.RowIndex < 0)
        {
            return;
        }

        if (grid.Rows[e.RowIndex].DataBoundItem is not TaskItem task)
        {
            return;
        }

        if (e.ColumnIndex == _completedRestoreColumn.Index)
        {
            RestoreTask(task);
            return;
        }

        if (e.ColumnIndex == _completedDeleteColumn.Index)
        {
            DeleteTask(task);
            return;
        }

        if (e.ColumnIndex == _completedMenuColumn.Index)
        {
            var cellBounds = grid.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
            ShowRowMenu(grid, task, isCompleted: true, cellBounds);
        }
    }

    private void CompletedGrid_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (sender is not DataGridView grid || e.RowIndex < 0)
        {
            return;
        }

        if (e.ColumnIndex == _completedRestoreColumn.Index)
        {
            e.Value = "復元";
            e.FormattingApplied = true;
            return;
        }

        if (e.ColumnIndex == _completedDeleteColumn.Index)
        {
            e.Value = "削除";
            e.FormattingApplied = true;
            return;
        }

        if (e.ColumnIndex == _completedMenuColumn.Index)
        {
            e.Value = "・・・";
            e.FormattingApplied = true;
        }
    }

    private void TasksGrid_CellValidating(object? sender, DataGridViewCellValidatingEventArgs e)
    {
        if (sender is not DataGridView grid || grid != _activeGrid)
        {
            return;
        }

        if (e.RowIndex < 0 || e.ColumnIndex != _activeProgressColumn.Index)
        {
            return;
        }

        var text = e.FormattedValue?.ToString() ?? string.Empty;
        if (!int.TryParse(text, out var value) || value < 0 || value > 100)
        {
            e.Cancel = true;
            grid.Rows[e.RowIndex].ErrorText = "進捗は0〜100の数値で入力してください。";
        }
    }

    private void TasksGrid_CellPainting(object? sender, DataGridViewCellPaintingEventArgs e)
    {
        if (sender is not DataGridView grid || e.RowIndex < 0)
        {
            return;
        }

        var progressColumn = grid == _activeGrid ? _activeProgressColumn : _completedProgressColumn;
        if (e.ColumnIndex != progressColumn.Index)
        {
            return;
        }

        var graphics = e.Graphics;
        if (graphics is null)
        {
            return;
        }

        e.Handled = true;
        e.PaintBackground(e.CellBounds, true);

        var valueText = e.Value?.ToString() ?? "0";
        var progress = int.TryParse(valueText, out var parsed) ? parsed : 0;
        progress = Math.Clamp(progress, 0, 100);

        var barBounds = new Rectangle(
            e.CellBounds.X + 4,
            e.CellBounds.Y + 6,
            Math.Max(0, e.CellBounds.Width - 8),
            Math.Max(0, e.CellBounds.Height - 12));

        using var backBrush = new SolidBrush(ProgressTrackColor);
        graphics.FillRectangle(backBrush, barBounds);

        var fillWidth = (int)Math.Round(barBounds.Width * (progress / 100.0));
        var fillRect = new Rectangle(barBounds.X, barBounds.Y, fillWidth, barBounds.Height);
        using var fillBrush = new SolidBrush(progress >= 100 ? SuccessColor : AccentDarkColor);
        graphics.FillRectangle(fillBrush, fillRect);

        using var borderPen = new Pen(BorderColor);
        graphics.DrawRectangle(borderPen, barBounds);

        var cellFont = e.CellStyle?.Font ?? _hintFont;
        TextRenderer.DrawText(
            graphics,
            $"{progress}%",
            cellFont,
            e.CellBounds,
            TextColor,
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);

        e.Paint(e.CellBounds, DataGridViewPaintParts.Border);
    }

    private void TasksGrid_DataError(object? sender, DataGridViewDataErrorEventArgs e)
    {
        MessageBox.Show(this, "無効な値です。", AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        e.ThrowException = false;
    }

    private void UpdateTrackingDisplay()
    {
        if (_activeGrid.IsDisposed || _activeGrid.Columns.Count == 0)
        {
            return;
        }

        if (_activeTasks.Any(task => task.IsTracking))
        {
            _activeGrid.InvalidateColumn(_activeWorkTimeColumn.Index);
        }
    }

    private void StopAllTracking()
    {
        if (_activeTasks.Count == 0)
        {
            return;
        }

        _suppressSave = true;
        try
        {
            foreach (var task in _activeTasks.Where(task => task.IsTracking))
            {
                task.StopTracking();
            }
        }
        finally
        {
            _suppressSave = false;
        }

        UpdateSummary();
        SaveTasks();
    }

    private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
    {
        if (_allowExit)
        {
            StopAllTracking();
            return;
        }

        e.Cancel = true;
        HideToTray();
    }

    private void MainForm_Resize(object? sender, EventArgs e)
    {
        if (WindowState == FormWindowState.Minimized)
        {
            HideToTray();
        }
    }

    private void HideToTray()
    {
        AutoSaveNoteOnHide(showTrayNotification: true);
        Hide();
        ShowInTaskbar = false;

        if (!_trayTipShown)
        {
            _trayIcon.ShowBalloonTip(1000, AppName, "バックグラウンドで動作中です。", ToolTipIcon.Info);
            _trayTipShown = true;
        }
    }

    private void ShowFromTray()
    {
        Show();
        ShowInTaskbar = true;
        WindowState = FormWindowState.Normal;
        Activate();
    }

    private void ExitFromTray()
    {
        AutoSaveNoteOnHide(showTrayNotification: true);
        _allowExit = true;
        _trayIcon.Visible = false;
        Close();
    }

    private CardPanel CreateCardPanel()
    {
        return new CardPanel
        {
            CornerRadius = 14,
            BorderColor = BorderColor,
            FillColor = CardColor,
            BorderThickness = 1,
            BackColor = BackgroundColor
        };
    }

    private sealed class CardPanel : Panel
    {
        public int CornerRadius { get; set; } = 12;
        public int BorderThickness { get; set; } = 1;
        public Color BorderColor { get; set; } = Color.FromArgb(230, 230, 230);
        public Color FillColor { get; set; } = Color.White;

        public CardPanel()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            DoubleBuffered = true;
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            Invalidate();
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            e.Graphics.Clear(BackColor);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            var bounds = ClientRectangle;
            bounds.Inflate(-1, -1);
            using var path = CreateRoundedPath(bounds, CornerRadius);
            using var fillBrush = new SolidBrush(FillColor);
            using var borderPen = new Pen(BorderColor, BorderThickness);
            e.Graphics.FillPath(fillBrush, path);
            e.Graphics.DrawPath(borderPen, path);
        }

        private static GraphicsPath CreateRoundedPath(Rectangle bounds, int radius)
        {
            if (bounds.Width <= 0 || bounds.Height <= 0)
            {
                return new GraphicsPath();
            }

            var safeRadius = Math.Max(0, Math.Min(radius, Math.Min(bounds.Width, bounds.Height) / 2));
            var path = new GraphicsPath();
            var diameter = safeRadius * 2;
            var arc = new Rectangle(bounds.X, bounds.Y, diameter, diameter);

            path.AddArc(arc, 180, 90);
            arc.X = bounds.Right - diameter;
            path.AddArc(arc, 270, 90);
            arc.Y = bounds.Bottom - diameter;
            path.AddArc(arc, 0, 90);
            arc.X = bounds.Left;
            path.AddArc(arc, 90, 90);
            path.CloseFigure();
            return path;
        }
    }
}
