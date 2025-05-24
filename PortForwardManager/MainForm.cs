using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Windows.Forms;
using System.Text;
using Newtonsoft.Json;

namespace PortForwardManager
{
    public partial class MainForm : Form
    {
        // Константы и пути
        private const string SettingsFileName = "portforward_settings.json";
        private const string BackupFolder = "Backups";
        private const string LogFileName = "portforward.log";
        private readonly string AppDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "PortForwardManager");

        // Данные приложения
        private AppSettings settings = new AppSettings();
        private RuleProfile currentProfile = new RuleProfile();
        private List<PortStatus> portStatuses = new List<PortStatus>();

        // Элементы управления
        private TextBox tbListenPort, tbTargetIP, tbTargetPort, tbDescription, txtLog;
        private ComboBox cbProtocol, cbProfiles;
        private ListBox lbRules;
        private DataGridView dgvPortStatus;
        private Button btnAdd, btnRemove, btnSaveProfile, btnNewProfile, btnDeleteProfile;
        private Button btnExportJson, btnImportJson, btnCheckPorts, btnShowVisualization;
        private CheckBox chkAutoBackup, chkFirewallRule;
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;
        private Panel graphPanel;
        private TabControl tabControl;

        public MainForm()
        {
            InitializeAppData();
            InitializeComponents();
            CheckAdminRights();
            SetupTrayIcon();
            LoadSettings();
            RefreshUI();
        }

        private void InitializeAppData()
        {
            if (!Directory.Exists(AppDataFolder))
                Directory.CreateDirectory(AppDataFolder);

            if (!Directory.Exists(Path.Combine(AppDataFolder, BackupFolder)))
                Directory.CreateDirectory(Path.Combine(AppDataFolder, BackupFolder));
        }

        private void InitializeComponents()
        {
            // Настройка главного окна
            this.Text = "Port Forward Manager";
            this.Size = new Size(900, 700);
            this.MinimumSize = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormClosing += MainForm_FormClosing;

            // Создаем вкладки                  
            tabControl = new TabControl { Dock = DockStyle.Fill };
            this.Controls.Add(tabControl);

            // Вкладка управления правилами
            var tabRules = new TabPage("Управление правилами");
            tabControl.TabPages.Add(tabRules);
            InitializeRulesTab(tabRules);

            // Вкладка профилей
            var tabProfiles = new TabPage("Профили");
            tabControl.TabPages.Add(tabProfiles);
            InitializeProfilesTab(tabProfiles);

            // Вкладка диагностики
            var tabDiagnostics = new TabPage("Диагностика");
            tabControl.TabPages.Add(tabDiagnostics);
            InitializeDiagnosticsTab(tabDiagnostics);

            // Вкладка настроек
            var tabSettings = new TabPage("Настройки");
            tabControl.TabPages.Add(tabSettings);
            InitializeSettingsTab(tabSettings);

            // Вкладка "О программе"
            var tabAbout = new TabPage("О программе");
            tabControl.TabPages.Add(tabAbout);
            InitializeAboutTab(tabAbout);
        }

        private void InitializeRulesTab(TabPage tab)
        {
            // Группа для ввода правил
            var groupInput = new GroupBox
            {
                Text = "Добавление правила",
                Location = new Point(10, 10),
                Size = new Size(400, 180)
            };
            tab.Controls.Add(groupInput);

            // Поля ввода
            tbListenPort = new TextBox { Location = new Point(120, 20), Width = 100 };
            tbTargetIP = new TextBox { Location = new Point(120, 50), Width = 150 };
            tbTargetPort = new TextBox { Location = new Point(120, 80), Width = 100 };
            tbDescription = new TextBox { Location = new Point(120, 110), Width = 250 };
            cbProtocol = new ComboBox { Location = new Point(120, 140), Width = 150 };

            // Заполнение протоколов
            cbProtocol.Items.AddRange(Enum.GetNames(typeof(PortProtocol)));
            cbProtocol.SelectedIndex = 0;

            // Метки
            groupInput.Controls.Add(new Label { Text = "Порт:", Location = new Point(20, 20) });
            groupInput.Controls.Add(tbListenPort);
            groupInput.Controls.Add(new Label { Text = "Целевой IP:", Location = new Point(20, 50) });
            groupInput.Controls.Add(tbTargetIP);
            groupInput.Controls.Add(new Label { Text = "Целевой порт:", Location = new Point(20, 80) });
            groupInput.Controls.Add(tbTargetPort);
            groupInput.Controls.Add(new Label { Text = "Описание:", Location = new Point(20, 110) });
            groupInput.Controls.Add(tbDescription);
            groupInput.Controls.Add(new Label { Text = "Протокол:", Location = new Point(20, 140) });
            groupInput.Controls.Add(cbProtocol);

            // Кнопки действий
            btnAdd = new Button { Text = "Добавить", Location = new Point(20, 200) };
            btnAdd.Click += (s, e) => AddPortRule();
            
            btnRemove = new Button { Text = "Удалить", Location = new Point(120, 200) };
            btnRemove.Click += (s, e) => RemovePortRule();
            
            chkFirewallRule = new CheckBox { 
                Text = "Открыть в брандмауэре", 
                Location = new Point(220, 200),
                Width = 150
            };
            
            tab.Controls.Add(btnAdd);
            tab.Controls.Add(btnRemove);
            tab.Controls.Add(chkFirewallRule);

            // Список правил
            lbRules = new ListBox
            {
                Location = new Point(10, 230),
                Size = new Size(400, 200),
                SelectionMode = SelectionMode.MultiExtended
            };
            lbRules.DoubleClick += (s, e) => LoadSelectedRule();
            tab.Controls.Add(lbRules);

            // Лог
            txtLog = new TextBox
            {
                Multiline = true,
                Location = new Point(420, 10),
                Size = new Size(450, 420),
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true
            };
            tab.Controls.Add(txtLog);
        }

        private void InitializeProfilesTab(TabPage tab)
        {
            // Выбор профиля
            cbProfiles = new ComboBox
            { 
                Location = new Point(20, 20), 
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            
            // Кнопки управления профилями
            btnNewProfile = new Button { Text = "Новый", Location = new Point(230, 20) };
            btnNewProfile.Click += (s, e) => CreateNewProfile();
            
            btnSaveProfile = new Button { Text = "Сохранить", Location = new Point(310, 20) };
            btnSaveProfile.Click += (s, e) => SaveCurrentProfile();
            
            btnDeleteProfile = new Button { Text = "Удалить", Location = new Point(400, 20) };
            btnDeleteProfile.Click += (s, e) => DeleteProfile();
            
            // Кнопки экспорта/импорта
            btnExportJson = new Button { Text = "Экспорт в JSON", Location = new Point(20, 60) };
            btnExportJson.Click += (s, e) => ExportToJson();
            
            btnImportJson = new Button { Text = "Импорт из JSON", Location = new Point(150, 60) };
            btnImportJson.Click += (s, e) => ImportFromJson();

            // Добавление элементов на вкладку
            tab.Controls.Add(new Label { Text = "Профили:", Location = new Point(20, 0) });
            tab.Controls.Add(cbProfiles);
            tab.Controls.Add(btnNewProfile);
            tab.Controls.Add(btnSaveProfile);
            tab.Controls.Add(btnDeleteProfile);
            tab.Controls.Add(btnExportJson);
            tab.Controls.Add(btnImportJson);
        }

        private void InitializeDiagnosticsTab(TabPage tab)
        {
            // Кнопка проверки портов
            btnCheckPorts = new Button
            { 
                Text = "Проверить порты", 
                Location = new Point(20, 20),
                Width = 150
            };
            btnCheckPorts.Click += (s, e) => CheckPorts();
            
            // Таблица статуса портов
            dgvPortStatus = new DataGridView
            {
                Location = new Point(20, 60),
                Size = new Size(500, 300),
                ReadOnly = true,
                AllowUserToAddRows = false
            };
            
            // Настройка колонок
            dgvPortStatus.Columns.Add("Port", "Порт");
            dgvPortStatus.Columns.Add("Status", "Статус");
            dgvPortStatus.Columns.Add("Service", "Сервис");
            dgvPortStatus.Columns.Add("Process", "Процесс");
            
            // Кнопка визуализации
            btnShowVisualization = new Button
            { 
                Text = "Показать схему", 
                Location = new Point(180, 20),
                Width = 150
            };
            btnShowVisualization.Click += (s, e) => ShowRulesVisualization();
            
            // Панель для графика
            graphPanel = new Panel
            {
                Location = new Point(530, 60),
                Size = new Size(300, 300),
                BorderStyle = BorderStyle.FixedSingle,
                AutoScroll = true
            };

            // Добавление элементов
            tab.Controls.Add(btnCheckPorts);
            tab.Controls.Add(dgvPortStatus);
            tab.Controls.Add(btnShowVisualization);
            tab.Controls.Add(graphPanel);
        }

        private void InitializeSettingsTab(TabPage tab)
        {
            // Настройки бэкапа
            chkAutoBackup = new CheckBox
            {
                Text = "Автоматическое создание резервных копий",
                Location = new Point(20, 20),
                Width = 300,
                Checked = settings.CreateBackups
            };
            chkAutoBackup.CheckedChanged += (s, e) =>
            {
                settings.CreateBackups = chkAutoBackup.Checked;
                SaveSettings();
            };
            
            // Максимальное количество бэкапов
            var lblMaxBackups = new Label
            { 
                Text = "Макс. количество бэкапов:", 
                Location = new Point(20, 50) 
            };
            
            var numMaxBackups = new NumericUpDown
            {
                Location = new Point(200, 50),
                Width = 100,
                Minimum = 1,
                Maximum = 100,
                Value = settings.MaxBackups
            };
            numMaxBackups.ValueChanged += (s, e) =>
            {
                settings.MaxBackups = (int)numMaxBackups.Value;
                SaveSettings();
            };
            
            // Кнопка очистки логов
            var btnClearLogs = new Button
            { 
                Text = "Очистить логи", 
                Location = new Point(20, 80) 
            };
            btnClearLogs.Click += (s, e) => ClearLogs();

            // Добавление элементов
            tab.Controls.Add(chkAutoBackup);
            tab.Controls.Add(lblMaxBackups);
            tab.Controls.Add(numMaxBackups);
            tab.Controls.Add(btnClearLogs);
        }

        private void InitializeAboutTab(TabPage tab)
        {
            var txtAbout = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Location = new Point(20, 20),
                Size = new Size(800, 220),
                Font = new Font("Segoe UI", 11F, FontStyle.Regular),
                Text =
                    "GUI интерфейс для PortForwarding для Windows Server 2016\r\n" +
                    "Разработан Мартынюком Сергеем для SoftServ.\r\n" +
                    "При прописывании правила форвардинга прописывается правило фаервола.\r\n" +
                    "При удалении правила форвардинга правило с файервола нужно удалять вручную (не дописал)\r\n" +
                    "Мартынюк Сергей\r\n" +
                    "admin@sinta-group.com.ua\r\n" +
                    "067-957-79-70"
            };
            tab.Controls.Add(txtAbout);
        }

        private void SetupTrayIcon()
        {
            trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("Открыть", null, (s, e) => ShowMainWindow());
            trayMenu.Items.Add("Выход", null, (s, e) => ExitApplication());

            // Используем собственную иконку из файла
            Icon trayAppIcon;
            string iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app.ico");
            if (File.Exists(iconPath))
            {
                trayAppIcon = new Icon(iconPath);
            }
            else
            {
                trayAppIcon = SystemIcons.Application; // fallback
            }

            trayIcon = new NotifyIcon
            {
                Icon = trayAppIcon,
                Text = "Port Forward Manager",
                ContextMenuStrip = trayMenu,
                Visible = true
            };

            trayIcon.DoubleClick += (s, e) => ShowMainWindow();
        }

        private void ShowMainWindow()
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.BringToFront();
        }

        private void ExitApplication()
        {
            trayIcon.Visible = false;
            Application.Exit();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
                trayIcon.ShowBalloonTip(1000, "Port Forward Manager", 
                    "Приложение продолжает работать в фоне", ToolTipIcon.Info);
            }
        }

        private void CheckAdminRights()
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            
            if (!principal.IsInRole(WindowsBuiltInRole.Administrator))
            {
                MessageBox.Show("Для корректной работы приложения требуются права администратора.",
                    "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void LoadSettings()
        {
            string settingsFile = Path.Combine(AppDataFolder, SettingsFileName);
            
            try
            {
                if (File.Exists(settingsFile))
                {
                    string json = File.ReadAllText(settingsFile);
                    settings = JsonConvert.DeserializeObject<AppSettings>(json) ?? new AppSettings();
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка загрузки настроек: {ex.Message}");
            }
        }

        private void SaveSettings()
        {
            string settingsFile = Path.Combine(AppDataFolder, SettingsFileName);
            
            try
            {
                string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                File.WriteAllText(settingsFile, json);
            }
            catch (Exception ex)
            {
                Log($"Ошибка сохранения настроек: {ex.Message}");
            }
        }

        private void RefreshUI()
        {
            RefreshProfilesList();
            RefreshRulesList();
            
            if (settings.Profiles.Any() && !string.IsNullOrEmpty(settings.LastActiveProfile))
            {
                cbProfiles.SelectedItem = settings.LastActiveProfile;
                LoadProfile(settings.LastActiveProfile);
            }
        }

        private void RefreshProfilesList()
        {
            cbProfiles.Items.Clear();
            cbProfiles.Items.AddRange(settings.Profiles.Select(p => p.Name).ToArray());
        }

        private void RefreshRulesList()
        {
            lbRules.Items.Clear();
            
            try
            {
                string output = ExecuteNetshCommand("interface portproxy show all");
                Log($"Вывод netsh:\n{output}");

                bool skipHeaders = true;
                foreach (var line in output.Split('\n'))
                {
                    if (string.IsNullOrWhiteSpace(line) || line.Contains("Прослушивать ipv4") || line.Contains("Адрес"))
                    {
                        skipHeaders = false;
                        continue;
                    }
                    
                    if (!skipHeaders && !line.Contains("---"))
                    {
                        var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length >= 4)
                        {
                            string formattedRule = $"{parts[0]}:{parts[1]} → {parts[2]}:{parts[3]}";
                            lbRules.Items.Add(formattedRule);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка обновления списка правил: {ex.Message}");
            }
        }

        private string ExecuteNetshCommand(string arguments)
        {
            try
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "netsh",
                        Arguments = arguments,
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true,
                        StandardOutputEncoding = Encoding.GetEncoding(866),
                        Verb = "runas"
                    }
                };
                
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                string errors = process.StandardError.ReadToEnd();
                process.WaitForExit();
                
                if (!string.IsNullOrEmpty(errors))
                    Log($"Ошибка netsh: {errors}");
                
                return output;
            }
            catch (Exception ex)
            {
                Log($"Ошибка выполнения команды netsh: {ex.Message}");
                return string.Empty;
            }
        }
        private void AddPortRule()
        {
            if (!ValidateInputs()) return;

            try
            {
                int listenPort = int.Parse(tbListenPort.Text);
                string targetIP = tbTargetIP.Text.ToLower() == "localhost" ? "127.0.0.1" : tbTargetIP.Text;
                int targetPort = int.Parse(tbTargetPort.Text);

                // Исправляем преобразование протокола
                string protocol = GetNetshProtocol(cbProtocol.SelectedItem.ToString());
                string description = tbDescription.Text;

                // Создаем резервную копию перед изменением
                if (settings.CreateBackups)
                    CreateBackup();

                // Добавляем правило
                string command = $"interface portproxy add {protocol} listenport={listenPort} " +
                               $"listenaddress=0.0.0.0 connectport={targetPort} connectaddress={targetIP}";

                Log($"Выполняем команду: netsh {command}");
                string output = ExecuteNetshCommand(command);
                Log($"Результат: {output}");

                // Открываем порт в брандмауэре если нужно
                if (chkFirewallRule.Checked)
                {
                    ExecuteNetshCommand(
                        $"advfirewall firewall add rule name=\"PortForward_{listenPort}\" " +
                        $"dir=in action=allow protocol=TCP localport={listenPort}");
                }

                // Обновляем UI
                RefreshRulesList();
                ClearInputs();

                Log($"Добавлено правило: {listenPort} → {targetIP}:{targetPort}");
            }
            catch (Exception ex)
            {
                Log($"Ошибка добавления правила: {ex.Message}");
                MessageBox.Show($"Ошибка добавления правила: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Новый метод для корректного преобразования протокола
        private string GetNetshProtocol(string protocol)
        {
            switch (protocol)
            {
                case "IPv4_to_IPv4": return "v4tov4";
                case "IPv4_to_IPv6": return "v4tov6";
                case "IPv6_to_IPv4": return "v6tov4";
                case "IPv6_to_IPv6": return "v6tov6";
                default: return "v4tov4";
            }
        }
        private bool ValidateInputs()
        {
            if (!int.TryParse(tbListenPort.Text, out int listenPort) || listenPort < 1 || listenPort > 65535)
            {
                MessageBox.Show("Некорректный номер порта (1-65535)");
                return false;
            }
            
            if (!int.TryParse(tbTargetPort.Text, out int targetPort) || targetPort < 1 || targetPort > 65535)
            {
                MessageBox.Show("Некорректный целевой порт (1-65535)");
                return false;
            }
            
            if (!IPAddress.TryParse(tbTargetIP.Text, out _) && tbTargetIP.Text != "localhost")
            {
                MessageBox.Show("Некорректный IP-адрес");
                return false;
            }
            
            return true;
        }

        private void RemovePortRule()
        {
            if (lbRules.SelectedItem == null)
            {
                MessageBox.Show("Выберите правило для удаления");
                return;
            }

            try
            {
                string selectedRule = lbRules.SelectedItem.ToString();
                var parts = selectedRule.Replace("→", "").Split(new[] { ' ', ':' }, StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length >= 4)
                {
                    // Исправляем определение протокола
                    string protocol = parts[0].Contains(":") ? "v6tov4" : "v4tov4";
                    int listenPort = int.Parse(parts[1]);
                    string listenAddress = parts[0] == "*" ? "0.0.0.0" : parts[0];

                    // Создаем резервную копию перед изменением
                    if (settings.CreateBackups)
                        CreateBackup();

                    // Удаляем правило
                    ExecuteNetshCommand(
                        $"interface portproxy delete {protocol} listenport={listenPort} listenaddress={listenAddress}");

                    // Обновляем UI
                    RefreshRulesList();

                    Log($"Удалено правило для порта {listenPort}");
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка удаления правила: {ex.Message}");
            }
        }
        private void LoadSelectedRule()
        {
            if (lbRules.SelectedItem == null) return;
            
            try
            {
                string selectedRule = lbRules.SelectedItem.ToString();
                var parts = selectedRule.Replace("→", "").Split(new[] { ' ', ':' }, StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length >= 4)
                {
                    tbListenPort.Text = parts[1];
                    tbTargetIP.Text = parts[2];
                    tbTargetPort.Text = parts[3];
                    
                    if (parts[0].Contains(":"))
                        cbProtocol.SelectedIndex = cbProtocol.FindStringExact("IPv6_to_IPv4");
                    else
                        cbProtocol.SelectedIndex = cbProtocol.FindStringExact("IPv4_to_IPv4");
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка загрузки правила: {ex.Message}");
            }
        }

        private void ClearInputs()
        {
            tbListenPort.Text = "";
            tbTargetIP.Text = "";
            tbTargetPort.Text = "";
            tbDescription.Text = "";
            cbProtocol.SelectedIndex = 0;
            chkFirewallRule.Checked = false;
        }

        private void CreateNewProfile()
        {
            string profileName = ShowInputDialog("Введите имя нового профиля:", "Создание профиля");
            
            if (!string.IsNullOrEmpty(profileName))
            {
                var newProfile = new RuleProfile
                {
                    Name = profileName,
                    Rules = GetCurrentRules()
                };
                
                settings.Profiles.Add(newProfile);
                SaveSettings();
                RefreshProfilesList();
                cbProfiles.SelectedItem = profileName;
                
                Log($"Создан новый профиль: {profileName}");
            }
        }

        private string ShowInputDialog(string text, string caption, string defaultValue = "")
        {
            Form prompt = new Form()
            {
                Width = 300,
                Height = 150,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = caption,
                StartPosition = FormStartPosition.CenterScreen
            };
            
            TextBox textBox = new TextBox() 
            { 
                Left = 10, 
                Top = 20, 
                Width = 260,
                Text = defaultValue
            };
            
            Button confirmation = new Button() { Text = "OK", Left = 110, Width = 80, Top = 70 };
            
            confirmation.Click += (sender, e) => { prompt.Close(); };
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.AcceptButton = confirmation;
            
            prompt.ShowDialog();
            return textBox.Text;
        }

        private List<PortRule> GetCurrentRules()
        {
            var rules = new List<PortRule>();
            string output = ExecuteNetshCommand("interface portproxy show all");
            
            bool skipHeaders = true;
            foreach (var line in output.Split('\n'))
            {
                if (string.IsNullOrWhiteSpace(line) || line.Contains("Прослушивать ipv4") || line.Contains("Адрес"))
                {
                    skipHeaders = false;
                    continue;
                }
                
                if (!skipHeaders && !line.Contains("---"))
                {
                    var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length >= 4)
                    {
                        rules.Add(new PortRule
                        {
                            Protocol = parts[0].Contains(":") ? PortProtocol.IPv6_to_IPv4 : PortProtocol.IPv4_to_IPv4,
                            ListenPort = int.Parse(parts[1]),
                            TargetIP = parts[2],
                            TargetPort = int.Parse(parts[3])
                        });
                    }
                }
            }
            
            return rules;
        }

        private void SaveCurrentProfile()
        {
            if (cbProfiles.SelectedItem == null) return;
            
            string profileName = cbProfiles.SelectedItem.ToString();
            var profile = settings.Profiles.FirstOrDefault(p => p.Name == profileName);
            
            if (profile != null)
            {
                profile.Rules = GetCurrentRules();
                profile.Modified = DateTime.Now;
                settings.LastActiveProfile = profileName;
                SaveSettings();
                
                Log($"Профиль '{profileName}' сохранен");
            }
        }

        private void DeleteProfile()
        {
            if (cbProfiles.SelectedItem == null) return;
            
            string profileName = cbProfiles.SelectedItem.ToString();
            
            if (MessageBox.Show($"Удалить профиль '{profileName}'?", "Подтверждение",
                MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                settings.Profiles.RemoveAll(p => p.Name == profileName);
                SaveSettings();
                RefreshProfilesList();
                
                Log($"Профиль '{profileName}' удален");
            }
        }

        private void LoadProfile(string profileName)
        {
            var profile = settings.Profiles.FirstOrDefault(p => p.Name == profileName);
            if (profile == null) return;

            try
            {
                // Очищаем текущие правила
                ExecuteNetshCommand("interface portproxy reset");

                // Применяем правила из профиля
                foreach (var rule in profile.Rules)
                {
                    string protocol = GetNetshProtocol(rule.Protocol.ToString());
                    ExecuteNetshCommand(
                        $"interface portproxy add {protocol} " +
                        $"listenport={rule.ListenPort} listenaddress=0.0.0.0 " +
                        $"connectport={rule.TargetPort} connectaddress={rule.TargetIP}");
                }

                // Обновляем UI
                RefreshRulesList();
                settings.LastActiveProfile = profileName;
                SaveSettings();

                Log($"Загружен профиль: {profileName}");
            }
            catch (Exception ex)
            {
                Log($"Ошибка загрузки профиля: {ex.Message}");
            }
        }

        private void ExportToJson()
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "JSON files (*.json)|*.json";
                dialog.FileName = $"portforward_profile_{DateTime.Now:yyyyMMdd}.json";
                
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var exportData = new
                        {
                            Profile = currentProfile,
                            Timestamp = DateTime.Now,
                            Version = Application.ProductVersion
                        };
                        
                        string json = JsonConvert.SerializeObject(exportData, Formatting.Indented);
                        File.WriteAllText(dialog.FileName, json);
                        
                        Log($"Профиль экспортирован в: {dialog.FileName}");
                        MessageBox.Show("Экспорт завершен успешно", "Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        Log($"Ошибка экспорта: {ex.Message}");
                        MessageBox.Show($"Ошибка экспорта: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ImportFromJson()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "JSON files (*.json)|*.json";
                
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string json = File.ReadAllText(dialog.FileName);
                        var importData = JsonConvert.DeserializeObject<RuleProfile>(json);
                        
                        if (importData != null)
                        {
                            var result = MessageBox.Show("Создать новый профиль на основе импортируемого?", "Импорт", 
                                MessageBoxButtons.YesNoCancel);
                            
                            if (result == DialogResult.Yes)
                            {
                                CreateNewProfileFromImport(importData);
                            }
                            else if (result == DialogResult.No)
                            {
                                ReplaceCurrentProfile(importData);
                            }
                            
                            Log($"Профиль импортирован из: {dialog.FileName}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"Ошибка импорта: {ex.Message}");
                        MessageBox.Show($"Ошибка импорта: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void CreateNewProfileFromImport(RuleProfile profile)
        {
            string profileName = ShowInputDialog(
                "Введите имя нового профиля:",
                "Создание профиля из импорта",
                profile.Name);
            
            if (!string.IsNullOrEmpty(profileName))
            {
                profile.Name = profileName;
                settings.Profiles.Add(profile);
                SaveSettings();
                RefreshProfilesList();
                cbProfiles.SelectedItem = profileName;
            }
        }

        private void ReplaceCurrentProfile(RuleProfile profile)
        {
            if (cbProfiles.SelectedItem == null) return;
            
            string currentProfileName = cbProfiles.SelectedItem.ToString();
            var existingProfile = settings.Profiles.FirstOrDefault(p => p.Name == currentProfileName);
            
            if (existingProfile != null)
            {
                existingProfile.Rules = profile.Rules;
                existingProfile.Modified = DateTime.Now;
                SaveSettings();
                LoadProfile(currentProfileName);
            }
        }

        private void CreateBackup()
        {
            try
            {
                string backupDir = Path.Combine(AppDataFolder, BackupFolder);
                string backupFile = Path.Combine(backupDir, $"backup_{DateTime.Now:yyyyMMdd_HHmmss}.json");
                
                var backupData = new
                {
                    Rules = GetCurrentRules(),
                    Timestamp = DateTime.Now,
                    Settings = settings
                };
                
                string json = JsonConvert.SerializeObject(backupData, Formatting.Indented);
                File.WriteAllText(backupFile, json);
                
                // Удаляем старые бэкапы
                var backupFiles = Directory.GetFiles(backupDir, "backup_*.json")
                                         .OrderByDescending(f => f)
                                         .Skip(settings.MaxBackups);
                
                foreach (var file in backupFiles)
                {
                    File.Delete(file);
                }
                
                Log($"Создана резервная копия: {backupFile}");
            }
            catch (Exception ex)
            {
                Log($"Ошибка создания резервной копии: {ex.Message}");
            }
        }

        private void CheckPorts()
        {
            try
            {
                dgvPortStatus.Rows.Clear();
                portStatuses.Clear();
                
                var currentRules = GetCurrentRules();
                
                foreach (var rule in currentRules)
                {
                    int pid = GetProcessIdFromPort(rule.ListenPort);
                    bool isAvailable = (pid == -1);
                    string serviceName = "Неизвестно";
                    string processInfo = "Неизвестно";
                    
                    if (!isAvailable)
                    {
                        try
                        {
                            var process = Process.GetProcessById(pid);
                            serviceName = process.ProcessName;
                            processInfo = $"{process.ProcessName} (PID: {process.Id})";
                        }
                        catch { }
                    }
                    
                    portStatuses.Add(new PortStatus
                    {
                        PortNumber = rule.ListenPort,
                        IsAvailable = isAvailable,
                        ServiceName = serviceName,
                        ProcessInfo = processInfo
                    });
                }
                
                // Отображаем результаты
                foreach (var status in portStatuses)
                {
                    int rowIndex = dgvPortStatus.Rows.Add(
                        status.PortNumber,
                        status.IsAvailable ? "Свободен" : "Занят",
                        status.ServiceName,
                        status.ProcessInfo);
                    
                    dgvPortStatus.Rows[rowIndex].DefaultCellStyle.BackColor = 
                        status.IsAvailable ? Color.LightGreen : Color.LightPink;
                }
                
                Log($"Проверено {portStatuses.Count} портов");
            }
            catch (Exception ex)
            {
                Log($"Ошибка проверки портов: {ex.Message}");
            }
        }

        [DllImport("iphlpapi.dll", SetLastError = true)]
        private static extern uint GetExtendedTcpTable(IntPtr pTcpTable, ref int dwOutBufLen, bool sort, int ipVersion, TCP_TABLE_CLASS tblClass, uint reserved);

        private enum TCP_TABLE_CLASS
        {
            TCP_TABLE_BASIC_LISTENER,
            TCP_TABLE_BASIC_CONNECTIONS,
            TCP_TABLE_BASIC_ALL,
            TCP_TABLE_OWNER_PID_LISTENER,
            TCP_TABLE_OWNER_PID_CONNECTIONS,
            TCP_TABLE_OWNER_PID_ALL,
            TCP_TABLE_OWNER_MODULE_LISTENER,
            TCP_TABLE_OWNER_MODULE_CONNECTIONS,
            TCP_TABLE_OWNER_MODULE_ALL
        }

        private static int GetProcessIdFromPort(int port)
        {
            IntPtr buffer = IntPtr.Zero;
            int bufferSize = 0;
            
            try
            {
                uint ret = GetExtendedTcpTable(IntPtr.Zero, ref bufferSize, false, 2, TCP_TABLE_CLASS.TCP_TABLE_OWNER_PID_ALL, 0);
                buffer = Marshal.AllocHGlobal(bufferSize);
                
                ret = GetExtendedTcpTable(buffer, ref bufferSize, false, 2, TCP_TABLE_CLASS.TCP_TABLE_OWNER_PID_ALL, 0);
                if (ret != 0) return -1;
                
                int entryCount = Marshal.ReadInt32(buffer);
                IntPtr rowPtr = (IntPtr)((long)buffer + 4);
                
                for (int i = 0; i < entryCount; i++)
                {
                    int localPort = Marshal.ReadInt32(rowPtr, 8) >> 16;
                    if (localPort == port)
                    {
                        return Marshal.ReadInt32(rowPtr, 24); // PID
                    }
                    rowPtr = (IntPtr)((long)rowPtr + Marshal.SizeOf(typeof(TcpRow)));
                }
            }
            finally
            {
                if (buffer != IntPtr.Zero)
                    Marshal.FreeHGlobal(buffer);
            }
            
            return -1;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct TcpRow
        {
            public int state;
            public int localAddr;
            public int localPort;
            public int remoteAddr;
            public int remotePort;
            public int owningPid;
        }

        private void ShowRulesVisualization()
        {
            graphPanel.Controls.Clear();
            
            var rules = GetCurrentRules();
            int yPos = 20;
            
            foreach (var rule in rules)
            {
                var rulePanel = new Panel
                {
                    Location = new Point(20, yPos),
                    Size = new Size(250, 60),
                    BorderStyle = BorderStyle.FixedSingle,
                    BackColor = Color.WhiteSmoke
                };
                
                var lblFrom = new Label
                {
                    Text = $"Порт: {rule.ListenPort}",
                    Location = new Point(10, 10),
                    AutoSize = true
                };
                
                var arrow = new Label
                {
                    Text = "→",
                    Font = new Font("Arial", 12, FontStyle.Bold),
                    Location = new Point(100, 10),
                    AutoSize = true
                };
                
                var lblTo = new Label
                {
                    Text = $"{rule.TargetIP}:{rule.TargetPort}",
                    Location = new Point(130, 10),
                    AutoSize = true
                };
                
                var lblProtocol = new Label
                {
                    Text = $"Протокол: {rule.Protocol}",
                    Location = new Point(10, 30),
                    AutoSize = true
                };
                
                rulePanel.Controls.Add(lblFrom);
                rulePanel.Controls.Add(arrow);
                rulePanel.Controls.Add(lblTo);
                rulePanel.Controls.Add(lblProtocol);
                graphPanel.Controls.Add(rulePanel);
                
                yPos += 70;
            }
            
            Log($"Отображена визуализация {rules.Count} правил");
        }

        private void ClearLogs()
        {
            string logFile = Path.Combine(AppDataFolder, LogFileName);
            
            if (File.Exists(logFile))
            {
                File.WriteAllText(logFile, string.Empty);
                txtLog.Clear();
                Log("Логи очищены");
            }
        }

        private void Log(string message)
        {
            string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
            
            // Вывод в текстовое поле
            txtLog.AppendText(logEntry + Environment.NewLine);
            txtLog.ScrollToCaret();
            
            // Запись в файл
            string logFile = Path.Combine(AppDataFolder, LogFileName);
            File.AppendAllText(logFile, logEntry + Environment.NewLine);
        }
    }
}