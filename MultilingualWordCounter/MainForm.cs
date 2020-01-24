// <copyright file="MainForm.cs" company="PublicDomain.com">
//     CC0 1.0 Universal (CC0 1.0) - Public Domain Dedication
//     https://creativecommons.org/publicdomain/zero/1.0/legalcode
// </copyright>

namespace MultilingualWordCounter
{
    // Directives
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Drawing;
    using System.IO;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Text.RegularExpressions;
    using System.Windows.Forms;
    using System.Xml.Serialization;
    using Microsoft.Win32;

    /// <summary>
    /// Main form.
    /// </summary>
    public partial class MainForm : Form
    {
        /// <summary>
        /// The mod control.
        /// </summary>
        private const int MODCONTROL = 0x0002; // Changed from MOD_CONTROL for StyleCop

        /// <summary>
        /// The mod shift.
        /// </summary>
        private const int MODSHIFT = 0x0004; // Changed from MOD_SHIFT for StyleCop

        /// <summary>
        /// The wm hotkey.
        /// </summary>
        private const int WMHOTKEY = 0x0312; // Changed from  for StyleCop

        /// <summary>
        /// The last clipboard text.
        /// </summary>
        private string lastClipboardText = string.Empty;

        /// <summary>
        /// The settings data.
        /// </summary>
        private SettingsData settingsData = new SettingsData();

        /// <summary>
        /// The assembly version.
        /// </summary>
        private Version assemblyVersion = Assembly.GetExecutingAssembly().GetName().Version;

        /// <summary>
        /// The semantic version.
        /// </summary>
        private string semanticVersion = string.Empty;

        /// <summary>
        /// The associated icon.
        /// </summary>
        private Icon associatedIcon = null;

        /// <summary>
        /// The friendly name of the program.
        /// </summary>
        private string friendlyName = "Multilingual Word Counter";

        /// <summary>
        /// The language speed wpm dictionary.
        /// </summary>
        private Dictionary<string, Dictionary<string, int>> languageSpeedWpmDictionary = new Dictionary<string, Dictionary<string, int>>();

        /// <summary>
        /// The message form.
        /// </summary>
        private MessageForm messageForm = new MessageForm();

        /// <summary>
        /// Initializes a new instance of the <see cref="T:MultilingualWordCounter.MainForm"/> class.
        /// </summary>
        public MainForm()
        {
            // The InitializeComponent() call is required for Windows Forms designer support.
            this.InitializeComponent();

            // Set message form icon
            this.messageForm.Icon = this.Icon;

            // Set notify icon
            this.mainNotifyIcon.Icon = this.Icon;

            // Set semantic version
            this.semanticVersion = this.assemblyVersion.Major + "." + this.assemblyVersion.Minor + "." + this.assemblyVersion.Build;

            /* Process languages */

            // Set languages file path
            var languageFilePath = "Languages.txt";

            // Check for languages file
            if (!File.Exists(languageFilePath))
            {
                // Inform user
                MessageBox.Show($"Missing \"{languageFilePath}\" file!", "Required", MessageBoxButtons.OK, MessageBoxIcon.Error);

                // Exit program
                this.Close();

                // Halt flow
                return;
            }

            // Load languages into list
            var languageList = new List<string>(File.ReadAllLines(languageFilePath));

            // Add to combo boxes
            foreach (var languageLine in languageList)
            {
                // Set language info
                List<string> languageInfo = new List<string>(languageLine.Split(new char[] { ',' }, 4)); // language, slow, average, fast

                // Check information is complete
                if (languageInfo.Count < 4)
                {
                    // Skip incomplete language info
                    continue;
                }

                // Set language name
                string languageName = languageInfo[0];

                /* Language */

                // Add to native combo box
                this.nativeComboBox.Items.Add(languageName);

                // Add to foreign combo box
                this.foreignComboBox.Items.Add(languageName);

                /* Speed */

                // Add language to words per minute dictionary
                this.languageSpeedWpmDictionary.Add(languageName, new Dictionary<string, int>());

                // Slow
                this.languageSpeedWpmDictionary[languageName].Add("Slow", int.Parse(languageInfo[1]));

                // Average
                this.languageSpeedWpmDictionary[languageName].Add("Average", int.Parse(languageInfo[2]));

                // Fast
                this.languageSpeedWpmDictionary[languageName].Add("Fast", int.Parse(languageInfo[3]));
            }

            /* Process settings */

            // Set settings file path
            var settingsFilePath = "SettingsData.txt";

            // Check for settings data file
            if (!File.Exists(settingsFilePath))
            {
                // Not present, assume first run and create it
                this.SaveSettingsData();

                // Inform user
                MessageBox.Show($"Created \"{settingsFilePath}\" file.{Environment.NewLine}Program icon will appear on system tray.", "First run", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            // Populate settings data
            this.settingsData = this.LoadSettingsData();

            // Set registry entry based on settings data
            this.ProcessRunAtStartupRegistry();

            // Set gui values
            this.SetGuiValuesFromSettingsData();

            // Set run at startup tool strip menu item check state
            this.runAtStartupToolStripMenuItem.Checked = this.settingsData.RunAtStartup;
        }

        /// <summary>
        /// Windows procedure.
        /// </summary>
        /// <param name="m">The message</param>
        protected override void WndProc(ref Message m)
        {
            // Chedk for hotkey message
            if (m.Msg == WMHOTKEY)
            {
                // Act on hotkey
                switch ((int)m.WParam)
                {
                    // CTRL+F6
                    case 1:
                        // Show message
                        this.messageForm.ShowMessage($"{this.CountWords(Clipboard.GetText())} copied words", "Clipboard word count");

                        // Halt flow
                        break;

                    // CTRL+SHIFT+F6
                    case 2:
                        // Show program window
                        this.RestoreFromSystemTray();

                        // Halt flow
                        break;

                    // CTRL+F7
                    case 3:
                        // Show message
                        this.messageForm.ShowMessage(this.TimeSpanToHumanReadable(this.GetSpeechTimeSpan(Clipboard.GetText(), this.nativeComboBox.SelectedItem.ToString(), this.languageSpeedWpmDictionary[this.nativeComboBox.SelectedItem.ToString()][this.speedComboBox.SelectedItem.ToString()])) + $" ({this.speedComboBox.SelectedItem.ToString()})", "Native speech time");

                        // Halt flow
                        break;

                    // CTRL+SHIFT+F7
                    case 4:
                        // Show message
                        this.messageForm.ShowMessage(this.TimeSpanToHumanReadable(this.GetSpeechTimeSpan(Clipboard.GetText(), this.foreignComboBox.SelectedItem.ToString(), this.languageSpeedWpmDictionary[this.foreignComboBox.SelectedItem.ToString()][this.speedComboBox.SelectedItem.ToString()])) + $" ({this.speedComboBox.SelectedItem.ToString()})", "Foreign speech time");

                        // Halt flow
                        break;

                    // CTRL+F8
                    case 5:
                        // Check for minimum speed
                        if (this.speedComboBox.SelectedIndex > 0)
                        {
                            // Lower speed
                            this.speedComboBox.SelectedIndex--;
                        }

                        // Halt flow
                        break;

                    // CTRL+SHIFT+F8
                    case 6:
                        // Check for maximum speed
                        if (this.speedComboBox.SelectedIndex < this.speedComboBox.Items.Count - 1)
                        {
                            // Increment speed
                            this.speedComboBox.SelectedIndex++;
                        }

                        // Halt flow
                        break;
                }
            }

            // Forward message
            base.WndProc(ref m);
        }

        /// <summary>
        /// Registers the hot key.
        /// </summary>
        /// <returns><c>true</c>, if hot key was registered, <c>false</c> otherwise.</returns>
        /// <param name="handle">The window handle.</param>
        /// <param name="id">The identifier.</param>
        /// <param name="modifiers">The modifiers.</param>
        /// <param name="vk">The virtual key.</param>
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr handle, int id, int modifiers, int vk);

        /// <summary>
        /// Unregisters the hot key.
        /// </summary>
        /// <returns><c>true</c>, if the hot key was unregistered, <c>false</c> otherwise.</returns>
        /// <param name="handle">The window handle.</param>
        /// <param name="id">The identifier.</param>
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr handle, int id);

        /// <summary>
        /// Counts the words.
        /// </summary>
        /// <returns>The counted words.</returns>
        /// <param name="text">The input text.</param>
        private int CountWords(string text)
        {
            // Return the word count
            return Regex.Matches(text, @"[A-Za-z0-9]+").Count;
        }

        /// <summary>
        /// Gets the speech time span.
        /// </summary>
        /// <returns>The speech time span.</returns>
        /// <param name="text">The text.</param>
        /// <param name="language">The language.</param>
        /// <param name="wordsPerMinute">Words per minute.</param>
        private TimeSpan GetSpeechTimeSpan(string text, string language, int wordsPerMinute)
        {
            // Return the timespan
            return TimeSpan.FromSeconds((this.CountWords(text) * 60) / wordsPerMinute);
        }

        /// <summary>
        /// Gets the human readable representation of passed time span.
        /// </summary>
        /// <returns>The human readable representation of passed time span..</returns>
        /// <param name="timeSpan">Time span.</param>
        private string TimeSpanToHumanReadable(TimeSpan timeSpan)
        {
            // Pieceslist
            var pieceList = new List<string>();

            // Check for days
            if (timeSpan.Days > 0)
            {
                // Add days
                pieceList.Add($"{timeSpan.Days} day{(timeSpan.Days > 1 ? "s" : string.Empty)}");
            }

            // Check for hours
            if (timeSpan.Hours > 0)
            {
                // Add hours
                pieceList.Add($"{timeSpan.Hours} hour{(timeSpan.Hours > 1 ? "s" : string.Empty)}");
            }

            // Check for minutes
            if (timeSpan.Minutes > 0)
            {
                // Add minutes
                pieceList.Add($"{timeSpan.Minutes} minute{(timeSpan.Minutes > 1 ? "s" : string.Empty)}");
            }

            // Check for seconds
            if (timeSpan.Seconds > 0)
            {
                // Add seconds
                pieceList.Add($"{timeSpan.Seconds} second{(timeSpan.Seconds > 1 ? "s" : string.Empty)}");
            }

            return string.Join(", ", pieceList);
        }

        /// <summary>
        /// Sets the GUI values from settings data.
        /// </summary>
        private void SetGuiValuesFromSettingsData()
        {
            // Set native language
            this.nativeComboBox.SelectedItem = this.settingsData.NativeLanguage;

            // Set foreign language
            this.foreignComboBox.SelectedItem = this.settingsData.ForeignLanguage;

            // Set speed 
            this.speedComboBox.SelectedItem = this.settingsData.SpeechSpeed;
        }

        /// <summary>
        /// Processes the run at startup registry action.
        /// </summary>
        private void ProcessRunAtStartupRegistry()
        {
            // Open registry key
            using (RegistryKey registryKey = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true))
            {
                // Check for run at startup in settings data
                if (this.settingsData.RunAtStartup)
                {
                    // Check for app value
                    if (registryKey.GetValue(Application.ProductName) == null)
                    {
                        // Add app value
                        registryKey.SetValue(Application.ProductName, $"\"{Application.ExecutablePath}\" /autostart");
                    }
                }
                else
                {
                    // Erase app value
                    registryKey.DeleteValue(Application.ProductName, false);
                }
            }
        }

        /// <summary>
        /// Handles the headquarters at patreon.com tool strip menu item click event.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void OnHeadquartersPatreoncomToolStripMenuItemClick(object sender, EventArgs e)
        {
            // Open Patreon headquarters
            Process.Start("https://www.patreon.com/publicdomain");
        }

        /// <summary>
        /// Handles the source code at github.com tool strip menu item click event.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void OnSourceCodeGithubcomToolStripMenuItemClick(object sender, EventArgs e)
        {
            // Open GitHub
            Process.Start("https://github.com/publicdomain");
        }

        /// <summary>
        /// Handles the original thread donation codercom tool strip menu item click event.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void OnOriginalThreadDonationCodercomToolStripMenuItemClick(object sender, EventArgs e)
        {
            // Open original thread @ DonationCoder
            Process.Start("https://www.donationcoder.com/forum/index.php?topic=47421.0");
        }

        /// <summary>
        /// Handles the about tool strip menu item click event.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void OnAboutToolStripMenuItemClick(object sender, EventArgs e)
        {
            // TODO Add code
        }

        /// <summary>
        /// Handles the new tool strip menu item click event.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void OnNewToolStripMenuItemClick(object sender, EventArgs e)
        {
            // Recreate initial settings data
            this.settingsData = new SettingsData();

            // Save settings data to disk
            this.SaveSettingsData();

            // Set default GUI values
            this.SetGuiValuesFromSettingsData();
        }

        /// <summary>
        /// Handles the exit tool strip menu item click event.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void OnExitToolStripMenuItemClick(object sender, EventArgs e)
        {
            // Close application
            this.Close();
        }

        /// <summary>
        /// Handles the options tool strip menu item drop down item clicked event.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void OnOptionsToolStripMenuItemDropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            // Toggle run at startup check box
            this.runAtStartupToolStripMenuItem.Checked = !this.runAtStartupToolStripMenuItem.Checked;
        }

        /// <summary>
        /// Handles the show tool strip menu item click event.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void OnShowToolStripMenuItemClick(object sender, EventArgs e)
        {
            // Restore window 
            this.RestoreFromSystemTray();
        }

        /// <summary>
        /// Handles the main form resize event.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void OnMainFormResize(object sender, EventArgs e)
        {
            // Check for minimized state
            if (this.WindowState == FormWindowState.Minimized)
            {
                // Send to the system tray
                this.SendToSystemTray();
            }
        }

        /// <summary>
        /// Sends the program to the system tray.
        /// </summary>
        private void SendToSystemTray()
        {
            // Hide main form
            this.Hide();

            // Show notify icon 
            this.mainNotifyIcon.Visible = true;
        }

        /// <summary>
        /// Restores the window back from system tray to the foreground.
        /// </summary>
        private void RestoreFromSystemTray()
        {
            // Make form visible again
            this.Show();

            // Return window back to normal
            this.WindowState = FormWindowState.Normal;

            // Hide system tray icon
            this.mainNotifyIcon.Visible = false;
        }

        /// <summary>
        /// Saves the settings data.
        /// </summary>
        private void SaveSettingsData()
        {
            // Use stream writer
            using (StreamWriter streamWriter = new StreamWriter("SettingsData.txt", false))
            {
                // Set xml serialzer
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(SettingsData));

                // Serialize settings data
                xmlSerializer.Serialize(streamWriter, this.settingsData);
            }
        }

        /// <summary>
        /// Loads the settings data.
        /// </summary>
        /// <returns>The settings data.</returns>ing
        private SettingsData LoadSettingsData()
        {
            // Use file stream
            using (FileStream fileStream = File.OpenRead("SettingsData.txt"))
            {
                // Set xml serialzer
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(SettingsData));

                // Return populated settings data
                return xmlSerializer.Deserialize(fileStream) as SettingsData;
            }
        }

        /// <summary>
        /// Handles the minimize tool strip menu item click event.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void OnMinimizeToolStripMenuItemClick(object sender, EventArgs e)
        {
            // Minimize program window
            this.WindowState = FormWindowState.Minimized;
        }

        /// <summary>
        /// Handles the main form load event.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void OnMainFormLoad(object sender, EventArgs e)
        {
            // Register hotkeys
            RegisterHotKey(this.Handle, 1, MODCONTROL, (int)Keys.F6); // Count words
            RegisterHotKey(this.Handle, 2, MODCONTROL + MODSHIFT, (int)Keys.F6); // Show program window
            RegisterHotKey(this.Handle, 3, MODCONTROL, (int)Keys.F7); // Native speech time
            RegisterHotKey(this.Handle, 4, MODCONTROL + MODSHIFT, (int)Keys.F7); // Foreign speech time
            RegisterHotKey(this.Handle, 5, MODCONTROL, (int)Keys.F8); // Decrease speed
            RegisterHotKey(this.Handle, 6, MODCONTROL + MODSHIFT, (int)Keys.F8); // Increase speed
        }

        /// <summary>
        /// Handles the main form form closing event.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void OnMainFormFormClosing(object sender, FormClosingEventArgs e)
        {
            // Set run at startup
            this.settingsData.RunAtStartup = this.runAtStartupToolStripMenuItem.Checked;

            // Set native language
            this.settingsData.NativeLanguage = this.nativeComboBox.SelectedItem.ToString();

            // Set foreign language
            this.settingsData.ForeignLanguage = this.foreignComboBox.SelectedItem.ToString();

            // Set speech speed
            this.settingsData.SpeechSpeed = this.speedComboBox.SelectedItem.ToString();

            // Save settings data to disk
            this.SaveSettingsData();

            // Unregister hotkeys
            UnregisterHotKey(this.Handle, 1);
            UnregisterHotKey(this.Handle, 2);
            UnregisterHotKey(this.Handle, 3);
            UnregisterHotKey(this.Handle, 4);
            UnregisterHotKey(this.Handle, 5);
            UnregisterHotKey(this.Handle, 6);
        }

        /// <summary>
        /// Handles the main notify icon mouse click event.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Mouse event arguments.</param>
        private void OnMainNotifyIconMouseClick(object sender, MouseEventArgs e)
        {
            // Check for left click
            if (e.Button == MouseButtons.Left)
            {
                // Restore window 
                this.RestoreFromSystemTray();
            }
        }

        /// <summary>
        /// Handles the main form shown event.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Mouse event arguments.</param>
        private void OnMainFormShown(object sender, EventArgs e)
        {
            // Minimize program window
            this.WindowState = FormWindowState.Minimized;
        }
    }
}
