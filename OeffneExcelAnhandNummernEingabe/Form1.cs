using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OeffneExcelAnhandNummernEingabe
{
    public partial class Form1 : Form
    {
        private string excelDir;

        public Form1()
        {
            this.InitializeComponent();
            if (!this.ReadConfig())
            {
                //exit!
                Environment.Exit(1);
            }

            //fenster immer im vordergrund:
            this.TopMost = true;
        }

        private bool ReadConfig()
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            if (config.AppSettings.Settings.Count == 0)
            {
                //MessageBox.Show(@"Das Verzeichnis ExcelDir (in dem sich alle Dateien befinden) muss noch angegeben werden!" + "\r\n" +
                //  @"app.config, Parameter: ExcelDir",
                //  @"fehlende Konfiguration1", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.excelDir = Application.StartupPath;
                return true;
            }
            if (
                config.AppSettings.Settings.Count == 1
                && config.AppSettings.Settings["ExcelDir"]?.Value == ""
            )
            {
                MessageBox.Show(
                    @"Das Verzeichnis ExcelDir (in dem sich alle Dateien befinden) muss noch angegeben werden!"
                        + "\r\n"
                        + @"app.config, Parameter: ExcelDir",
                    @"fehlende Konfiguration2",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation
                );
                return false;
            }
            this.excelDir =
                config.AppSettings.Settings["ExcelDir"]?.Value ?? Application.StartupPath;

            if (!Directory.Exists(this.excelDir))
            {
                MessageBox.Show(
                    $"Das angegebene Verzeichnis ExcelDir (in dem sich alle Dateien befinden) scheint nicht zu existieren. Wert aus der app.config: {this.excelDir}"
                        + "\r\n"
                        + @"app.config, Parameter: ExcelDir",
                    @"falsche Konfiguration3",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation
                );
                return false;
            }
            return true;
        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.OpenExcel(this.textBox1.Text);
            }
        }

        private void OpenExcel(string text)
        {
            var file = Path.Combine(this.excelDir, text + @".xlsx");
            //datei vorhanden?
            if (File.Exists(file))
            {
                WriteLog($"OpenExcel({file})");

                this.label1.Text = "";

                //start!
                System.Diagnostics.Process.Start(file);
            }
            else
            {
                this.label1.Text = $"{text}.xlsx konnte nicht gefunden werden!";
                WriteLog($"OpenExcel: {text}.xlsx konnte nicht gefunden werden!");
            }
            //falls ja, oeffnen

            //text selektieren, focus setzen
            this.textBox1.Text = "";
            this.textBox1.Focus();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.OpenExcel(this.textBox1.Text);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //try to save settings if not saved yet..
            this.SaveConfig();
        }

        private void SaveConfig()
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            bool save =
                config.AppSettings.Settings.Count == 0
                || config.AppSettings.Settings["ExcelDir"]?.Value != this.excelDir;
            if (save)
            {
                var keyValueConfigurationElement = config.AppSettings.Settings["ExcelDir"];
                if (keyValueConfigurationElement != null)
                    keyValueConfigurationElement.Value = this.excelDir;
            }
        }

        private void textBox1_MouseEnter(object sender, EventArgs e)
        {
            if (this.textBox1.Text.Length >= 0)
            {
                this.textBox1.SelectAll();
            }
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            if (this.textBox1.Text.Length >= 0)
            {
                this.textBox1.SelectAll();
            }
        }

        private void WriteLog(string logEntry)
        {
            var finalLogEntry = $"{DateTime.Now.ToString()}\t{logEntry}{Environment.NewLine}";
            File.AppendAllText(Path.Combine(this.excelDir, "8999.txt"), finalLogEntry);
        }

        private void Form1_Enter(object sender, EventArgs e)
        {
            if (this.textBox1.Text.Length >= 0)
            {
                this.textBox1.SelectAll();
            }
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            if (this.textBox1.Text.Length >= 0)
            {
                this.textBox1.SelectAll();
            }
        }
    }
}
