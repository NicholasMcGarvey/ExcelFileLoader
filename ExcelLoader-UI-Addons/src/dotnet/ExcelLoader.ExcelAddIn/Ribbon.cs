
using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace ExcelLoader.ExcelAddIn
{
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;
        public string GetCustomUI(string ribbonID) => GetRibbonXml();
        public void OnLoad(Office.IRibbonUI ribbonUI) => _ribbon = ribbonUI;

        public void OnSelect(Office.IRibbonControl c) => SelectFile(false);
        public void OnPreview(Office.IRibbonControl c) => SelectFile(true);
        public void OnDryRun(Office.IRibbonControl c) => RunLoader(true);
        public void OnLoadClick(Office.IRibbonControl c) => RunLoader(false);
        public void OnErrors(Office.IRibbonControl c) => MessageBox.Show("SELECT * FROM repo.LoadRunError ORDER BY LoadRunErrorId DESC", "ExcelLoader");
        public void OnSettings(Office.IRibbonControl c) { using var dlg = new SettingsForm(); dlg.ShowDialog(); }

        private static string _lastFile;
        private static string _connection;
        private static string _fileGroup = "COL";
        private static string _entity = "LenderA";
        private static string _loadType = "daily";
        private static string _cliPath = Environment.GetEnvironmentVariable("EXCELLOADER_CLI") ?? "";

        private void SelectFile(bool preview)
        {
            using var ofd = new OpenFileDialog { Filter = "Excel files|*.xlsx;*.xlsm|All files|*.*" };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                _lastFile = ofd.FileName;
                if (preview) MessageBox.Show($"Selected: {_lastFile}\n(Preview UI TBD)", "ExcelLoader");
            }
        }

        private void RunLoader(bool dryRun)
        {
            EnsureSettings();
            if (string.IsNullOrWhiteSpace(_lastFile)) { MessageBox.Show("Select a file first."); return; }
            if (!File.Exists(_cliPath)) { MessageBox.Show("CLI not found. Set EXCELLOADER_CLI or use Settings."); return; }

            var args = $"\"{_cliPath}\" --connection \"{_connection}\" --file \"{_lastFile}\" --fileGroup {_fileGroup} --entity {_entity} --loadType {_loadType}";
            if (dryRun) args += " --dryRun";
            var psi = new ProcessStartInfo("dotnet", args) { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true };
            var p = Process.Start(psi);
            var output = p.StandardOutput.ReadToEnd();
            var error = p.StandardError.ReadToEnd();
            p.WaitForExit();
            MessageBox.Show(string.IsNullOrWhiteSpace(error) ? output : error, dryRun ? "[Dry Run]" : "[Load]");
        }

        private void EnsureSettings()
        {
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelLoader", "addin.json");
            if (File.Exists(path))
            {
                var s = System.Text.Json.JsonSerializer.Deserialize<AddinSettings>(File.ReadAllText(path));
                _connection = s.Connection;
                _fileGroup = string.IsNullOrWhiteSpace(s.FileGroup) ? _fileGroup : s.FileGroup;
                _entity = string.IsNullOrWhiteSpace(s.Entity) ? _entity : s.Entity;
                _loadType = string.IsNullOrWhiteSpace(s.LoadType) ? _loadType : s.LoadType;
                _cliPath = string.IsNullOrWhiteSpace(s.CliPath) ? _cliPath : s.CliPath;
            }
        }

        private static string GetRibbonXml()
        {
            using var s = typeof(Ribbon).Assembly.GetManifestResourceStream("ExcelLoader.ExcelAddIn.Ribbon.xml");
            using var r = new StreamReader(s);
            return r.ReadToEnd();
        }

        private class AddinSettings
        {
            public string Connection { get; set; }
            public string FileGroup { get; set; }
            public string Entity { get; set; }
            public string LoadType { get; set; }
            public string CliPath { get; set; }
        }
    }
}
