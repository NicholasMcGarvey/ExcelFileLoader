
using System;
using System.IO;
using System.Linq;
using System.Windows;
using Dapper;
using Microsoft.Data.SqlClient;
using System.Diagnostics;

namespace ExcelLoader.Admin.Wpf
{
    public partial class MainWindow : Window
    {
        private string _conn;

        public MainWindow()
        {
            InitializeComponent();
            _conn = LoadConn();
            txtConn.Text = _conn;
            RefreshAll();
        }

        private string LoadConn()
        {
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelLoader", "admin.json");
            if (File.Exists(path)) return System.Text.Json.JsonSerializer.Deserialize<ConnObj>(File.ReadAllText(path)).Connection;
            return "";
        }

        private void SaveConn_Click(object sender, RoutedEventArgs e)
        {
            _conn = txtConn.Text;
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelLoader");
            Directory.CreateDirectory(dir);
            File.WriteAllText(Path.Combine(dir, "admin.json"),
                System.Text.Json.JsonSerializer.Serialize(new ConnObj { Connection = _conn }));
            MessageBox.Show("Saved.");
            RefreshAll();
        }

        private async void TestConn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await using var c = new SqlConnection(_conn);
                await c.OpenAsync();
                var ver = await c.ExecuteScalarAsync<string>("SELECT @@VERSION");
                MessageBox.Show("Connected.\n\n" + ver);
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        private async void RefreshAll()
        {
            try
            {
                await using var c = new SqlConnection(_conn);
                var fgs = await c.QueryAsync("SELECT * FROM repo.FileGroup");
                var ents = await c.QueryAsync("SELECT * FROM repo.Entity");
                var fs = await c.QueryAsync("SELECT * FROM repo.FileSpec");
                var fm = await c.QueryAsync("SELECT TOP 500 * FROM repo.FieldMap ORDER BY FieldMapId DESC");

                gridFileGroups.ItemsSource = fgs.ToList();
                gridEntities.ItemsSource = ents.ToList();
                gridFileSpecs.ItemsSource = fs.ToList();
                gridFieldMaps.ItemsSource = fm.ToList();
            }
            catch { }
        }

        private async void DryRun_Click(object sender, RoutedEventArgs e) => await RunCli(true);
        private async void Load_Click(object sender, RoutedEventArgs e) => await RunCli(false);

        private async System.Threading.Tasks.Task RunCli(bool dry)
        {
            txtOutput.Clear();
            var cli = Environment.GetEnvironmentVariable("EXCELLOADER_CLI");
            if (string.IsNullOrWhiteSpace(cli) || !File.Exists(cli))
            {
                txtOutput.Text = "Set EXCELLOADER_CLI to the built CLI DLL path.";
                return;
            }
            var args = $"\"{cli}\" --connection \"{_conn}\" --file \"{txtPath.Text}\" --fileGroup {txtFG.Text} --entity {txtEnt.Text} --loadType {txtLT.Text}";
            if (dry) args += " --dryRun";
            var psi = new ProcessStartInfo("dotnet", args) { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true };
            var p = Process.Start(psi);
            var std = await p.StandardOutput.ReadToEndAsync();
            var err = await p.StandardError.ReadToEndAsync();
            p.WaitForExit();
            txtOutput.Text = string.IsNullOrWhiteSpace(err) ? std : err;
        }

        private class ConnObj { public string Connection { get; set; } }
    }
}
