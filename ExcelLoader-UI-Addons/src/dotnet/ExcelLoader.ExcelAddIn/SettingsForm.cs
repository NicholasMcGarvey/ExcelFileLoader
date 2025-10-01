
using System;
using System.IO;
using System.Windows.Forms;

namespace ExcelLoader.ExcelAddIn
{
    public class SettingsForm : Form
    {
        TextBox txtConn = new TextBox(){ Width = 420 };
        TextBox txtFG = new TextBox(){ Width = 120, Text = "COL" };
        TextBox txtEnt = new TextBox(){ Width = 120, Text = "LenderA" };
        TextBox txtLT = new TextBox(){ Width = 120, Text = "daily" };
        TextBox txtCli = new TextBox(){ Width = 420 };
        Button btnBrowse = new Button(){ Text = "Browse..." };
        Button btnSave = new Button(){ Text = "Save", Width = 100 };

        public SettingsForm()
        {
            Text = "ExcelLoader Settings"; Width = 600; Height = 260;
            var table = new TableLayoutPanel(){ Dock = DockStyle.Fill, ColumnCount = 3 };

            table.Controls.Add(new Label(){ Text="Connection"}, 0, 0); table.Controls.Add(txtConn, 1,0);
            table.Controls.Add(new Label(){ Text="FileGroup"}, 0, 1); table.Controls.Add(txtFG, 1,1);
            table.Controls.Add(new Label(){ Text="Entity"},    0, 2); table.Controls.Add(txtEnt, 1,2);
            table.Controls.Add(new Label(){ Text="LoadType"},  0, 3); table.Controls.Add(txtLT, 1,3);
            table.Controls.Add(new Label(){ Text="CLI (dll)"}, 0, 4); table.Controls.Add(txtCli, 1,4); table.Controls.Add(btnBrowse, 2,4);
            table.Controls.Add(btnSave, 1, 5);
            Controls.Add(table);

            btnBrowse.Click += (s,e)=> { using var ofd = new OpenFileDialog(){ Filter = "DLL|*.dll" }; if (ofd.ShowDialog()== DialogResult.OK) txtCli.Text = ofd.FileName; };
            btnSave.Click += (s,e)=> Save();
        }

        private void Save() {
            var obj = new { Connection = txtConn.Text, FileGroup = txtFG.Text, Entity = txtEnt.Text, LoadType = txtLT.Text, CliPath = txtCli.Text };
            var json = System.Text.Json.JsonSerializer.Serialize(obj, new System.Text.Json.JsonSerializerOptions{ WriteIndented=true });
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelLoader");
            Directory.CreateDirectory(dir);
            File.WriteAllText(Path.Combine(dir, "addin.json"), json);
            MessageBox.Show("Saved."); Close();
        }
    }
}
