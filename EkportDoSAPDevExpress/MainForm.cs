using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace EkportDoSAPDevExpress
{
    public partial class MainForm : DevExpress.XtraEditors.XtraForm
    {
        public string ZazGUID = "";
        public string OperatorIdent = "";

        public MainForm(string GUID, string OpeIdent)
        {
            InitializeComponent();
            ZazGUID = GUID;
            OperatorIdent = OpeIdent;
            gridControl1.DataSource = nXP_EksportDoSAPTableAdapter.GetData(ZazGUID);
  
        }

        private void ResetData(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gridControl1.DataSource = nXP_EksportDoSAPTableAdapter.GetData(ZazGUID);
        }

        private void SaveToSap(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridControl1.MainView.DataRowCount != 0)
            {
                SaveFile.FileName += DateTime.Now.ToString().Replace(":", "_").Replace(" ", "_");
                if (SaveFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    
                    //Zapis do bazy
                    SqlConnection Conn = new SqlConnection(Properties.Settings.Default.AGRANA);
                    SqlCommand Usun = new SqlCommand();
                    Usun.CommandText = string.Format("DELETE FROM CDN.NXT_ListaWyslawnychDoSAP WHERE ZazGUID='{0}'", ZazGUID);
                    if (Conn.State == ConnectionState.Closed)
                    {
                        Conn.Open();
                    }
                    Usun.Connection = Conn;
                    Usun.ExecuteNonQuery();
                    for (int i = 0; i < gridControl1.MainView.DataRowCount; i++)
                    {
                        SqlCommand Cmd = new SqlCommand();
                        string sql = "INSERT INTO CDN.NXT_ListaWyslawnychDoSAP\n"
                           + "(\n"
                           + "    CDN.NXT_ListaWyslawnychDoSAP.CompanyCode,\n"
                           + "    CDN.NXT_ListaWyslawnychDoSAP.PlantCode,\n"
                           + "    CDN.NXT_ListaWyslawnychDoSAP.[Date],\n"
                           + "    CDN.NXT_ListaWyslawnychDoSAP.Account,\n"
                           + "    CDN.NXT_ListaWyslawnychDoSAP.CostCenter,\n"
                           + "    CDN.NXT_ListaWyslawnychDoSAP.InternalOrder,\n"
                           + "    CDN.NXT_ListaWyslawnychDoSAP.Currency,\n"
                           + "    CDN.NXT_ListaWyslawnychDoSAP.CreditDebit,\n"
                           + "    CDN.NXT_ListaWyslawnychDoSAP.GLindivator,\n"
                           + "    CDN.NXT_ListaWyslawnychDoSAP.Ammount,\n"
                           + "    CDN.NXT_ListaWyslawnychDoSAP.Reference,\n"
                           + "    CDN.NXT_ListaWyslawnychDoSAP.AccountingText,\n"
                           + "    CDN.NXT_ListaWyslawnychDoSAP.BookingDescription,\n"
                           + "    CDN.NXT_ListaWyslawnychDoSAP.ZazGUID,\n"
                           + "    CDN.NXT_ListaWyslawnychDoSAP.Operator,\n"
                           + "    CDN.NXT_ListaWyslawnychDoSAP.FileName,\n"
                           + "    CDN.NXT_ListaWyslawnychDoSAP.TStamp\n"
                           + ")\n"
                           + "VALUES\n"
                           + "(\n"
                           + "    '{0}', -- CompanyCode - varchar\n"
                           + "    '{1}', -- PlantCode - varchar\n"
                           + "    '{2}', -- Date - varchar\n"
                           + "    '{3}', -- Account - varchar\n"
                           + "    '{4}', -- CostCenter - varchar\n"
                           + "    '{5}', -- InternalOrder - varchar\n"
                           + "    '{6}', -- Currency - varchar\n"
                           + "    '{7}', -- CreditDebit - varchar\n"
                           + "    '{8}', -- GLindivator - varchar\n"
                           + "    REPLACE('{9}',',','.'), -- Ammount - decimal\n"
                           + "    '{10}', -- Reference - varchar\n"
                           + "    '{11}', -- AccountingText - varchar\n"
                           + "    '{12}', -- BookingDescription - varchar\n"
                           + "    '{13}', -- ZazGUID - varchar\n"
                           + "    '{14}', -- Operator - varchar\n"
                           + "    '{15}', --FileName -varchar\n"
                           + "    Current_timestamp -- TStamp - datetime\n"
                           + ")";
                        Cmd.CommandText = string.Format(sql,
                            ListyPlac.GetRowCellValue(i, ListyPlac.Columns[0]).ToString(),  //Company Code
                            ListyPlac.GetRowCellValue(i, ListyPlac.Columns[1]).ToString(),  //Plant Code
                            ListyPlac.GetRowCellValue(i, ListyPlac.Columns[2]).ToString(),  //Data
                            ListyPlac.GetRowCellValue(i, ListyPlac.Columns[3]).ToString(),  //Account
                            ListyPlac.GetRowCellValue(i, ListyPlac.Columns[4]).ToString(),  //CostCenter
                            ListyPlac.GetRowCellValue(i, ListyPlac.Columns[5]).ToString(),  //InternalOrder
                            ListyPlac.GetRowCellValue(i, ListyPlac.Columns[6]).ToString(),  //Currency
                            ListyPlac.GetRowCellValue(i, ListyPlac.Columns[7]).ToString(),  //CreditDebit
                            ListyPlac.GetRowCellValue(i, ListyPlac.Columns[8]).ToString(),  //GLindivator
                            ListyPlac.GetRowCellValue(i, ListyPlac.Columns[9]).ToString(),  //Ammount
                            ListyPlac.GetRowCellValue(i, ListyPlac.Columns[10]).ToString(), //Reference
                            ListyPlac.GetRowCellValue(i, ListyPlac.Columns[11]).ToString(), //AccountingText
                            ListyPlac.GetRowCellValue(i, ListyPlac.Columns[12]).ToString(), //BookingDescription
                            ZazGUID,
                            OperatorIdent,
                            SaveFile.FileName
                            );

                  
                        Cmd.Connection = Conn;
                        Cmd.ExecuteNonQuery();
                    }
                    Conn.Close();
                    gridControl2.DataSource = mM_EksportDoSAPPogrupowanyTableAdapter.GetData(ZazGUID);
                    gridControl2.ExportToCsv(SaveFile.FileName, new DevExpress.XtraPrinting.CsvExportOptionsEx
                    {
                        ExportType = DevExpress.Export.ExportType.WYSIWYG
                    });
                    DialogResult ToCloseOrNotToClose = MessageBox.Show("Poprawnie wyeksportowano plik" + Environment.NewLine + "Czy checsz zamknąć aplikację?", "Informacja", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    if (ToCloseOrNotToClose == DialogResult.Yes)
                    {
                        this.Close();
                    }
                  
                }
            }
            else
            {
                MessageBox.Show("Brak danych do eksportu", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void VisitWWW(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.netrix.com.pl");
        }

        private void SaveToXLS(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            SaveToXLSX.FileName += DateTime.Now.ToString().Replace(":", "_").Replace(" ", "_");
            if (SaveToXLSX.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                DevExpress.XtraPrinting.XlsxExportOptionsEx ExcelSettings = new DevExpress.XtraPrinting.XlsxExportOptionsEx
                {
                    ShowGridLines = false,
                    ShowColumnHeaders = DevExpress.Utils.DefaultBoolean.True,
                    ShowTotalSummaries = DevExpress.Utils.DefaultBoolean.False,
                    TextExportMode = DevExpress.XtraPrinting.TextExportMode.Value,
                    RawDataMode = true

                };

                gridControl1.ExportToXlsx(SaveToXLSX.FileName, ExcelSettings);
            }
            DialogResult ToCloseOrNotToClose = MessageBox.Show("Poprawnie wyeksportowano plik" + Environment.NewLine + "Czy checsz zamknąć aplikację?", "Informacja", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (ToCloseOrNotToClose == DialogResult.Yes)
            {
                this.Close();
            }
        }

        private void pivotGridControl1_Click(object sender, EventArgs e)
        {
        }

        private void barStaticItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            System.Diagnostics.Process.Start("http://prospeo.com.pl/");
        }
    }
}