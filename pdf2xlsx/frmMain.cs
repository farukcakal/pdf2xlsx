using Aspose.Pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace pdf2xlsx
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            selectFile.Filter = "PDF Dosyaları (*.pdf)|*.pdf";
            DialogResult resultOpenFile = selectFile.ShowDialog();

            if(resultOpenFile == DialogResult.OK)
            {
                Document pdfDocument = new Document(selectFile.FileName);


                ExcelSaveOptions excelSave = new ExcelSaveOptions()
                {
                    MinimizeTheNumberOfWorksheets = true,
                    Format = ExcelSaveOptions.ExcelFormat.XLSX,
                    
                };

                saveFile.Filter = "Excel Dosyaları (*.xlsx)|*.xlsx*";
                DialogResult resultSaveFile = saveFile.ShowDialog();

                if(resultSaveFile == DialogResult.OK)
                {

                    pdfDocument.Save(saveFile.FileName + ".xlsx", excelSave);
                    //document.Save(saveFile.FileName + ".xlsx", SaveFormat.Excel, excelSaveOptions);
                    MessageBox.Show("Başarıyla dönüştürüldü.");
                }
                else
                {
                    MessageBox.Show("İşlem kullanıcı tarafından iptal edildi.");
                }
            }
            else
            {
                MessageBox.Show("İşlem kullanıcı tarafından iptal edildi.");
            }
        }
    }
}
