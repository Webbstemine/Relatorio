using Relatorio;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RelatorioWinForms
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Relatorio.Relatorio relatorio = new Relatorio.Relatorio(new ConfiguracaoRelatorio()
            {
                MargemSuperior = 5
            });
            relatorio.AbrirArquivo();
            relatorio.AdicionarParagrafo("Teste");
            relatorio.CriarTabela();
            for (int i = 0; i < 100; i++)
            {
                relatorio.CriarCelula($"Coluna 1 - {i}");
                relatorio.CriarCelula($"Coluna 2 - {i}");
            }
            relatorio.FecharArquivo();
        }
    }
}
