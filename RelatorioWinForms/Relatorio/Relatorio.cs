using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Diagnostics;
using System.IO;

namespace Relatorio
{
    public class Relatorio
    {
        private ConfiguracaoRelatorio Configuracao;

        private Document Pdf;
        private FileStream Arquivo;
        private PdfWriter Writer;
        private PdfPTable Tabela;

        public Relatorio(ConfiguracaoRelatorio Configuracao)
        {
            this.Configuracao = Configuracao;
            Tabela = new PdfPTable(Configuracao.NumeroColunas);
            Tabela.DefaultCell.BorderWidth = 0;
            Tabela.WidthPercentage = 100;
        }

        public void AbrirArquivo()
        {
            if (Configuracao == null)
                throw new Exception("Configuracao invalida!");

            //Criando o documento
            Pdf = new Document(PageSize.A4, Configuracao.getEsquerda(), Configuracao.getDireita(),
                Configuracao.getSuperior(), Configuracao.getInferior());
            Arquivo = new FileStream(Configuracao.CaminhoArquivo, FileMode.Create);
            Writer = PdfWriter.GetInstance(Pdf, Arquivo);
            Pdf.Open();
        }

        public void AdicionarParagrafo(string Texto, int AlinhamentoHorizontal = Element.ALIGN_CENTER, float TamanhoFonte = 32,
            int Estilo = Font.NORMAL, BaseColor CorTexto = null)
        {
            if (CorTexto == null)
                CorTexto = BaseColor.Black;
            Font FonteParagrafo = new Font(Configuracao.FonteBase, TamanhoFonte, Estilo, CorTexto);
            Paragraph Paragrafo = new Paragraph(Texto + "\n\n", FonteParagrafo);
            Paragrafo.Alignment = AlinhamentoHorizontal;
            Pdf.Add(Paragrafo);
        }

        public void CriarTabela()
        {
            //Adicionar os headers
            foreach (string Header in Configuracao.TitulosColuna)
            {
                CriarCelula(Header);
            }
        }

        public void CriarCelula(string Texto, int AlinhamentoHorizontal = Element.ALIGN_LEFT, float TamanhoFonte = 12,
            int Estilo = Font.NORMAL, BaseColor CorTexto = null)
        {
            Font FonteCelula = new Font(Configuracao.FonteBase, TamanhoFonte, Estilo, CorTexto);
            PdfPCell Celula = new PdfPCell(new Phrase(Texto, FonteCelula));
            Celula.HorizontalAlignment = AlinhamentoHorizontal;
            Celula.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
            Celula.Border = 0;
            Celula.BorderWidthBottom = 1;
            Celula.FixedHeight = 25;
            Tabela.AddCell(Celula);
        }

        public void FecharArquivo(bool Abrir = false)
        {
            Pdf.Add(Tabela);
            Pdf.Close();
            Arquivo.Close();
            if (Abrir)
            {
                AbriRelatorio();
            }
        }

        private void AbriRelatorio()
        {
            if (File.Exists(Configuracao.CaminhoArquivo))
            {
                Process.Start(new ProcessStartInfo()
                {
                    Arguments = $"/c start {Configuracao.CaminhoArquivo}",
                    FileName = "cmd.exe",
                    CreateNoWindow = true,
                });
            }
        }


    }
}
