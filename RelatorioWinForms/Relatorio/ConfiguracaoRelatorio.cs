using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Relatorio
{
    public class ConfiguracaoRelatorio
    {
        public int NumeroColunas { get; set; }
        public string[] TitulosColuna { get; set; }

        #region Margem
        public float MargemSuperior { private get; set; } = 15;
        public float MargemInferior { private get; set; } = 15;
        public float MargemEsquerda { private get; set; } = 15;
        public float MargemDireita { private get; set; } = 15;
        #endregion

        #region Documento
        public BaseFont FonteBase { get; set; }
        public string CaminhoArquivo { get; set; }
        public bool EstiloZebrado { get; set; } = false;

        #endregion

        private const float PxPorMm = 2.857142857142857f;

        public ConfiguracaoRelatorio() 
        {
            FonteBase = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
            CaminhoArquivo = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,$"PDF\\{DateTime.Now.ToString("ddMMyyHHmmss")}.pdf");
            if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + "PDF"))
                Directory.CreateDirectory(CaminhoArquivo);
            NumeroColunas = 2;
            TitulosColuna = new string[] { "Titulo1", "Titulo2" };
        }
        public ConfiguracaoRelatorio(int NumeroColunas, string[] TitulosColuna, 
            BaseFont FonteBase, string CaminhoArquivo, bool EstiloZebrado)
        {
            this.TitulosColuna = TitulosColuna;
            this.NumeroColunas = NumeroColunas;
            this.FonteBase = FonteBase;
            this.CaminhoArquivo = CaminhoArquivo;
            this.EstiloZebrado = EstiloZebrado;
        }

        public float getSuperior()
        {
            return MargemSuperior * PxPorMm;
        }

        public float getInferior()
        {
            return MargemInferior * PxPorMm;
        }

        public float getEsquerda()
        {
            return MargemEsquerda * PxPorMm;
        }

        public float getDireita()
        {
            return MargemDireita * PxPorMm;
        }

    }
}
