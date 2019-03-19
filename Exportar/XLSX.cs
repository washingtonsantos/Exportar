using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Exportar
{
    public class XLSX
    {               
        /// <summary>
        /// Gerar Arquivo Excel a partir de uma lista Generica
        /// Parâmetro "salvarPlanilha" é usado para adicionar mais de uma aba na planilha
        /// Ex: Exportar.Exportar.XLSX(Objeto1, @"C:\Temp\relatorio.xlsx",false);
        ///     Exportar.Exportar.XLSX(Objeto2, @"C:\Temp\relatorio.xlsx",true);
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="listaGenerica"></param>
        /// <param name="caminhoParaSalvarArquivo"></param>
        /// <param name="filtroAtivadoNoCabecalho"></param>
        /// <param name="backgroundCabecalho"></param>
        /// <param name="foregroundCabecalho"></param>
        /// <param name="nomeDaAbaDaPlanilha"></param>
        /// <param name="formatarColunadoubleParaMonetario"></param>
        /// <param name="formatarColunaDateTimeRemovendoTime"></param>
        public byte[] GerarArquivo<T>(List<T> listaGenerica, string caminhoParaSalvarArquivo = @"C:\")
            where T : class
        {
            byte[] array = null;
            var fi = new FileInfo(caminhoParaSalvarArquivo);

            if (listaGenerica != null && listaGenerica.Count > 0)
            {
                using (var pck = new ExcelPackage(fi))
                {
                    //Reflection da Lista de <T>
                    var mi = typeof(T).GetProperties().Where(pi => pi.Name != "Col7")
                     .Select(pi => (MemberInfo)pi).ToArray();

                    var aba = DateTime.Now.ToString().Trim();
                    var existeAba = pck.Workbook.Worksheets.Where(w => w.Name == aba).FirstOrDefault()?.ToString() ?? "";
                    var novaAba = existeAba.Length > 0 ? existeAba + "1" : aba;
                    var worksheet = pck.Workbook.Worksheets.Add(novaAba);

                    //Processamento da Planilha
                    worksheet.Cells.LoadFromCollection(listaGenerica, true, TableStyles.None, BindingFlags.Public | BindingFlags.Instance, mi);
                    var ultimaLinha = worksheet.Dimension.End.Row;
                    var ultimaColuna = worksheet.Dimension.End.Column;

                    //Metodos de Formatação da Planilha
                    FormataCabecalho(worksheet, ultimaLinha, ultimaColuna);
                    FormataBordas(worksheet, ultimaLinha, ultimaColuna);

                    worksheet.Cells.AutoFitColumns();

                    if (caminhoParaSalvarArquivo.Length > 5)
                    {
                        pck.Save();
                    }

                    array = pck.GetAsByteArray();

                    //return array;
                }
            }

            return array;
        }

        public byte[] GerarArquivo<T>(List<T> listaGenerica, string caminhoParaSalvarArquivo = @"C:\", string nomeDaAbaDaPlanilha = "Nome do Objeto", bool filtroAtivadoNoCabecalho = true, string backgroundCabecalho = "Red", string foregroundCabecalho = "White", bool formatarColunadoubleParaMonetario = true, bool formatarColunaDateTimeRemovendoTime = true)
          where T : class
        {
            byte[] array = null;
            var fi = new FileInfo(caminhoParaSalvarArquivo);

            if (listaGenerica != null && listaGenerica.Count > 0)
            {
                using (var pck = new ExcelPackage(fi))
                {
                    //Reflection da Lista de <T>
                    var mi = typeof(T).GetProperties().Where(pi => pi.Name != "Col7")
                     .Select(pi => (MemberInfo)pi).ToArray();

                    var aba = nomeDaAbaDaPlanilha == "Nome do Objeto" ? mi[0].ReflectedType.Name : nomeDaAbaDaPlanilha;
                    var existeAba = pck.Workbook.Worksheets.Where(w => w.Name == aba).FirstOrDefault()?.ToString() ?? "";
                    var novaAba = existeAba.Length > 0 ? existeAba + "1" : aba;
                    var worksheet = pck.Workbook.Worksheets.Add(novaAba);

                    //Processamento da Planilha
                    worksheet.Cells.LoadFromCollection(listaGenerica, true, TableStyles.None, BindingFlags.Public | BindingFlags.Instance, mi);
                    var ultimaLinha = worksheet.Dimension.End.Row;
                    var ultimaColuna = worksheet.Dimension.End.Column;

                    //Metodos de Formatação da Planilha
                    FormataCabecalho(worksheet, backgroundCabecalho, foregroundCabecalho, filtroAtivadoNoCabecalho, ultimaLinha, ultimaColuna);
                    FormataBordas(worksheet, ultimaLinha, ultimaColuna);

                    //Formatar Celulas
                    FormataCelulas(listaGenerica, worksheet, formatarColunadoubleParaMonetario, formatarColunaDateTimeRemovendoTime, ultimaLinha, ultimaColuna);
                    worksheet.Cells.AutoFitColumns();

                    if (caminhoParaSalvarArquivo.Length > 5)
                    {
                        pck.Save();
                    }

                    array = pck.GetAsByteArray();

                    //return array;
                }
            }

            return array;
        }
        private void FormataCelulas<T>(List<T> listaGenericaT, ExcelWorksheet worksheet, bool formatarColunadoubleParaMonetario, bool formatarColunaDateTimeRemovendoTime, int ultimaLinha, int ultimaColuna)
            where T : class
        {
            var propriedades = typeof(T);
            int col = 1;
            int colunaAtual = 0;

            foreach (var item in propriedades.GetProperties())
            {
                var propriedade = item.PropertyType.Name;

                if (propriedade == "Double" && formatarColunadoubleParaMonetario == true)
                {
                    worksheet.Cells[2, col, ultimaLinha, col].Style.Numberformat.Format = "_-R$* #,##0.00_-;-R$* #,##0.00_-;_-R$* \"-\"??_-;_-@_-";
                }
                else if (propriedade == "DateTime" && formatarColunaDateTimeRemovendoTime == true)
                {
                    worksheet.Cells[2, colunaAtual, ultimaLinha, col].Style.Numberformat.Format = "dd/mm/yyyy";
                }

                colunaAtual += 1;
                col += 1;
            }
        }
       
        private void FormataCabecalho(ExcelWorksheet worksheet, string backgroundCabecalho, string foregroundCabecalho, bool filtroAtivadoNoCabecalho, int ultimaLinha, int ultimaColuna)
        {
            //Estilho Cabeçalho
            worksheet.Cells[1, 1, 1, ultimaColuna].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[1, 1, 1, ultimaColuna].Style.Fill.BackgroundColor.SetColor(color: System.Drawing.Color.FromName(backgroundCabecalho));
            worksheet.Cells[1, 1, 1, ultimaColuna].Style.Font.Color.SetColor(color: System.Drawing.Color.FromName(foregroundCabecalho));
            worksheet.Cells[1, 1, 1, ultimaColuna].Style.Font.Bold = true;
            worksheet.Cells[1, 1, 1, ultimaColuna].AutoFilter = filtroAtivadoNoCabecalho;
            worksheet.View.FreezePanes(2, 1);
        }
        private void FormataCabecalho(ExcelWorksheet worksheet, int ultimaLinha, int ultimaColuna)
        {
            //Estilho Cabeçalho
            worksheet.Cells[1, 1, 1, ultimaColuna].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[1, 1, 1, ultimaColuna].Style.Fill.BackgroundColor.SetColor(color: System.Drawing.Color.FromName("Red"));
            worksheet.Cells[1, 1, 1, ultimaColuna].Style.Font.Color.SetColor(color: System.Drawing.Color.FromName("White"));
            worksheet.Cells[1, 1, 1, ultimaColuna].Style.Font.Bold = true;
            worksheet.Cells[1, 1, 1, ultimaColuna].AutoFilter = true;
            worksheet.View.FreezePanes(2, 1);
        }

        private void FormataBordas(ExcelWorksheet worksheet, int ultimaLinha, int ultimaColuna)
        {
            //Estilho de Bordas do Documento.
            worksheet.Cells[1, 1, ultimaLinha, ultimaColuna].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[1, 1, ultimaLinha, ultimaColuna].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[1, 1, ultimaLinha, ultimaColuna].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[1, 1, ultimaLinha, ultimaColuna].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[1, 1, ultimaLinha, ultimaColuna].Style.Border.Top.Color.SetColor(color: System.Drawing.Color.Black);
            worksheet.Cells[1, 1, ultimaLinha, ultimaColuna].Style.Border.Bottom.Color.SetColor(color: System.Drawing.Color.Black);
            worksheet.Cells[1, 1, ultimaLinha, ultimaColuna].Style.Border.Left.Color.SetColor(color: System.Drawing.Color.Black);
            worksheet.Cells[1, 1, ultimaLinha, ultimaColuna].Style.Border.Right.Color.SetColor(color: System.Drawing.Color.Black);
        }


    }
}
