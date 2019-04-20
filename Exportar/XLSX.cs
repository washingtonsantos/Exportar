using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Exportar
{
    public class XLSX
    {               

        /// <summary>
        /// exportar excel a partir de uma lista genérica.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="listaGenerica"></param>
        public byte[] GerarArquivo<T>(List<T> listaGenerica)
            where T : class
        {
            byte[] array = null;

            FileInfo fi = new FileInfo("temp");

            if (listaGenerica != null && listaGenerica.Count > 0)
            {
                using (var pck = new ExcelPackage(fi))
                {
                    //Reflection da Lista de <T>
                    var mi = typeof(T).GetProperties()
                     .Select(pi => (MemberInfo)pi).ToArray();

                    var nomePlanilha = GerarNomeDaPlanilha<T>(pck);

                    var worksheet = pck.Workbook.Worksheets.Add(nomePlanilha);

                    //Processamento da Planilha
                    worksheet.Cells.LoadFromCollection(listaGenerica, true, TableStyles.None, BindingFlags.Public | BindingFlags.Instance, mi);
                 
                    worksheet.Cells.AutoFitColumns();

                    array = pck.GetAsByteArray();
                }
            }

            return array;
        }

        /// <summary>
        /// exportar excel a partir de uma lista genérica, necessita caminho válido com extensão '.xlsx'.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="listaGenerica"></param>
        /// <param name="caminhoParaSalvarArquivo"></param>
        public byte[] GerarArquivo<T>(List<T> listaGenerica, string caminhoParaSalvarArquivo)
            where T : class
        {
            byte[] array = null;

            DirectoryInfo directoryInfo = new DirectoryInfo(caminhoParaSalvarArquivo);

            if (!directoryInfo.Root.Exists)
                return null;

            if (!directoryInfo.Parent.Exists)
                return null;

            var fi = new FileInfo(caminhoParaSalvarArquivo);

            if (listaGenerica != null && listaGenerica.Count > 0)
            {
                using (var pck = new ExcelPackage(fi))
                {
                    //Reflection da Lista de <T>
                    var mi = typeof(T).GetProperties()
                     .Select(pi => (MemberInfo)pi).ToArray();

                    var nomePlanilha = GerarNomeDaPlanilha<T>(pck);

                    var worksheet = pck.Workbook.Worksheets.Add(nomePlanilha);

                    //Processamento da Planilha
                    worksheet.Cells.LoadFromCollection(listaGenerica, true, TableStyles.None, BindingFlags.Public | BindingFlags.Instance, mi);
                                     
                    worksheet.Cells.AutoFitColumns();

                    if (caminhoParaSalvarArquivo.Length > 2)
                    {
                        pck.Save();
                    }

                    array = pck.GetAsByteArray();
                }
            }

            return array;
        }

        /// <summary>
        /// exportar excel a partir de uma lista genérica, necessita caminho válido com extensão '.xlsx',informe o nome da planilha.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="listaGenerica"></param>
        /// <param name="caminhoParaSalvarArquivo"></param>
        /// <param name="nomeDaPlanilha"></param>
        /// <param name="backgroundCabecalho"></param>
        /// <param name="foregroundCabecalho"></param>
        /// <param name="formatarColunadoubleParaMonetario"></param>
        /// <param name="formatarColunaDateTimeRemovendoTime"></param>
        public byte[] GerarArquivo<T>(List<T> listaGenerica,  string caminhoParaSalvarArquivo, string nomeDaPlanilha)
          where T : class
        {
            byte[] array = null;

            DirectoryInfo directoryInfo = new DirectoryInfo(caminhoParaSalvarArquivo);


            if (!directoryInfo.Root.Exists)
                return null;

            if (!directoryInfo.Parent.Exists)
                return null;

            var fi = new FileInfo(caminhoParaSalvarArquivo);

            if (listaGenerica != null && listaGenerica.Count > 0)
            {
                using (var pck = new ExcelPackage(fi))
                {
                    //Reflection da Lista de <T>
                    var mi = typeof(T).GetProperties().Where(pi => pi.Name != "Col7")
                     .Select(pi => (MemberInfo)pi).ToArray();

                    var nomePlanilha = GerarNomeDaPlanilha<T>(pck);

                    var worksheet = pck.Workbook.Worksheets.Add(nomePlanilha);

                    //Processamento da Planilha
                    worksheet.Cells.LoadFromCollection(listaGenerica, true, TableStyles.None, BindingFlags.Public | BindingFlags.Instance, mi);

                    worksheet.Cells.AutoFitColumns();

                    if (caminhoParaSalvarArquivo.Length > 2 && fi.Extension.ToLower() == ".xlsx")
                            pck.Save();

                    array = pck.GetAsByteArray();

                }
            }

            return array;
        }

        /// <summary>
        /// exportar excel a partir de uma lista genérica, informe o nome da planilha, necessita caminho válido com extensão '.xlsx', informe a cor do background do cabeçalho (cor deve ser informada no idioma EN ex.: 'RED'),informe o foreground (cor deve ser informada no idioma EN ex.: 'WHITE').
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="listaGenerica"></param>
        /// <param name="caminhoParaSalvarArquivo"></param>
        /// <param name="nomeDaPlanilha"></param>
        /// <param name="backgroundCabecalho"></param>
        /// <param name="foregroundCabecalho"></param>
        /// <param name="formatarColunadoubleParaMonetario"></param>
        /// <param name="formatarColunaDateTimeRemovendoTime"></param>
        public byte[] GerarArquivo<T>(List<T> listaGenerica, string caminhoParaSalvarArquivo, string nomeDaPlanilha, string backgroundCabecalho = "Red", string foregroundCabecalho = "White")
          where T : class
        {
            byte[] array = null;

            DirectoryInfo directoryInfo = new DirectoryInfo(caminhoParaSalvarArquivo);

            if (!directoryInfo.Root.Exists)
                return null;

            if (!directoryInfo.Parent.Exists)
                return null;

            var fi = new FileInfo(caminhoParaSalvarArquivo);

            if (listaGenerica != null && listaGenerica.Count > 0)
            {
                using (var pck = new ExcelPackage(fi))
                {
                    //Reflection da Lista de <T>
                    var mi = typeof(T).GetProperties().Where(pi => pi.Name != "Col7")
                     .Select(pi => (MemberInfo)pi).ToArray();

                    var nomePlanilha = GerarNomeDaPlanilha<T>(pck);

                    var worksheet = pck.Workbook.Worksheets.Add(nomePlanilha);

                    //Processamento da Planilha
                    worksheet.Cells.LoadFromCollection(listaGenerica, true, TableStyles.None, BindingFlags.Public | BindingFlags.Instance, mi);
                  
                    //Metodos de Formatação da Planilha
                    FormataCabecalho(worksheet, backgroundCabecalho, foregroundCabecalho);
                    FormataBordas(worksheet);

                    worksheet.Cells.AutoFitColumns();

                    if (caminhoParaSalvarArquivo.Length > 5)
                    {
                        pck.Save();
                    }

                    array = pck.GetAsByteArray();
                }
            }

            return array;
        }

        private string GerarNomeDaPlanilha<T>(ExcelPackage pck)
        {
            var nomeDoObjeto = typeof(T).Name;

            var ultimaPlanilha = pck.Workbook.Worksheets.Select(x => x.Name).LastOrDefault();

            return ultimaPlanilha == null ? nomeDoObjeto : nomeDoObjeto + pck.Workbook.Worksheets.Count;
        }
        private void FormataCabecalho(ExcelWorksheet worksheet, string backgroundCabecalho, string foregroundCabecalho)
        {
            var ultimaLinha = worksheet.Dimension.End.Row;
            var ultimaColuna = worksheet.Dimension.End.Column;

            //Estilho Cabeçalho
            worksheet.Cells[1, 1, 1, ultimaColuna].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[1, 1, 1, ultimaColuna].Style.Fill.BackgroundColor.SetColor(color: System.Drawing.Color.FromName(backgroundCabecalho));
            worksheet.Cells[1, 1, 1, ultimaColuna].Style.Font.Color.SetColor(color: System.Drawing.Color.FromName(foregroundCabecalho));
            worksheet.Cells[1, 1, 1, ultimaColuna].Style.Font.Bold = true;
            worksheet.Cells[1, 1, 1, ultimaColuna].AutoFilter = true;
            worksheet.View.FreezePanes(2, 1);
        }
        [Obsolete]
        private void FormataCabecalho(ExcelWorksheet worksheet)
        {
            var ultimaLinha = worksheet.Dimension.End.Row;
            var ultimaColuna = worksheet.Dimension.End.Column;
            //Estilho Cabeçalho
            worksheet.Cells[1, 1, 1, ultimaColuna].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[1, 1, 1, ultimaColuna].Style.Fill.BackgroundColor.SetColor(color: System.Drawing.Color.FromName("Red"));
            worksheet.Cells[1, 1, 1, ultimaColuna].Style.Font.Color.SetColor(color: System.Drawing.Color.FromName("White"));
            worksheet.Cells[1, 1, 1, ultimaColuna].Style.Font.Bold = true;
            worksheet.Cells[1, 1, 1, ultimaColuna].AutoFilter = true;
            worksheet.View.FreezePanes(2, 1);
        }
        private void FormataBordas(ExcelWorksheet worksheet)
        {
            var ultimaLinha = worksheet.Dimension.End.Row;
            var ultimaColuna = worksheet.Dimension.End.Column;

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
