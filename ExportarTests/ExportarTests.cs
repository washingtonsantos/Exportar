using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExportarTests.Builder;

namespace ExportarTests
{
    [TestClass]
    public class ExportarTests
    {
        [TestMethod]
        public void ExportarPessoa()
        {                       
            var listaPessoas = Factory.Pessoas();

            Exportar.XLSX exp = new Exportar.XLSX();
            exp.GerarArquivo(listaPessoas, @"C:\washington_teste.xlsx");
        }
    }
}
