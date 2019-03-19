using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExportarTests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var pessoa = Builder.Pessoa.PessoaBuilder();
           
            var listaPessoas = Builder.Pessoa.Pessoas(pessoa);

            Exportar.XLSX exp = new Exportar.XLSX();
            exp.GerarArquivo(listaPessoas, @"C:\washington_teste.xlsx");
        }
    }
}
