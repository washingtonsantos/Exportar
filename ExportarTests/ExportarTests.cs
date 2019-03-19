using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExportarTests
{
    [TestClass]
    public class ExportarTests
    {
        [TestMethod]
        public void ExportarPessoa()
        {
            var pessoa = Builder.PessoaBuilder.Novo().Build();
           
            var listaPessoas = Builder.PessoaBuilder.Pessoas(pessoa);

            Exportar.XLSX exp = new Exportar.XLSX();
            exp.GerarArquivo(listaPessoas, @"C:\washington_teste.xlsx");
        }
    }
}
