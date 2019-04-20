using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExportarTests.Builders;

namespace ExportarTests
{
    [TestClass]
    public class ExportarTests
    {
        [TestMethod]
        public void DeveExportarUmaListaDePessoas()
        {                       
            var pessoas = Factory.Pessoas();

            var exp = new Exportar.XLSX();
            exp.GerarArquivo(pessoas, @"C:\washington_teste.xlsx");
        }

        //exportando diversas 'planilhas'
        [TestMethod]
        public void DeveExportarTresPlanilhasNoMesmoArquivo()
        {
            for (int novaPlanilha = 0; novaPlanilha < 3; novaPlanilha++)
            {
                var pessoas = Factory.Pessoas();

                var exp = new Exportar.XLSX();
                exp.GerarArquivo(pessoas, @"C:\washington_teste.xlsx");
            }           
        }

        //exportando diversas planilhas de objetos diferentes
        [TestMethod]
        public void DeveExportarListasDeObjetosDiferentes()
        {          
                var pessoas = Factory.Pessoas();
                var carros = Factory.Carros();

                var exp = new Exportar.XLSX();
                exp.GerarArquivo(pessoas, @"C:\washington_teste.xlsx");
                exp.GerarArquivo(carros, @"C:\washington_teste.xlsx");
        }

        [TestMethod]
        public void DeveRetornarUmArrayDeBytesDeUmaPlanilhaPronta()
        {
            var pessoas = Factory.Pessoas();

            var exp = new Exportar.XLSX();
            exp.GerarArquivo(pessoas);
        }

        [TestMethod]
        public void DeveRenomearONomeDaAbaDaPlanilha()
        {
            var pessoas = Factory.Pessoas();

            var exp = new Exportar.XLSX();
            exp.GerarArquivo(pessoas,@"C:\washington_teste.xlsx", "Nome Personalizado");
        }

        [TestMethod]
        public void DeveAlterarACorDoBackGroundDoCabecalhoDaPlanilha()
        {
            var pessoas = Factory.Pessoas();

            var exp = new Exportar.XLSX();
            exp.GerarArquivo(pessoas, @"C:\washington_teste.xlsx", "Nome Personalizado","Blue","Black");
        }

        [TestMethod]
        public void DeveSerInvalidaAExportacaoCasoOCaminhoInformadoSejaInvalido()
        {
            var pessoas = Factory.Pessoas();

            var exp = new Exportar.XLSX();
            exp.GerarArquivo(pessoas, @"C:\CaminhoNaoExiste\washington_teste.xlsx");
        }
    }
}
