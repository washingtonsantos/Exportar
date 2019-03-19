using ExportarTests.Entitie;
using System.Collections.Generic;

namespace ExportarTests.Builder
{
    public class Factory
    {
         public static List<Pessoa> Pessoas()
        {
            List<Pessoa> pessoas = new List<Pessoa>();
            pessoas.Add(PessoaBuilder.Novo().Build());

            return pessoas;
        }
    }
}
