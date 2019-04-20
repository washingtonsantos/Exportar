using ExportarTests.Builder;
using ExportarTests.Entities;
using System.Collections.Generic;

namespace ExportarTests.Builders
{
    public class Factory
    {
         public static List<Pessoa> Pessoas()
        {
            List<Pessoa> pessoas = new List<Pessoa>();
            pessoas.Add(PessoaBuilder.Novo().Build());

            return pessoas;
        }

        public static List<Carro> Carros()
        {
            List<Carro> carros = new List<Carro>();
            carros.Add(CarroBuilder.Novo().Build());

            return carros;
        }
    }
}

