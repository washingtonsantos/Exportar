using Bogus;
using System.Collections.Generic;

namespace ExportarTests.Builder
{
    public class PessoaBuilder
    {
        private int Id { get;set; }
        private string Nome { get; set; }

        public static PessoaBuilder Novo()
        {            
            return new PessoaBuilder();
        }

        public PessoaBuilder Build()
        {
            PessoaBuilder pessoa = new PessoaBuilder();
            var _faker = new Faker();
            pessoa.Id = _faker.Random.Int();
            pessoa.Nome = _faker.Person.FirstName;
            return pessoa;
        }
             
        public PessoaBuilder ComNome(string nome)
        {
            PessoaBuilder pessoa = new PessoaBuilder();
            pessoa.Nome = nome;
            return pessoa;
        }

        public int ComId(int id)
        {
            PessoaBuilder pessoa = new PessoaBuilder();
            pessoa.Id = id;
            return pessoa.Id;
        }

        public static List<PessoaBuilder> Pessoas(PessoaBuilder pessoa)
        {
            List<PessoaBuilder> pessoas = new List<PessoaBuilder>();
            pessoas.Add(pessoa);
             
            return pessoas;
        }

    }
}
