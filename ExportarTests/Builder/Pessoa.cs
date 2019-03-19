using Bogus;
using System.Collections.Generic;

namespace ExportarTests.Builder
{
    public class Pessoa
    {
        public int Id { get;set; }
        public string Nome { get; set; }

        public static Pessoa Novo()
        {            
            return new Pessoa();
        }

        public static Pessoa PessoaBuilder()
        {
            Pessoa pessoa = new Pessoa();
            var _faker = new Faker();
            pessoa.Id = _faker.Random.Int();
            pessoa.Nome = _faker.Person.FirstName;
            return pessoa;
        }

        public static Pessoa NomePessoa(string nome)
        {
            Pessoa pessoa = new Pessoa();
            pessoa.Nome = nome;
            return pessoa;
        }

        public static int IdPessoa(int id)
        {
            Pessoa pessoa = new Pessoa();
            pessoa.Id = id;
            return pessoa.Id;
        }

        public static List<Pessoa> Pessoas(Pessoa pessoa)
        {
            List<Pessoa> pessoas = new List<Pessoa>();
            pessoas.Add(pessoa);
             
            return pessoas;
        }

    }
}
