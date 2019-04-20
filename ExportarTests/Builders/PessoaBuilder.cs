using Bogus;
using ExportarTests.Entities;

namespace ExportarTests.Builder
{
    public class PessoaBuilder
    {       
        private Faker _faker = new Faker();
        private int Id;
        private string Nome;
        private int Idade;
        private string Sexo;

        public PessoaBuilder()
        {
            Id = _faker.Random.Int(0, 1000000);
            Nome = _faker.Person.FirstName;
            Idade = _faker.Random.Int(0,90);
            Sexo = _faker.Person.Gender.ToString();
        }

        public static PessoaBuilder Novo()
        {
            return new PessoaBuilder();
        }

        public Pessoa Build()
        {
            Pessoa pessoa = new Pessoa();
            pessoa.Id = Id;
            pessoa.Nome = Nome;
            pessoa.Idade = Idade;
            pessoa.Sexo = Sexo;
            return pessoa;
        }
             
        public PessoaBuilder ComNome(string nome)
        {
            Nome = nome;
            return this;
        }

        public PessoaBuilder ComId(int id)
        {
            Id = id;
            return this;
        }       

    }
}
