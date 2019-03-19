using Bogus;
using ExportarTests.Entitie;

namespace ExportarTests.Builder
{
    public class PessoaBuilder
    {       
        private Faker _faker = new Faker();
        private int Id { get; set; }
        private string Nome { get; set; }

        public PessoaBuilder()
        {
            Id = _faker.Random.Int(0, 1000000);
            Nome = _faker.Person.FirstName;
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
