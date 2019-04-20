using Bogus;
using ExportarTests.Entities;

namespace ExportarTests.Builders
{
    public class CarroBuilder
    {
        private Faker _faker = new Faker();
        private int Portas;
        private string Modelo;
        private string Fabricante;
        private string Combustivel;

        public CarroBuilder()
        {
            Portas = _faker.Random.Int(2,5);
            Modelo = _faker.Vehicle.Model();
            Fabricante = _faker.Vehicle.Manufacturer();
            Combustivel = _faker.Vehicle.Fuel();
        }

        public static CarroBuilder Novo()
        {
            return new CarroBuilder();
        }

        public Carro Build()
        {
            var carro = new Carro();
            carro.Portas = Portas;
            carro.Modelo = Modelo;
            carro.Fabricante = Fabricante;
            carro.Combustivel = Combustivel;

            return carro;
        }
    }
}
