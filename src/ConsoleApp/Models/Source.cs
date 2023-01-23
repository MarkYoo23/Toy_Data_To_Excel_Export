namespace ConsoleApp.Models
{
    public class Source
    {
        public int Id { get; set; }
        public DateTime Created { get; set; }

        public decimal SampleTemperature { get; set; }
        public decimal ReferenceTemperature { get; set; }
        public decimal GasTemperature { get; set; }
        public decimal TargetTemperature { get; set; }
        public decimal ContainerPlatePosition { get; set; }
        public decimal ProcessingMinuites { get; set; }

        public decimal GasTemperatureDifferential { get; set; }
        public decimal SampleTemperatureDifferential { get; set; }
        public decimal ReferenceTemperatureDifferential { get; set; }
        public decimal TargetTemperatureDifferential { get; set; }
    }
}
