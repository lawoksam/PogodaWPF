using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ApkaPogodowa2
{
     class CityWeather
    {
        #region Properties

        public string Name { get; set; }
        public string WeatherType { get; set; }
        public string WeatherDescription { get; set; }
        public double TempC { get; set; }
        public double TempF { get; set; }
        public double Pressure { get; set; }
        public int Humidity { get; set; }
        public double WindSpeed { get; set; }
        public int CloudIness { get; set; }

        #endregion
        #region Constructor
        public CityWeather(string name, string weatherType, string weatherDescription, double tempC, double tempF, double pressure, int humidity, double windSpeed, int cloudIness)
        {
            Name = name;
            WeatherType = weatherType;
            WeatherDescription = weatherDescription;
            TempC = tempC;
            TempF = tempF;
            Pressure = pressure;
            Humidity = humidity;
            WindSpeed = windSpeed;
            CloudIness = cloudIness;
        }
        #endregion Constructor

        public void WypiszInfo()
        {
            Console.WriteLine($"Miasto: {Name}");
            Console.WriteLine($"Pogoda: {WeatherType}");
            Console.WriteLine($"Opis: {WeatherDescription}");
            Console.WriteLine($"Temperatura w C: {TempC}");
            Console.WriteLine($"Temperatura w F: {TempF}");
            Console.WriteLine($"Ciśnienie: {Pressure}");
            Console.WriteLine($"Wilgotność: {Humidity}");
            Console.WriteLine($"Prędkość wiatru: {WindSpeed}");
            Console.WriteLine($"Zachmurzenie: {CloudIness}");


        }

    }
}
