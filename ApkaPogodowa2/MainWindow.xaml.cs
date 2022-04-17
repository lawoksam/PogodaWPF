using ApkaPogodowa2.WeatherApi;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ApkaPogodowa2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
            isFileCreated();
            var file = new FileInfo(@"c:\c#\pogoda.xlsx");
            using (var package = new ExcelPackage(file))
            {
                var sheet = package.Workbook.Worksheets.Add("Pogoda");
                sheet.Cells["A1"].Value = "City";
                sheet.Cells["B1"].Value = "Temperature";
                sheet.Cells["C1"].Value = "Cloudiness";
                sheet.Cells["D1"].Value = "Humidity";
                sheet.Cells["E1"].Value = "Pressure";
                sheet.Cells["F1"].Value = "Wind Speed";
                sheet.Cells["G1"].Value = "Date";
                package.Save();
            }
        }
        /// <summary>
        /// checks existance of file
        /// </summary>
        private static void isFileCreated()
        {
            if (File.Exists(@"c:\c#\pogoda.xlsx"))
            {
                File.Delete(@"c:\c#\pogoda.xlsx");
            }
        }
        #region GetCityWeatherInfoFromApi
        /// <summary>
        /// Get city weather data from API depending on input city name
        /// </summary>
        /// <param name="cityName">City name</param>
        static string GetCityWeatherInfoFromAPI(string cityName)
        {
            string responseFromServer;
            WebRequest request = WebRequest.Create(
                $"https://api.openweathermap.org/data/2.5/weather?q=" + cityName + "&appid=da64753049352fb4609a3ac38800a9e6");
            try
            {
                using (WebResponse response = request.GetResponse())
                {
                    using (Stream dataStream = response.GetResponseStream())
                    {
                        StreamReader reader = new StreamReader(dataStream);
                        responseFromServer = reader.ReadToEnd();

                    }
                }
                return responseFromServer;
            }
            catch (WebException)
            {
                return null;
            }
        }
        #endregion
        #region CreateCityWeatherObject
        static CityWeather CreateCityWeatherObject(string weatherInfo)
        {

            CityWeatherFromAPI cityWeatherData = JsonConvert.DeserializeObject<CityWeatherFromAPI>(weatherInfo);
            double tempC = Math.Round(cityWeatherData.main.Temp - 273.15, 2);
            double tempF = Math.Round((tempC * 1.8) + 32);
            CityWeather city = new CityWeather(cityWeatherData.Name, cityWeatherData.weather[0].Main, cityWeatherData.weather[0].Description, tempC, tempF, cityWeatherData.main.Pressure, cityWeatherData.main.Humidity, cityWeatherData.wind.Speed, cityWeatherData.clouds.all);

            return city;
        }
        #endregion
        #region GetCityWeatherInfo
        static CityWeather GetCityWeatherInfo(string cityName)
        {
            string weatherInfo = GetCityWeatherInfoFromAPI(cityName);
            if (weatherInfo == null)
            {
                return null;
            }
            else
            {
                return CreateCityWeatherObject(weatherInfo);

            }

        }
        #endregion
        private void miasto_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void przycisk_Click(object sender, RoutedEventArgs e)
        {
            Thread t = new Thread(() =>delayLoop());
            t.Start();

        }
        /// <summary>
        /// Checks weather every declared time
        /// </summary>
        private void delayLoop()
        {
            int i = 2;
            //Lista.Items.Clear();
            for (int z = 1; z <= 3; z++)
            {
                Dispatcher.Invoke(() => miasto.Text = $"Number of updates: {z}");
                Dispatcher.Invoke(() => timeCounter.Text = $"Last update: {DateTime.Now.ToString("HH:mm")}");
                List<CityWeather> cityWeathersUpToDate = citiesListWeatherUpdate();
                i = CitiesToExcelLoop(i, cityWeathersUpToDate);
                System.Threading.Thread.Sleep(6000);

            }
        }

        private void timerCounter(int i)
        {
            miasto.Text = $"Pozostało {240 - i} minut";
        }

        private static List<CityWeather> citiesListWeatherUpdate()
        {
            List<CityWeather> cityWeathers = new List<CityWeather>(); // Lista miast do pobierania pogody
            cityWeathers.Add(GetCityWeatherInfo("Gdansk"));
            cityWeathers.Add(GetCityWeatherInfo("Koscierzyna"));
            cityWeathers.Add(GetCityWeatherInfo("Nowa Karczma"));
            cityWeathers.Add(GetCityWeatherInfo("Stara kiszewa"));
            cityWeathers.Add(GetCityWeatherInfo("Warszawa"));
            cityWeathers.Add(GetCityWeatherInfo("Halle"));
            return cityWeathers;
        }
        #region Cities To Excel
        private int CitiesToExcelLoop(int i, List<CityWeather> cityWeathers)
        {

            for (int j = 0; j < cityWeathers.Count; j++) // Pętla po miastach
            {
                if (cityWeathers[j] != null)
                {
                    writeCityWeatherToFile(i, cityWeathers[j]);
                    i++;
                }
                else
                {
                    miasto.Text = "Niepoprawna nazwa miasta";
                }
            }

            return i;
        }

        private void writeCityWeatherToFile(int i, CityWeather city)
        {
            //miasto.Text = city.Name;
            //Lista.Items.Add(("Temperatura w C: " + city.TempC));
            //Lista.Items.Add(("Temperatura w F: " + city.TempF));
            //Lista.Items.Add(("Zachmurzenie: " + city.CloudIness + "%"));
            //Lista.Items.Add(("Wilgotność: " + city.Humidity + "%"));
            //Lista.Items.Add(("Ciśnienie: " + city.Pressure + "kpa"));
            //Lista.Items.Add(("Prędkość wiatru: " + city.WindSpeed + "km/h"));

            var file = new FileInfo(@"c:\c#\pogoda.xlsx");
            using (var package = new ExcelPackage(file))
            {
                var sheet = package.Workbook.Worksheets[0];
                sheet.Cells["A" + i].Value = city.Name;
                sheet.Cells["B" + i].Value = city.TempC;
                sheet.Cells["C" + i].Value = city.CloudIness;
                sheet.Cells["D" + i].Value = city.Humidity;
                sheet.Cells["E" + i].Value = city.Pressure;
                sheet.Cells["F" + i].Value = city.WindSpeed;
                sheet.Cells["G" + i].Value = DateTime.Now.ToString();

                // Save to file
                package.Save();
            }
        }
        #endregion

        private void ListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }


        private void TextBox_TextChanged2(object sender, TextChangedEventArgs e)
        {

        }
    }
}
