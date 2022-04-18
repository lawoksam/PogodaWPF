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
        int defTime = 2;
        private void przycisk_Click(object sender, RoutedEventArgs e)
        {
            
            if(timeCounter.Text != "Set the measurement time (minutes)")
            {
                try
                {
                    if (int.Parse(timeCounter.Text) < 10)
                    { defTime = 1; }
                    else
                    { defTime = int.Parse(timeCounter.Text) / 10; }
                    
                }
                catch (FormatException)
                {

                }
            }
            Thread t = new Thread(() =>delayLoop());
            t.Start();

        }


        List<CityWeather> cityWeathersUpToDate = citiesListWeatherUpdate(); // public list of cities
        bool isAddClicked = false;

        private void przycisk_Click_Add(object sender, RoutedEventArgs e)
        {
            if (isAddClicked = false)
            {
                Lista.Items.Clear();
                isAddClicked = true;
            }
            if(GetCityWeatherInfoFromAPI(numberOfUpdates.Text) != null)
            {
                
                cityWeathersUpToDate.Add(GetCityWeatherInfo(numberOfUpdates.Text));
                Lista.Items.Add(GetCityWeatherInfo(numberOfUpdates.Text).Name+"\t\t");
            }
            else
            {
                numberOfUpdates.Text = "Invalid city name";
                
            }
        }
        /// <summary>
        /// Checks weather every declared time
        /// </summary>
        private void delayLoop()
        {
            int i = 2;
            
            for (int z = 1; z <= defTime; z++)
            {
                Dispatcher.Invoke(() => Lista.Items.Clear());
                Dispatcher.Invoke(() => numberOfUpdates.Text = $"Number of updates: {z} from {defTime}");
                Dispatcher.Invoke(() => timeCounter.Text = $"Last update: {DateTime.Now.ToString("HH:mm")}");
                i = CitiesToExcelLoop(i, cityWeathersUpToDate);
                System.Threading.Thread.Sleep(600000);

            }
            Dispatcher.Invoke(() => numberOfUpdates.Text = "Finish"); // End of program
        }

        private void timerCounter(int i)
        {
            numberOfUpdates.Text = $"Pozostało {240 - i} minut";
        }

        private static List<CityWeather> citiesListWeatherUpdate()
        {
            List<CityWeather> cityWeathers = new List<CityWeather>(); // Lista miast do pobierania pogody
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
                    numberOfUpdates.Text = "Niepoprawna nazwa miasta";
                }
            }

            return i;
        }

        private void writeCityWeatherToFile(int i, CityWeather city)
        {
            cityTempToList(city);

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

        private void cityTempToList(CityWeather city)
        {
            if (city.TempC < 0)
            { Dispatcher.Invoke(() => Lista.Items.Add($"{city.TempC} degrees Celsius in {city.Name} \t\t")); }
            else
            { Dispatcher.Invoke(() => Lista.Items.Add($" {city.TempC} degrees Celsius in {city.Name} \t\t")); }
        }
        #endregion

        private void ListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }


        private void TextBox_TextChanged2(object sender, TextChangedEventArgs e)
        {

        }

        private void Lista_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void timeCounter_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void button_Add_Click(object sender, RoutedEventArgs e)
        {

        }

        
    }
}
