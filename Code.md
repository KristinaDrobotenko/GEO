# GEO
Сode for the program "Geolocation of users"
*******************************************

//Код формы Major.cs:
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GMap.NET;
using GMap.NET.WindowsForms;
//Пространство имен для работы с картой
using GMap.NET.WindowsForms.ToolTips;
using GMap.NET.WindowsForms.Markers;
//Пространсво имен для работа с XML
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.XPath;
using System.IO;
//Представляет постоянное регулярное выражение
using System.Text.RegularExpressions;
//Пространство имен для работы со стилем формы
using MetroFramework.Components;
using MetroFramework.Forms;

namespace mapsgto
{
    public partial class Major : MetroForm
    {
        //Список для хранения пути к изображению.
        public List<string> PathImage;
        //Переменная путь к XML файлу 
        public static string infoxml = Application.StartupPath + "\\GeoXML.xml";
        // Создание документа DOM и загрузка данных XML в него
        public XDocument dom = XDocument.Load(infoxml);
        //Список маркеров-меток
        public GMapOverlay markersOverlay;
        //Начальное значение широты и долготы
        public double lat = 0, lng = 0;      
        public Major()
        {
            InitializeComponent();
            //Цвет формы
            MetroStyleManager.Default.Style = MetroFramework.MetroColorStyle.Orange;
            //Тема формы
            MetroStyleManager.Default.Theme = MetroFramework.MetroThemeStyle.Default;
        }
        private void Major_Load(object sender, EventArgs e)
        {
            //Настройки для компонента GMap
            gMapBKO.Bearing = 0;
            //CanDragMap - Если параметр установлен в True, пользователь может перетаскивать карту с попмощью левой кнопки мыши
            gMapBKO.CanDragMap = true;
            //Указываем, что перетаскивание карты осуществляется с использованием левой клавиши мыши.
            gMapBKO.DragButton = MouseButtons.Left;
            gMapBKO.GrayScaleMode = true;
            //MarkersEnsbled - Если параметр установлен в True, любые маркеры, заданные в ручную будут показаны. Если нет, они не появятся.
            gMapBKO.MarkersEnabled = true;
            //Указываем значение максимального приближения.
            gMapBKO.MaxZoom = 16;
            //Указываем значение минимального приближения.
            gMapBKO.MinZoom = 2;
            //Устанавливаем центр приближения/удаления курсор мыши.
            gMapBKO.MouseWheelZoomType = GMap.NET.MouseWheelZoomType.MousePositionAndCenter;
            //Отказываемся от негативного режима.
            gMapBKO.NegativeMode = false;
            //Разрешаем полигоны
            gMapBKO.PolygonsEnabled = true;
            //Разрешаем маршруты
            gMapBKO.RoutesEnabled = true;
            //Скрываем внешнюю сетку карты с заголовками.
            gMapBKO.ShowTileGridLines = false;
            //Указываем, что при загрузки карты будет использоваться 16-ти кратное приближение.
            gMapBKO.Zoom = 16;
            //Указываем, что все края элемента управления закрепляются у краев содержащего его элементы управления(главной формой), а их размеры изменяются соответствующим образом.
            gMapBKO.Dock = DockStyle.Fill;
            //Указывваем, что будем использовать карты YANDEX.       
            gMapBKO.MapProvider = GMap.NET.MapProviders.GMapProviders.YandexMap;
            GMap.NET.GMaps.Instance.Mode = GMap.NET.AccessMode.ServerOnly;
            //Указываем местонахождения АО "БКО"
            gMapBKO.Position = new GMap.NET.PointLatLng(58.387815, 33.865559);
            //Заполнение списка отделов в comboBox
            FillComboBox(cBDepartForSearch);
            //Заполнение данных в statusStrip
            FillStatusStrip();
            //Заполнение списка рабочих мест в comboBox
            FillCbRM();
        } 
        private void FillCbRM()//Метод на заполнение списка рабочих мест в comboBox
        {
            //Переменная для загрузки данных из XML файла
            dom = XDocument.Load(infoxml);
            //Очищение списка в comboBox
            cBRM.Items.Clear();
            //Путь к узлу РМ в XML файле GeoXML
            IEnumerable<XElement> element = dom.XPathSelectElements("/Отделы/Отдел/Подотдел/РМ");
            //Перебор всех элементов в файле GeoXML
            foreach (XElement el in element)
            {
                //Проверка того, что атрибут Name неотсутствует и не пустой
                if (el.Attribute("Name") != null && el.Attribute("Name").Value != "")
                    //Загрузка всех элементов с атрибутом Name в comboBox
                    cBRM.Items.Add(el.Attribute("Name").Value.Trim());
            }
        }
        private void FillStatusStrip()//Метод на заполнение списка рабочих мест в comboBox
        {   //Путь к узлу Отдел в XML файле GeoXML
            IEnumerable<XElement> elementsOtdel = dom.XPathSelectElements("/Отделы/Отдел");
            //Путь к узлу Подотдел в XML файле GeoXML
            IEnumerable<XElement> elementsPodotdel = dom.XPathSelectElements("/Отделы/Отдел/Подотдел");
            //Путь к узлу РМ в XML файле GeoXML
            IEnumerable<XElement> elementsRM = dom.XPathSelectElements("/Отделы/Отдел/Подотдел/РМ");
            //Путь к узлу Пользователь в XML файле GeoXML
            IEnumerable<XElement> elementsUsers = dom.XPathSelectElements("/Отделы/Отдел/Подотдел/РМ/Пользователь");
            //Подсчет отделов в XML файле GeoXML и вывод их количества в statusStrip
            tStrStLabDepart.Text = "Отделов: " + elementsOtdel.Count().ToString();
            //Подсчет подразделений в XML файле GeoXML и вывод их количества в statusStrip
            tStrStaLabSubdiv.Text = "Подразделений: " + elementsPodotdel.Count().ToString();
            //Подсчет рабочих мест в XML файле GeoXML и вывод их количества в statusStrip
            tStrStaLabRM.Text = "Рабочих мест: " + elementsRM.Count().ToString();
            //Подсчет пользователей в XML файле GeoXML и вывод их количества в statusStrip
            tStrStaLabUser.Text = "Пользователей: " + elementsUsers.Count().ToString();
        }
        private void gMapBKO_Load(object sender, EventArgs e)//Загрузка карты
        {
            MAP();//Загрузка карты
        } 
        public void MAP()//Метод загрузки карты
        {
            //Переменная путь к XML файлу 
            string uri = Application.StartupPath + @"\GeoXML.xml";
            //Переменная для загрузки документа
            XmlDocument xml = new XmlDocument();
            //Загрузка документа
            xml.Load(uri);
            //Переменная для загрузки данных из XML файла
            XmlNodeList xnList = xml.SelectNodes("/Отделы/Отдел");
            //Создание нового списка маркеров-меток
            markersOverlay = new GMapOverlay("marker");
            //Перебор всех элментов в XML файле GeoXML в узле Отдел
            foreach (XmlNode xn in xnList)
            {
                //Проверка того, что Атрибуты lat и lng не пустые
                if (xn.Attributes["lat"] != null && xn.Attributes["lng"] != null)
                {
                    //объявление переменной lat, атрибут для которой берется из файла GeoXML
                    double lat = Convert.ToDouble(xn.Attributes["lat"].Value);
                    //объявление переменной lng, атрибут для которой берется из файла GeoXML
                    double lng = Convert.ToDouble(xn.Attributes["lng"].Value);
                    //объявление переменной name, атрибут для которой берется из файла GeoXML
                    string name = xn.Attributes["Name"].Value;
                    //объявление переменной address, атрибут для которой берется из файла GeoXML
                    string address = xn.Attributes["adress"].Value;
                    //Инициилизация маркеров, с указанием их координат  
                    GMarkerGoogle markerG = new GMarkerGoogle(new PointLatLng(lat, lng), GMarkerGoogleType.blue_dot);
                    //Текст отображаемый при наведении на маркер
                    markerG.ToolTipText = name + Environment.NewLine + address;
                    //label1.Text += markerG.ToolTipText;
                    //Добавляем маркер в список маркеров 
                    markersOverlay.Markers.Add(markerG);
                }
            }
            //Добавляем в компонент, список маркеров
            gMapBKO.Overlays.Add(markersOverlay);
        }
        public void FillComboBox(ComboBox cmbDevelop)//Метод загрузки списка отделов в comboBox
        {
            try
            {
                //Путь к узлу Отдел в XML файле GeoXML
                IEnumerable<XElement> elements = dom.XPathSelectElements("/Отделы/Отдел");
                //Перебор всех элментов в XML файле GeoXML в узле Отдел
                foreach (XElement el in elements)
                {
                    //Проверка того, что атрибут Name не отсутствует и не пустой
                    if (el.Attribute("Name") != null && el.Attribute("Name").Value != "")                    
                        //Загрузка всех элементов с атрибутом Name в comboBox
                        cBDepartForSearch.Items.Add(el.Attribute("Name").Value.Trim());
                    
                }
            }
            catch { }
        }
        private void btShowOnMap_Click(object sender, EventArgs e)//Обработчик на кнопку "Найти" для поиска подразделения 
        {
            //Проверка, если в cBDepartForSearch нет текста, то вывести соответствующее сообщение
            if (cBDepartForSearch.Text == "")
            {
                MessageBox.Show("Выберите отдел и подразделение для того, чтобы выполнить поиск", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else
            {
                try
                {
                    //Переменная для выбранного элемента в cBSubdivForSearch
                    string name = Convert.ToString(cBSubdivForSearch.SelectedItem);
                    //Переменные широты и долготы 
                    double lat = 0, lng = 0;
                    //Путь к узлу Подотдел в XML файле GeoXML
                    IEnumerable<XElement> elements = dom.XPathSelectElements("/Отделы/Отдел/Подотдел");
                    //Перебор всех элментов в XML файле GeoXML в узле Подотдел
                    foreach (XElement el in elements)
                    {
                       //объявлнение списка маркеров - меток
                        GMapOverlay markersOverlay = new GMapOverlay("marker");
                        //Проверка того, что атрибут Name неотсутствует и не пустой                      
                        if (el.Attribute("Name").Value.Trim() == name)
                        {
                            //Присваивание переменным lat и lng значение соответствующих атребутов
                            lat = Convert.ToDouble(el.Attribute("lat").Value.Trim());
                            lng = Convert.ToDouble(el.Attribute("lng").Value.Trim());
                            //Инициилизация маркеров, с указанием их координат  
                            GMarkerGoogle markerG = new GMarkerGoogle(new PointLatLng(lat, lng), GMarkerGoogleType.green_dot);
                            //Текст отображаемый при наведении на маркер
                            markerG.ToolTipText = name;
                            //Добавляем маркер в список маркеров 
                            markersOverlay.Markers.Add(markerG);
                            //Добавляем в компонент, список маркеров
                            gMapBKO.Overlays.Add(markersOverlay);
                            break;
                        }
                    }
                    //Очищение comboBox'ов 
                    cBSubdivForSearch.Text = "";
                    cBDepartForSearch.Text = "";
                    //Зум карты при отображнии
                    gMapBKO.Zoom = 16;
                    //Позиция карты соответствует широте и долготе
                    gMapBKO.Position = new GMap.NET.PointLatLng(lat, lng);
                }
                catch
                { }
            }            
        }
        private void cBDepartForSearch_SelectedIndexChanged(object sender, EventArgs e)
        {
            //включаем элемент cBSubdivForSearch
            cBSubdivForSearch.Enabled = true;
            //Очистка списока cBSubdivForSearch
            cBSubdivForSearch.Items.Clear();
            //Текст в cBSubdivForSearch отсутствует
            cBSubdivForSearch.Text = "";
            //Путь к узлу Подотдел в XML файле GeoXML
            IEnumerable<XElement> elements = dom.XPathSelectElements("/Отделы/Отдел/Подотдел");
            //Перебор всех элментов в XML файле GeoXML в узле Подотдел
            foreach (XElement el in elements)
            {
                //Проверка: равен ли атрибут Name выбранному из списка названию отдела
                if (el.Parent.Attribute("Name").Value.Trim() == Convert.ToString(cBDepartForSearch.SelectedItem))
                    //Загрузка элемента с атрибутом Name в cBSubdivForSearch
                    cBSubdivForSearch.Items.Add(el.Attribute("Name").Value.Trim());
            }
        }
        private void цСПToStMI_Click(object sender, EventArgs e)//Отображение  отдела
        {
            //Присваивание переменной name значения цСПToStMI
            string name = Convert.ToString(цСПToStMI.Text);
            //Переменные широты и долготы 
            double lat = 0, lng = 0;
            //Путь к узлу Отдел в XML файле GeoXML
            IEnumerable<XElement> elements = dom.XPathSelectElements("/Отделы/Отдел");
            //Перебор всех элментов в XML файле GeoXML в узле Отдел
            foreach (XElement el in elements)
            {
                //Проверка равен ли атрибут перменной name
                if (el.Attribute("Name").Value.Trim() == name)
                {
                    //Присваивание переменным lat и lng значений атрибутов lat и lng соответствующего отдела
                    lat = Convert.ToDouble(el.Attribute("lat").Value.Trim());
                    lng = Convert.ToDouble(el.Attribute("lng").Value.Trim());
                    break;
                }
            }
            //Зум карты при отображнии
            gMapBKO.Zoom = 16;
            //Позиция карты соответствует широте и долготе
            gMapBKO.Position = new GMap.NET.PointLatLng(lat, lng);
        }
        private void АдмToStMI_Click(object sender, EventArgs e)//Отображение  отдела
        {
            //Присваивание переменной name значения АдмToStMI
            string name = Convert.ToString(АдмToStMI.Text);
            //Переменные широты и долготы 
            double lat = 0, lng = 0;
            //Путь к узлу Отдел в XML файле GeoXML
            IEnumerable<XElement> elements = dom.XPathSelectElements("/Отделы/Отдел");
            //Перебор всех элментов в XML файле GeoXML в узле Отдел
            foreach (XElement el in elements)
            {
                //Проверка равен ли атрибут перменной name
                if (el.Attribute("Name").Value.Trim() == name)
                {
                    //Присваивание переменным lat и lng значений атрибутов lat и lng соответствующего отдела
                    lat = Convert.ToDouble(el.Attribute("lat").Value.Trim());
                    lng = Convert.ToDouble(el.Attribute("lng").Value.Trim());
                    break;
                }
            }
            //Зум карты при отображнии
            gMapBKO.Zoom = 16;
            //Позиция карты соответствует широте и долготе
            gMapBKO.Position = new GMap.NET.PointLatLng(lat, lng);
        }
        private void эНЕЦЕХToStMI_Click(object sender, EventArgs e)//Отображение  отдела
        {
            //Присваивание переменной name значения эНЕЦЕХToStMI
            string name = Convert.ToString(эНЕЦЕХToStMI.Text);
            //Переменные широты и долготы 
            double lat = 0, lng = 0;
            //Путь к узлу Отдел в XML файле GeoXML
            IEnumerable<XElement> elements = dom.XPathSelectElements("/Отделы/Отдел");
            //Перебор всех элментов в XML файле GeoXML в узле Отдел
            foreach (XElement el in elements)
            {
                //Проверка равен ли атрибут перменной name
                if (el.Attribute("Name").Value.Trim() == name)
                {
                    //Присваивание переменным lat и lng значений атрибутов lat и lng соответствующего отдела
                    lat = Convert.ToDouble(el.Attribute("lat").Value.Trim());
                    lng = Convert.ToDouble(el.Attribute("lng").Value.Trim());
                    break;
                }
            }
            //Зум карты при отображнии
            gMapBKO.Zoom = 16;
            //Позиция карты соответствует широте и долготе
            gMapBKO.Position = new GMap.NET.PointLatLng(lat, lng);
        }
        private void фОКToStMI_Click(object sender, EventArgs e)//Отображение  отдела
        {
            //Присваивание переменной name значения фОКToStMI
            string name = Convert.ToString(фОКToStMI.Text);
            //Переменные широты и долготы 
            double lat = 0, lng = 0;
            //Путь к узлу Отдел в XML файле GeoXML
            IEnumerable<XElement> elements = dom.XPathSelectElements("/Отделы/Отдел");
            //Перебор всех элментов в XML файле GeoXML в узле Отдел
            foreach (XElement el in elements)
            {
                //Проверка равен ли атрибут перменной name
                if (el.Attribute("Name").Value.Trim() == name)
                {
                    //Присваивание переменным lat и lng значений атрибутов lat и lng соответствующего отдела
                    lat = Convert.ToDouble(el.Attribute("lat").Value.Trim());
                    lng = Convert.ToDouble(el.Attribute("lng").Value.Trim());
                    break;
                }
            }
            //Зум карты при отображнии
            gMapBKO.Zoom = 16;
            //Позиция карты соответствует широте и долготе
            gMapBKO.Position = new GMap.NET.PointLatLng(lat, lng);
        }
        private void торгоЦToStMI_Click(object sender, EventArgs e)//Отображение  отдела
        {
            //Присваивание переменной name значения торгоЦToStMI
            string name = Convert.ToString(торгоЦToStMI.Text);
            //Переменные широты и долготы 
            double lat = 0, lng = 0;
            //Путь к узлу Отдел в XML файле GeoXML
            IEnumerable<XElement> elements = dom.XPathSelectElements("/Отделы/Отдел");
            //Перебор всех элментов в XML файле GeoXML в узле Отдел
            foreach (XElement el in elements)
            {
                //Проверка равен ли атрибут перменной name
                if (el.Attribute("Name").Value.Trim() == name)
                {
                    //Присваивание переменным lat и lng значений атрибутов lat и lng соответствующего отдела
                    lat = Convert.ToDouble(el.Attribute("lat").Value.Trim());
                    lng = Convert.ToDouble(el.Attribute("lng").Value.Trim());
                    break;
                }
            }
            //Зум карты при отображнии
            gMapBKO.Zoom = 16;
            //Позиция карты соответствует широте и долготе
            gMapBKO.Position = new GMap.NET.PointLatLng(lat, lng);
        }
        private void торгйДомTlStMI_Click(object sender, EventArgs e)//Отображение  отдела
        {
            //Присваивание переменной name значения торгйДомTlStMItem
            string name = Convert.ToString(торгйДомTlStMItem.Text);
            //Переменные широты и долготы 
            double lat = 0, lng = 0;
            //Путь к узлу Отдел в XML файле GeoXML
            IEnumerable<XElement> elements = dom.XPathSelectElements("/Отделы/Отдел");
            //Перебор всех элментов в XML файле GeoXML в узле Отдел
            foreach (XElement el in elements)
            {
                //Проверка равен ли атрибут перменной name
                if (el.Attribute("Name").Value.Trim() == name)
                {
                    //Присваивание переменным lat и lng значений атрибутов lat и lng соответствующего отдела
                    lat = Convert.ToDouble(el.Attribute("lat").Value.Trim());
                    lng = Convert.ToDouble(el.Attribute("lng").Value.Trim());
                    break;
                }
            }
            //Зум карты при отображнии
            gMapBKO.Zoom = 16;
            //Позиция карты соответствует широте и долготе
            gMapBKO.Position = new GMap.NET.PointLatLng(lat, lng);
        }
        private void мСЧToStMI_Click(object sender, EventArgs e)
        {
            //Присваивание переменной name значения мСЧToStMI
            string name = Convert.ToString(мСЧToStMI.Text);
            //Переменные широты и долготы 
            double lat = 0, lng = 0;
            //Путь к узлу Отдел в XML файле GeoXML
            IEnumerable<XElement> elements = dom.XPathSelectElements("/Отделы/Отдел");
            //Перебор всех элментов в XML файле GeoXML в узле Отдел
            foreach (XElement el in elements)
            {
                //Проверка равен ли атрибут перменной name
                if (el.Attribute("Name").Value.Trim() == name)
                {
                    //Присваивание переменным lat и lng значений атрибутов lat и lng соответствующего отдела
                    lat = Convert.ToDouble(el.Attribute("lat").Value.Trim());
                    lng = Convert.ToDouble(el.Attribute("lng").Value.Trim());
                    break;
                }
            }
            //Зум карты при отображнии
            gMapBKO.Zoom = 16;
            //Позиция карты соответствует широте и долготе
            gMapBKO.Position = new GMap.NET.PointLatLng(lat, lng);
        }
        private void btnShowRMonMap_Click(object sender, EventArgs e)//Обработчик на отображение рм
        {
            //Проверка: если cBRM пуст, то вывести сообщение
            if (cBRM.Text == "")
            {
                MessageBox.Show("Выберите рабочее место для отображения на карте", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else
            {
                try
                {
                    //Присваивание переменной значения текста из cBRM
                    string name = Convert.ToString(cBRM.SelectedItem);
                    //Переменные широты и долготы 
                    double lat = 0, lng = 0;
                    //Путь к узлу РМ в XML файле GeoXML
                    IEnumerable<XElement> elements = dom.XPathSelectElements("/Отделы/Отдел/Подотдел/РМ");
                    //Перебор всех элментов в XML файле GeoXML в узле РМ
                    foreach (XElement el in elements)
                    {
                        //Создание нового списка маркеров
                        GMapOverlay markersOverlay = new GMapOverlay("marker");
                        //Проверка: равен ли атрибут Name переменной name
                        if (el.Attribute("Name").Value.Trim() == name)
                        {
                            //Присваивание lat и lng значений из моответствующих атрибутов
                            lat = Convert.ToDouble(el.Attribute("lat").Value.Trim());
                            lng = Convert.ToDouble(el.Attribute("lng").Value.Trim());

                            //Инициилизация маркеров, с указанием их координат  
                            GMarkerGoogle markerG = new GMarkerGoogle(new PointLatLng(lat, lng), GMarkerGoogleType.orange_dot);
                            //Текст отображаемый при наведении на маркер
                            markerG.ToolTipText = name;
                            //Добавляем маркер в список маркеров 
                            markersOverlay.Markers.Add(markerG);
                            //Добавляем в компонент, список маркеров
                            gMapBKO.Overlays.Add(markersOverlay);
                            break;
                        }
                    }
                    //Очищение списка cBRM
                    cBRM.Text = "";
                    //Зум карты при отображении
                    gMapBKO.Zoom = 16;
                    //Позиция карты соответствует широте и долготе
                    gMapBKO.Position = new GMap.NET.PointLatLng(lat, lng);
                }
                catch
                { }
            }
        }
        private void btnInfoAboutRM_Click(object sender, EventArgs e)//Переход на форму информация
        {
            //Проверка: если cBRM пуст, то вывести сообщение
            if (cBRM.Text == "")
            {
                MessageBox.Show("Для просмотра информации выберите рабочее место", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else
            {
                InformationAboutMarkers InfAbMarkers = new InformationAboutMarkers();//Создание новой переменной формы InformationAboutMarkers
                InfAbMarkers.labInfo.Text = cBRM.Text;//Присваивание labInfo текста из cBRM
                InfAbMarkers.ShowDialog();//Демонстрация формы
            }
        }
        private void AddDepartToStMI_Click(object sender, EventArgs e)//Переход на форму добавления отдела
        {     
            DobavlenieOtdela addotdel = new DobavlenieOtdela();//Создание новой переменной формы DobavleniePodotdela
            addotdel.lat = lat;//Передача данных lat в переменную addotdel
            addotdel.lng = lng;//Передача данных lng в переменную addotdel
            addotdel.ShowDialog(); //Демонстрация формы
        }
        private void btnSearh_Click(object sender, EventArgs e)//Поиск сотрудника
        {
            if (tBNayti.Text == "")//Если tBNayti пустой, вывод сообщения
            {
                MessageBox.Show("Введите информацию для поиска", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else//Если tBNayti не пустой
            {
                try
                {   //Создание нового списка маркеров-меток
                    GMapOverlay markersOverlay = new GMapOverlay("marker");
                    //Присваивание переменной name значение текста из tBNayti
                    string name = tBNayti.Text.Trim();
                    //Переменная, которой присваивается значение из cBDataSearch
                    string selitem = cBDataSearch.SelectedItem.ToString();
                    //Переменные широты и долготы 
                    double lat = 0, lng = 0;
                    //Переменная для загрузки документа
                    XmlDocument doc = new XmlDocument();
                    //Загрузка докуменита в переменную doc
                    doc.Load(infoxml);
                    //Организация выбора
                    switch (selitem)
                    {
                        case "ФИО"://Поиск по ФИО
                            //Путь к узлу Пользователь в XML файле GeoXML
                            IEnumerable<XElement> elementsF = dom.XPathSelectElements("/Отделы/Отдел/Подотдел/РМ/Пользователь");
                            //Перебор всех элментов в XML файле GeoXML в узле Пользователь
                            foreach (XElement xn in elementsF)
                            {
                                Match match = Regex.Match(Convert.ToString(xn.Attribute(selitem).Value.Trim()), name, RegexOptions.IgnoreCase);
                                if (match.Success)
                                {
                                    if (xn.Ancestors("Подотдел").Single().Attribute("lat").Value.Trim() != String.Empty || xn.Ancestors("Подотдел").Single().Attribute("lng").Value.Trim() != String.Empty)
                                    {
                                        //Присваивание переменным lat и lng значений атрибутов lat и lng соответствующего подотдела
                                        lat = Convert.ToDouble(xn.Ancestors("Подотдел").Single().Attribute("lat").Value.Trim());
                                        lng = Convert.ToDouble(xn.Ancestors("Подотдел").Single().Attribute("lng").Value.Trim());
                                    }
                                    else//иначе присваиваем значения отдела
                                    {
                                        //Присваивание переменным lat и lng значений атрибутов lat и lng соответствующего отдела
                                        lat = Convert.ToDouble(xn.Ancestors("Отдел").Single().Attribute("lat").Value.Trim());
                                        lng = Convert.ToDouble(xn.Ancestors("Отдел").Single().Attribute("lng").Value.Trim());
                                    }
                                    //Инициилизация маркеров, с указанием их координат  
                                    GMarkerGoogle markerG = new GMarkerGoogle(new PointLatLng(lat, lng), GMarkerGoogleType.orange_dot);
                                    //Текст отображаемый при наведении на маркер
                                    markerG.ToolTipText = name;
                                    //Добавляем маркер в список маркеров 
                                    markersOverlay.Markers.Add(markerG);
                                    //Добавляем в компонент, список маркеров
                                    gMapBKO.Overlays.Add(markersOverlay);
                                    break;
                                }
                            }
                            break;


                        case "РМ"://Поиск по рабочему месту
                            //Путь к узлу РМ в XML файле GeoXML
                            IEnumerable<XElement> elementsR = dom.XPathSelectElements("/Отделы/Отдел/Подотдел/РМ");
                            //Перебор всех элментов в XML файле GeoXML в узле РМ
                            foreach (XElement xn in elementsR)
                            {
                                Match match = Regex.Match(Convert.ToString(xn.Attribute("Name").Value.Trim()), name, RegexOptions.IgnoreCase);
                                if (match.Success)
                                {
                                    if (xn.Ancestors("Подотдел").Single().Attribute("lat").Value.Trim() != String.Empty || xn.Ancestors("Подотдел").Single().Attribute("lng").Value.Trim() != String.Empty)
                                    {
                                        //Присваивание переменным lat и lng значений атрибутов lat и lng соответствующего подотдела
                                        lat = Convert.ToDouble(xn.Ancestors("Подотдел").Single().Attribute("lat").Value.Trim());
                                        lng = Convert.ToDouble(xn.Ancestors("Подотдел").Single().Attribute("lng").Value.Trim());
                                    }
                                    else
                                    {
                                        //Присваивание переменным lat и lng значений атрибутов lat и lng соответствующего отдела
                                        lat = Convert.ToDouble(xn.Ancestors("Отдел").Single().Attribute("lat").Value.Trim());
                                        lng = Convert.ToDouble(xn.Ancestors("Отдел").Single().Attribute("lng").Value.Trim());
                                    }
                                    //Инициилизация маркеров, с указанием их координат  
                                    GMarkerGoogle markerG = new GMarkerGoogle(new PointLatLng(lat, lng), GMarkerGoogleType.orange_dot);
                                    //Текст отображаемый при наведении на маркер
                                    markerG.ToolTipText = name;
                                    //Добавляем маркер в список маркеров 
                                    markersOverlay.Markers.Add(markerG);
                                    //Добавляем в компонент, список маркеров
                                    gMapBKO.Overlays.Add(markersOverlay);
                                    break;
                                }
                            }
                            break;

                        case "Сетевое имя"://Поиск по сетевому имени
                            //Путь к узлу Пользователь в XML файле GeoXML
                            IEnumerable<XElement> elementsS = dom.XPathSelectElements("/Отделы/Отдел/Подотдел/РМ/Пользователь");
                            //Перебор всех элментов в XML файле GeoXML в узле Пользователь
                            foreach (XElement xn in elementsS)
                            {
                                Match match = Regex.Match(Convert.ToString(xn.Attribute("Сетевое_имя").Value.Trim()), name, RegexOptions.IgnoreCase);
                                if (match.Success)
                                {
                                    if (xn.Ancestors("Подотдел").Single().Attribute("lat").Value.Trim() != String.Empty || xn.Ancestors("Подотдел").Single().Attribute("lng").Value.Trim() != String.Empty)
                                    {
                                        //Присваивание переменным lat и lng значений атрибутов lat и lng соответствующего подотдела
                                        lat = Convert.ToDouble(xn.Ancestors("Подотдел").Single().Attribute("lat").Value.Trim());
                                        lng = Convert.ToDouble(xn.Ancestors("Подотдел").Single().Attribute("lng").Value.Trim());
                                    }
                                    else
                                    {
                                        //Присваивание переменным lat и lng значений атрибутов lat и lng соответствующего отдела
                                        lat = Convert.ToDouble(xn.Ancestors("Отдел").Single().Attribute("lat").Value.Trim());
                                        lng = Convert.ToDouble(xn.Ancestors("Отдел").Single().Attribute("lng").Value.Trim());
                                    }
                                    //Инициилизация маркеров, с указанием их координат  
                                    GMarkerGoogle markerG = new GMarkerGoogle(new PointLatLng(lat, lng), GMarkerGoogleType.orange_dot);
                                    //Текст отображаемый при наведении на маркер
                                    markerG.ToolTipText = name;
                                    //Добавляем маркер в список маркеров 
                                    markersOverlay.Markers.Add(markerG);
                                    //Добавляем в компонент, список маркеров
                                    gMapBKO.Overlays.Add(markersOverlay);
                                    break;
                                }
                            }
                            break;
                    }
                    //Если lat == 0, то вывести сообщение
                    if (lat == 0) { MessageBox.Show("Не найдено!"); lat = 58.387815; lng = 33.865559; }
                    //Зум карты при отображнии                    
                    gMapBKO.Zoom = 16;
                    //Позиция карты соответствует широте и долготе
                    gMapBKO.Position = new GMap.NET.PointLatLng(lat, lng);
                }
                catch (Exception) { }
            }
        }
        private void gMapBKO_MouseMove(object sender, MouseEventArgs e)
        {
            lat = gMapBKO.FromLocalToLatLng(e.X, e.Y).Lat;
            lng = gMapBKO.FromLocalToLatLng(e.X, e.Y).Lng;
        }
        private void AddSubdivToStMI_Click(object sender, EventArgs e)//Переход на форму добавления подразделения
        {
            DobavleniePodotdela addotdel = new DobavleniePodotdela();//Создание новой переменной формы DobavleniePodotdela
            addotdel.lat = lat;//Передача данных lat в переменную addotdel
            addotdel.lng = lng;//Передача данных lng в переменную addotdel
            addotdel.ShowDialog();//Демонстрация формы
        }
        private void AddRMToStMI_Click(object sender, EventArgs e)//Переход на форму добавления рабочего места
        {
            DobavlenieRM addotdel = new DobavlenieRM();//Создание новой переменной формы DobavleniePodotdela
            addotdel.lat = lat;//Передача данных lat в переменную addotdel
            addotdel.lng = lng;//Передача данных lng в переменную addotdel
            addotdel.ShowDialog();//Демонстрация формы
        }
        private void btnApply_Click(object sender, EventArgs e)//Фильтрация
        {
            //Проверка, если нет галочки на chkBDepart/chkBSubdiv/chkBRM вывести сообщение
            if (chkBDepart.Checked == false & chkBSubdiv.Checked == false & chkBRM.Checked == false)
            {
                MessageBox.Show("Выберите ограниччения для фильтрации маркеров - меток", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else
            {
                //Переменная для загрузки документа
                XmlDocument xml = new XmlDocument();
                //Загрузка документа
                xml.Load(infoxml);
                //Добавляем в компонент, список маркеров
                XmlNodeList xnList;
                //Очищаем карту от маркеров;
                markersOverlay.Markers.Clear();
                //Очищаем список маркеров;
                gMapBKO.Overlays.Clear();
                //Проверка включен ли чекбокс, отвечающий за отдел
                if (chkBDepart.Checked == true)
                {
                    //Перечисляем список отделов
                    xnList = xml.SelectNodes("/Отделы/Отдел");
                    //Перебор всех элментов в XML файле GeoXML в узле Отдел
                    foreach (XmlNode xn in xnList)
                    {
                        try
                        {
                            //Проверяем вбиты ли координаты
                            if (xn.Attributes["lat"].Value != String.Empty || xn.Attributes["lng"].Value != String.Empty)
                            {
                                //Присваивание переменным lat и lng значений из соответствующих атрибутов
                                double lat = Convert.ToDouble(xn.Attributes["lat"].Value);
                                double lng = Convert.ToDouble(xn.Attributes["lng"].Value);
                                //Присваивание переменной name значения из оответствующего атрибута
                                string name = xn.Attributes["Name"].Value;
                                //Присваивание переменной adress значения из оответствующего атрибута
                                string address = xn.Attributes["adress"].Value;
                                //Инициилизация маркеров, с указанием их координат  
                                GMarkerGoogle markerG = new GMarkerGoogle(new PointLatLng(lat, lng), GMarkerGoogleType.blue_dot);
                                //Текст отображаемый при наведении на маркер
                                markerG.ToolTipText = name + Environment.NewLine + address;
                                //Добавляем маркер в список маркеров 
                                markersOverlay.Markers.Add(markerG);
                                //Отображение маркеров
                                gMapBKO.UpdateMarkerLocalPosition(markerG);
                            }
                        }
                        catch
                        { }
                    }
                }

                //Проверка включен ли чекбокс, отвечающий за подразделение
                if (chkBSubdiv.Checked == true)
                {
                    //Перечисляем список подотделов
                    xnList = xml.SelectNodes("/Отделы/Отдел/Подотдел");
                    foreach (XmlNode xn in xnList)
                    {
                        try
                        {
                            //Проверяем вбиты ли координаты
                            if (xn.Attributes["lat"].Value != String.Empty || xn.Attributes["lng"].Value != String.Empty)
                            {
                                //Присваивание переменным lat и lng значений из соответствующих атрибутов
                                double lat = Convert.ToDouble(xn.Attributes["lat"].Value);
                                double lng = Convert.ToDouble(xn.Attributes["lng"].Value);
                                //Присваивание переменной name значения из оответствующего атрибута
                                string name = xn.Attributes["Name"].Value;
                                //Инициилизация маркеров, с указанием их координат  
                                GMarkerGoogle markerG = new GMarkerGoogle(new PointLatLng(lat, lng), GMarkerGoogleType.green_dot);
                                //Текст отображаемый при наведении на маркер
                                markerG.ToolTipText = name + Environment.NewLine;
                                //Добавляем маркер в список маркеров 
                                markersOverlay.Markers.Add(markerG);
                                gMapBKO.UpdateMarkerLocalPosition(markerG);
                            }
                        }
                        catch
                        { }
                    }
                }
                //Проверка включен ли чекбокс, отвечающий за РМ
                if (chkBRM.Checked == true)
                {
                    //Перечисляем список подотделов
                    xnList = xml.SelectNodes("/Отделы/Отдел/Подотдел");
                    foreach (XmlNode xn in xnList)
                    {
                        try
                        {
                            //Проверяем вбиты ли координаты, и принадлежность к тому или иному подотделу, если да, то берем координаты подотдела и ставим маркер
                            if (xn.Attributes["lat"].Value != String.Empty || xn.Attributes["lng"].Value != String.Empty)
                            {
                                //Перечисляем список рабочих мест принадлежащих этому подотделу
                                XmlNodeList xnSubdivision = xml.SelectNodes("/Отделы/Отдел/Подотдел[@Name='" + xn.Attributes["Name"].Value + "']/РМ");
                                //Прогоняем каждое рабочее место и добавляем маркер
                                foreach (XmlNode sn in xnSubdivision)
                                {
                                    //Присваивание переменным lat и lng значений из соответствующих атрибутов
                                    double lat = Convert.ToDouble(xn.Attributes["lat"].Value);
                                    double lng = Convert.ToDouble(xn.Attributes["lng"].Value);
                                    //Присваивание переменной name значения из оответствующего атрибута
                                    string name = sn.Attributes["Name"].Value;
                                    //Инициилизация маркеров, с указанием их координат  
                                    GMarkerGoogle markerG = new GMarkerGoogle(new PointLatLng(lat, lng), GMarkerGoogleType.orange_dot);
                                    //Текст отображаемый при наведении на маркер
                                    markerG.ToolTipText = name + Environment.NewLine + xn.Attributes["Name"].Value + Environment.NewLine;
                                    //Добавляем маркер в список маркеров 
                                    markersOverlay.Markers.Add(markerG);
                                    gMapBKO.UpdateMarkerLocalPosition(markerG);
                                }
                            }
                        }
                        catch
                        { }
                    }
                }      
                //Загрузка в список маркеров
                gMapBKO.Overlays.Add(markersOverlay);
                //Зум карты при отображнии 
                gMapBKO.Zoom = 16;
            }  
        }
        private void AddUsersToStMI_Click(object sender, EventArgs e)//Переход на фому добавления
        {
            DobavleniePolzovatelya addotdel = new DobavleniePolzovatelya();//Создание новой переменной формы DobavleniePodotdela
            addotdel.lat = lat;//Передача данных lat на форму DobavleniePolzovatelya
            addotdel.lng = lng;//Передача данных lng на форму DobavleniePolzovatelya
            addotdel.ShowDialog();//Демонстрация формы
        }
        private void AboutTheProgToStMI_Click(object sender, EventArgs e)//Переход на форму о программе
        {
            AB ab = new AB();//Создание новой переменной формы AB
            ab.ShowDialog();//Демонстрация формы
        }
        private void HelpToStMI_Click(object sender, EventArgs e)//Обработчик перехода на форму справка
        {
            Help.ShowHelp(this, "Справка.chm");//подключение справки
        } 
        private void ScreenMapTSMI_Click(object sender, EventArgs e)//Обработчик для снимка карты
        {
            //Вывод сообщения о совершении действия
            DialogResult result = MessageBox.Show("Вы точно хотите сделать снимок карты?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No) //Если пользователь нажал Нет
            {              
            }
            if (result == DialogResult.Yes) //Если пользователь нажал Да
            {
                //выполнение снимка
                Bitmap bitmap = new Bitmap(gMapBKO.Width, gMapBKO.Height);
                using (Graphics gr = Graphics.FromImage(bitmap))
                {
                    gr.CopyFromScreen(gMapBKO.PointToScreen(Point.Empty), Point.Empty, gMapBKO.Size);
                }
                
                if (bitmap != null)//Открытие диалогового окна сохранения изображения
                {
                    sFDScreen.Filter = "Png image|*.png|All files|*.*";//фильтр на файлы
                    sFDScreen.Title = "Сохранение скриншота";//Заголовок
                    if (sFDScreen.ShowDialog() == DialogResult.OK)//Нажатие ОК
                    {
                        bitmap.Save(sFDScreen.FileName);//Сохранение снимка
                    }
                    //Сообщение о сохранение
                    MessageBox.Show("Скриншот сохранен!","Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }           
        }                   
    }      
        private void AothorizTSMI_Click(object sender, EventArgs e)//Обработчик перехода на форму авторизации
        {
            ChangingThePasswordAndLogin changLogPass = new ChangingThePasswordAndLogin();//Создание новой переменной формы ChangingThePasswordAndLogin
            changLogPass.ShowDialog();//Демонстрация формы
        }
        private void Major_FormClosed(object sender, FormClosedEventArgs e)//Обработчик закрытия формы
        {
            Application.Exit();//закрытие формы}}}
//Код формы Authorization.cs:
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//Пространство имен для работы со стилем формы
using MetroFramework.Components;
using MetroFramework.Forms;
//Пространсво имен для работа с XML
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.XPath;

namespace mapsgto
{
    public partial class Authorization : MetroForm
    {
        //Переменная путь к XML файлу 
        public static string Authoriz = Application.StartupPath + "\\log&pass.xml";
        // Создание документа DOM и загрузка данных XML в него.
        public XDocument add = XDocument.Load(Authoriz);
        public Authorization()
        {
            InitializeComponent();
        }
        private void btnLogin_Click(object sender, EventArgs e)//Сохранение
        {
            ////Переменная входа, 0 - доступ закрыт, 1 - доступ открыт
            int status = 0;
            switch (cheBAdmin.Checked)
            {
                case true://Если чекбокс нажат
                    //Путь к узлу Сотрудник в XML файле log&pass
                    IEnumerable<XElement> elementsauthoAdmin = add.XPathSelectElements("/Авторизация/Сотрудник");
                    //Перебор всех элементов в файле log&pass
                    foreach (XElement el in elementsauthoAdmin)
                    {
                        //проверка: если атрибут Login равен текству введенному в tBLog и атрибут Password равен текству введенному в tBPassw
                        if (el.Attribute("Login").Value == tBLog.Text && el.Attribute("Password").Value == tBPassw.Text)
                        {
                            status = 1;//Присваивание переменной status значения, что доступ открыт
                            Major f1 = new Major();//Создание новой переменной формы Major 
                            f1.Show();//Демонстрация формы Major
                            Hide();//Закрытие формы Authorization  
                            break;
                        }
                    }
                    //Если переменная status равно 0. выводим сообщение
                    if (status == 0) MessageBox.Show("Не правильно введен логин или пароль", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                case false: //Если чекбокс не нажат
                    //Путь к узлу Сотрудник в XML файле log&pass
                    IEnumerable<XElement> elementsauthoUser = add.XPathSelectElements("/Авторизация/Сотрудник");
                    //Перебор всех элементов в файле log&pass
                    foreach (XElement el in elementsauthoUser)
                    {
                        //проверка: если атрибут Login равен текству введенному в tBLog и атрибут Password равен текству введенному в tBPassw
                        if (el.Attribute("Login").Value == tBLog.Text && el.Attribute("Password").Value == tBPassw.Text)
                        {
                            status = 1; //Присваивание переменной status значения, что доступ открыт
                            Major f1 = new Major(); //Создание новой переменной формы Major 
                            f1.Show();//Демонстрация формы Major 
                            Hide(); //Закрытие формы Authorization  
                        }
                    }
                    //Если переменная status равно 0. выводим сообщение
                    if (status == 0) MessageBox.Show("Не правильно введен логин или пароль", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);break;}}}}
//Код формы ChangingThePasswordAndLogin.cs: 
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//Пространство имен для работы со стилем формы
using MetroFramework.Components;
using MetroFramework.Forms;
//Пространсво имен для работа с XML
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.XPath;
using System.IO;

namespace mapsgto
{
    public partial class ChangingThePasswordAndLogin : MetroForm
    {
        //путь к файлу
        public static string LogPass = Application.StartupPath + "\\log&pass.xml";
        // Создание документа addLogPass и загрузка данных XML в него.
        public XDocument addLogPass = XDocument.Load(LogPass);
        public ChangingThePasswordAndLogin()
        {
            InitializeComponent();
        }
        private void ChangingThePasswordAndLogin_Load(object sender, EventArgs e)
        {
            FillDataGrid(dGVDann);//Заполнение датагрида
        }
        private void FillDataGrid(DataGridView dgv)
        {
            //Путь к узлу Сотрудник в XML файле log&pass
            IEnumerable<XElement> elementsAdmin = addLogPass.XPathSelectElements("/Авторизация/Сотрудник");
            //Очищение датагрида
            dGVDann.Rows.Clear();
            int n = 0;
            //Перебор всех элментов в XML файле GeoXML в узле Сотрудник
            foreach (XElement el in elementsAdmin)
            {
                dGVDann.Rows.Add(); //Загрузка данных в датагрид
                if (el.Attribute("Name") != null && el.Attribute("Name").Value != "")//Проверка:отсутствует ли атрибут Name или пустой он
                {
                    dGVDann.Rows[n].Cells[0].Value = el.Attribute("Name").Value;//Отображения атрибута Name в dGVDann
                }
                if (el.Attribute("Login") != null && el.Attribute("Login").Value != "")//Проверка:отсутствует ли атрибут Name или пустой он
                {
                    dGVDann.Rows[n].Cells[1].Value = el.Attribute("Login").Value;//Отображения атрибута Name в dGVDann
                }
                if (el.Attribute("Password") != null && el.Attribute("Password").Value != "")//Проверка:отсутствует ли атрибут Name или пустой он
                {
                    dGVDann.Rows[n].Cells[2].Value = el.Attribute("Password").Value;//Отображения атрибута Name в dGVDann
                }
                n++;
            }            
        }     
        private void btnSave_Click(object sender, EventArgs e)//Обработчик сохранения введенных данных
        {
            switch (chBAdmin.Checked)
            {
                case true://Если нажат чекбокс

                    if (tBLogin.Text == "")//Если текстбокс пустой
                    {
                        tStSInfo.ForeColor = Color.Red;//цвет шрифтра 
                        tStSInfo.Text = "Заполните все поля!";//сообщение
                    }
                    else
                    {
                        addAdmin();//Сохранение
                    }
                    break;

                case false:

                    if (tBLogin.Text == "")//Если текстбокс пустой
                    {
                        tStSInfo.ForeColor = Color.Red;//цвет шрифтра 
                        tStSInfo.Text = "Заполните все поля!";//сообщение
                    }
                    else
                    {
                        addUser();//Сохранение                  
                    }
                     break;
            }
        }
        private void addAdmin()//Метод сохранения данных администратора
        {
            XmlDocument xml = new XmlDocument();//Переменная для загрузки документа
            xml.Load(LogPass);//Загрузка документа
            XmlNode root = xml.DocumentElement;
            //Создаем новый узел.
            XmlElement elem = xml.CreateElement("Сотрудник");
            elem.SetAttribute("Name", chBAdmin.Text);//запись атибута Name из chBAdmin
            elem.SetAttribute("Login", tBLogin.Text);//запись атибута Login из tBLogin
            elem.SetAttribute("Password", tBPass.Text);//Password атибута tBPass из tBLogin
            //Добавляем узел в документе.
            root.InsertAfter(elem, root.LastChild);
            xml.Save(LogPass);//Загрузка документа
            tStSInfo.ForeColor = Color.Green;//цвет шрифтра 
            tStSInfo.Text = "Логин и пароль " + " успешно добавлены!";//сообщение
            int n = dGVDann.Rows.Add();//Загрузка введенных данных в датагрид
            dGVDann.Rows[n].Cells[0].Value = chBAdmin.Text; // столбец Сотрудник
            dGVDann.Rows[n].Cells[1].Value = tBLogin.Text; // столбец Логин
            dGVDann.Rows[n].Cells[2].Value = tBPass.Text; // столбец Пароль
            chBAdmin.Checked = false;//снятие галочки
            tBLogin.Text = "";//пустой текст
            tBPass.Text = "";//пустой текст
        }
        private void addUser()//Метод сохранения данных пользователя
        {
            XmlDocument xml = new XmlDocument();//Переменная для загрузки документа
            xml.Load(LogPass);//Загрузка документа
            XmlNode root = xml.DocumentElement;
            //Создаем новый узел.
            XmlElement elem = xml.CreateElement("Сотрудник");
            elem.SetAttribute("Name", labUser.Text);//запись атибута Name из labUser
            elem.SetAttribute("Login", tBLogin.Text);//запись атибута Login из tBLogin
            elem.SetAttribute("Password", tBPass.Text);//Password атибута tBPass из tBLogin

            //Добавляем узел в документе.
            root.InsertAfter(elem, root.LastChild);
            xml.Save(LogPass);//Загрузка документа
            tStSInfo.ForeColor = Color.Green;//цвет шрифтра 
            tStSInfo.Text = "Логин и пароль " + " успешно добавлены!";//сообщение
            int n = dGVDann.Rows.Add();//Загрузка введенных данных в датагрид
            dGVDann.Rows[n].Cells[0].Value = labUser.Text; // столбец Сотрудник
            dGVDann.Rows[n].Cells[1].Value = tBLogin.Text; // столбец Логин
            dGVDann.Rows[n].Cells[2].Value = tBPass.Text; // столбец Пароль
            chBAdmin.Checked = false;
            tBLogin.Text = "";
            tBPass.Text = "";      
        }
        private void dGVDann_MouseDown(object sender, MouseEventArgs e)//обработчик на клик по мыши
        {
            if (e.Button == MouseButtons.Right)//Если было нажатие правого клика
            {
                var h = dGVDann.HitTest(e.X, e.Y);//Попадание коорд. точек в dGVDann
                var n = dGVDann.HitTest(e.X, e.Y);//Попадание коорд. точек в dGVDann
                if (h.Type == DataGridViewHitTestType.Cell)
                {
                    dGVDann.ClearSelection();//снять выделение всех выбранных ячеек 
                    dGVDann.Rows[h.RowIndex].Cells[n.ColumnIndex].Selected = true;
                }
            }
        }
        private void DeleteTSMI_Click(object sender, EventArgs e)//Обработчик на удаление 
        {
            int ind = dGVDann.SelectedCells[0].RowIndex;//Задаем переменной ind, что всегда брать нулевуб ячейку
            IEnumerable<XElement> elements = addLogPass.XPathSelectElements("/Авторизация/Сотрудник");//Путь к узлу Сотрудник в XML файле log&pass            
            foreach (XElement el in elements)//Перебор всех элментов в XML файле log&pass в узле Сотрудник
            {
                //Проверка: если атрибут Name в XML файле log&pass равен нулевой ячейки из dGVDann, то удалить элемент из log&pass
                if (el.Attribute("Name").Value == dGVDann.CurrentRow.Cells[0].Value.ToString())
                {
                    el.Remove();//удаление элемента из XML файла log&pass
                }
            }
            addLogPass.Save(LogPass);//Сохранение изменений
            FillDataGrid(dGVDann);//Заполнение строк в dGVDann}}}
//Код формы DobavlenieOtdela.cs:
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GMap.NET;
using GMap.NET.WindowsForms;
using GMap.NET.WindowsForms.ToolTips;
using GMap.NET.WindowsForms.Markers;
//Пространсво имен для работа с XML
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;     
using System.Xml.XPath;
using System.IO;
//Пространство имен для работы со стилем формы
using MetroFramework.Components;
using MetroFramework.Forms;

namespace mapsgto
{
    public partial class DobavlenieOtdela : MetroForm
    {
        public double flat = 0, flng = 0;//Присваивание переменным широты и долготы нулевых значений 
        public static string infoxml = Application.StartupPath + "\\GeoXML.xml";//Переменная путь к XML файлу GeoXML        
        public XDocument dom = XDocument.Load(infoxml);// Создание документа DOM и загрузка данных XML в него.        
        public int ncount = 0;//переменная для индекса элемента в документе GeoXML
        public double lat//метод для доступа к lat
        {
            get { return lat; }
            set { flat = value; }
        }
        public double lng//метод для доступа к lng
        {
            get { return lng; }
            set { flng = value; }
        }        
        public DobavlenieOtdela()
        {
            InitializeComponent();
        }
        private void DobavlenieOtdela_Load(object sender, EventArgs e)
        {
            FillDataGrid(dGVListAbDepart);//Заполнение датагрида
        }
        private void FillDataGrid(DataGridView dgv)//Метод для заполнения датагрида
         {
            //Получаем координаты курсора
            int CursorX = Cursor.Position.X;
            int CursorY = Cursor.Position.Y;
            //Присваиваем tBflat и tBflng значения flat и flng
            tBflat.Text = Convert.ToString(flat);
            tBflng.Text = Convert.ToString(flng);
            string uri = Application.StartupPath + @"\GeoXML.xml";//Переменная путь к XML файлу GeoXML            
            IEnumerable<XElement> elements = dom.XPathSelectElements("/Отделы/Отдел");//Путь к узлу Отдел в XML файле GeoXML
            ncount = elements.Count();//Подсчет элементов
            dGVListAbDepart.Rows.Clear();//Очищение датагрида
            int n = 0;//присваивание переменной n нулевого значения            
            foreach (XElement el in elements)//Перебор всех элементов в файле GeoXML 
            {
                dGVListAbDepart.Rows.Add();//Загрузка строк в датагрид
                if (el.Attribute("Name") != null && el.Attribute("Name").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVListAbDepart.Rows[n].Cells[0].Value = el.Attribute("Name").Value;//Загрузка соответствующего атрибута в датагрид
                }

                if (el.Attribute("adress") != null && el.Attribute("adress").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVListAbDepart.Rows[n].Cells[1].Value = el.Attribute("adress").Value;//Загрузка соответствующего атрибута в датагрид
                }

                if (el.Attribute("location") != null && el.Attribute("location").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVListAbDepart.Rows[n].Cells[2].Value = el.Attribute("location").Value;//Загрузка соответствующего атрибута в датагрид
                }
                n++;//прибавление к n
            }
        }
        private void btnSave_Click(object sender, EventArgs e)//Обработчик сохранения введенных данных
        {
            if (tBDepart.Text == "")//Если текстбокс пустой
            {
                tSSLInfo.ForeColor = Color.Red;//цвет шрифтра 
                tSSLInfo.Text = "Заполните все поля!";//сообщение
            }
            else
            {
                Alladd();//Сохранение
            }
            this.Close();//Закрытие формы
        }
        private void Alladd()//Метод на сохранение введенной информации
        {
                string uri = Application.StartupPath + @"\GeoXML.xml";//Переменная путь к XML файлу GeoXML
                XmlDocument xml = new XmlDocument();//Переменная для загрузки документа
                xml.Load(uri);//Загрузка документа
                XmlNode root = xml.DocumentElement;
                //Создаем новый узел.
                XmlElement elem = xml.CreateElement("Отдел");
                elem.SetAttribute("id", Convert.ToString(ncount + 1));//Добавление атрибута id с присваиванием n+1
                elem.SetAttribute("Name", tBDepart.Text);//Добавление атрибута Name с тексот из tBDepart
                elem.SetAttribute("adress", tBAdress.Text);//Добавление атрибута adress с тексот из tBAdress
                elem.SetAttribute("location", tBLocation.Text);//Добавление атрибута location с тексот из tBLocation
                elem.SetAttribute("lat", Convert.ToString(flat));//Добавление атрибута lat
                elem.SetAttribute("lng", Convert.ToString(flng));//Добавление атрибута lng
                //Добавляем узел в документе.
                root.InsertAfter(elem, root.LastChild);
                xml.Save(uri);//Сохранение введенных данных в XML файл GeoXML
                tSSLInfo.Text = "Отдел " + tBDepart.Text + " успешно добавлено!";//Вывод сообщения о сохранении
                int n = dGVListAbDepart.Rows.Add();//Загрузка столбцов
                dGVListAbDepart.Rows[n].Cells[0].Value = tBDepart.Text; // столбец Название отдела
                dGVListAbDepart.Rows[n].Cells[1].Value = tBAdress.Text; // столбец адрес
                dGVListAbDepart.Rows[n].Cells[2].Value = tBLocation.Text; // столбец местонахождение
                tBDepart.Text = "";//Очищение текстбокса
                tBAdress.Text = "";//Очищение текстбокса
                tBLocation.Text = "";//Очищение текстбокса
        }
        private void DeleteTSMI_Click(object sender, EventArgs e)//Обработчик на удаление 
        {
            int ind = dGVListAbDepart.SelectedCells[0].RowIndex;//Задаем переменной ind, что всегда брать нулевуб ячейку
            IEnumerable<XElement> elementsdeletedgv = dom.XPathSelectElements("/Отделы/Отдел");//Путь к узлу Отдел в XML файле GeoXML
            foreach (XElement el in elementsdeletedgv)//Перебор всех элментов в XML файле GeoXML в узле Отдел
            {
                //Проверка: если атрибут Name в XML файле GeoXML равен нулевой ячейки из dGVListAbDepart, то удалить элемент из GeoXML
                if (el.Attribute("Name").Value == dGVListAbDepart.CurrentRow.Cells[0].Value.ToString())
                {
                    el.Remove();//удаление элемента из XML файла GeoXML
                }
            }
            dom.Save(infoxml);//Сохранение изменений
            FillDataGrid(dGVListAbDepart);//Заполнение строк в dGVListAbDepart
        }
        private void dGVListAbDepart_MouseDown(object sender, MouseEventArgs e)//обработчик на клик по мыши
        {
            if (e.Button == MouseButtons.Right)//Если было нажатие правого клика
            {
                var h = dGVListAbDepart.HitTest(e.X, e.Y);//Попадание коорд. точек в dGVListAbDepart
                var n = dGVListAbDepart.HitTest(e.X, e.Y);//Попадание коорд. точек в dGVListAbDepart
                if (h.Type == DataGridViewHitTestType.Cell)
                {
                    dGVListAbDepart.ClearSelection();//снять выделение всех выбранных ячеек 
                    dGVListAbDepart.Rows[h.RowIndex].Cells[n.ColumnIndex].Selected = true;}}}}}
//Код формы DobavleniePodotdela.cs:
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GMap.NET;
using GMap.NET.WindowsForms;
using GMap.NET.WindowsForms.ToolTips;
using GMap.NET.WindowsForms.Markers;
//Пространсво имен для работа с XML
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.XPath;
using System.IO;
//Пространство имен для работы со стилем формы
using MetroFramework.Components;
using MetroFramework.Forms;  

namespace mapsgto
{
    public partial class DobavleniePodotdela : MetroForm
    {
        public double flat = 0, flng = 0;//Присваивание переменным широты и долготы нулевых значений 
        public static string infoxml = Application.StartupPath + "\\GeoXML.xml";//Переменная путь к XML файлу GeoXML        
        public XDocument dom = XDocument.Load(infoxml);// Создание документа DOM и загрузка данных XML в него.
        public int ncount = 0;//переменная для индекса элемента в документе GeoXML
        public double lat//метод для доступа к lat
        {
            get { return lat; }
            set { flat = value; }
        }
        public double lng//метод для доступа к lng
        {
            get { return lng; }
            set { flng = value; }
        }
        
        public DobavleniePodotdela()
        {
            InitializeComponent();
        }
        private void DobavleniePodotdela_Load(object sender, EventArgs e)
        {
            FillDataGrid(dGVListOFUnits);//Заполнения датагрида            
        }
        private void FillDataGrid(DataGridView dgv)//Метод для заполнения датагрида
        {
            //Получаем координаты курсора
            int CursorX = Cursor.Position.X;
            int CursorY = Cursor.Position.Y;
            //Присваиваем tBflat и tBflng значения flat и flng
            tBLon.Text = Convert.ToString(flat);
            tBLat.Text = Convert.ToString(flng);
            string uri = Application.StartupPath + @"\GeoXML.xml";//Переменная путь к XML файлу GeoXML
            dGVListOFUnits.Rows.Clear();//Очищение датагрида
            IEnumerable<XElement> elementscombo = dom.XPathSelectElements("/Отделы/Отдел");//Путь к узлу Отдел в XML файле GeoXML
            foreach (XElement el in elementscombo)//Перебор всех элементов в файле GeoXML 
            {
                if (el.Attribute("Name") != null && el.Attribute("Name").Value != "")//проверка того, что атрибуты не пустые
                {
                    cBDepart.Items.Add(el.Attribute("Name").Value.Trim());//Загрузка соответствующего атрибута в датагрид
                }
            }
            int n = 0;//переменная для индекса элемента в документе GeoXML
            IEnumerable<XElement> elementsdata = dom.XPathSelectElements("/Отделы/Отдел/Подотдел");//Путь к узлу Подотдел в XML файле GeoXML
            ncount = elementsdata.Count();//Подсчет элементов
            foreach (XElement el in elementsdata)//Перебор всех элементов в файле GeoXML 
            {
                dGVListOFUnits.Rows.Add();//Загрузка строк в датагрид
                if (el.Ancestors("Отдел").Single().Attribute("Name") != null && el.Ancestors("Отдел").Single().Attribute("Name").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVListOFUnits.Rows[n].Cells[0].Value = el.Ancestors("Отдел").Single().Attribute("Name").Value;//Загрузка соответствующего атрибута в датагрид
                }
                if (el.Attribute("Name") != null && el.Attribute("Name").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVListOFUnits.Rows[n].Cells[1].Value = el.Attribute("Name").Value;//Загрузка соответствующего атрибута в датагрид
                }
                if (el.Attribute("adress") != null && el.Attribute("adress").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVListOFUnits.Rows[n].Cells[2].Value = el.Attribute("adress").Value;//Загрузка соответствующего атрибута в датагрид
                }
                if (el.Attribute("info") != null && el.Attribute("info").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVListOFUnits.Rows[n].Cells[3].Value = el.Attribute("info").Value;//Загрузка соответствующего атрибута в датагрид
                }
                n++;//прибавление к n
            }
        }
        private void btnSave_Click(object sender, EventArgs e)//Обработчик сохранения введенных данных
        {
            if (tBSubdiv.Text == "")//Если текстбокс пустой
            {
                tSSLInfo.ForeColor = Color.Red;//цвет шрифтра 
                tSSLInfo.Text = "Заполните все поля!";//сообщение
            }
            else
            {
                addAll();//сохранение
            }
            this.Close();//Закрытие формы
        }
        private void addAll ()//Метод на сохранение
        {
            string selitem = cBDepart.SelectedItem.ToString();//присваивание переменной значения выбранного элемента в комбобокс
            string uri = Application.StartupPath + @"\GeoXML.xml";//Переменная путь к XML файлу GeoXML
            XmlDocument xml = new XmlDocument();//Создание нового документа
            xml.Load(uri);//Загрузка документа
            XmlNode root = xml.DocumentElement;
            XmlNode xn = xml.SelectSingleNode("/Отделы/Отдел[@Name='" + selitem + "']");//Создаем новый узел.
            XmlElement elem = xml.CreateElement("Подотдел");//Создаем новаый элемент
            elem.SetAttribute("id", Convert.ToString(ncount + 1));//Добавление атрибута id с присваиванием n+1
            elem.SetAttribute("Name", tBSubdiv.Text);//Добавление атрибута Name с тексот из tBSubdiv
            elem.SetAttribute("adress", tBAddress.Text);//Добавление атрибута adress с тексот из tBAdress
            elem.SetAttribute("info", tBAdditInfo.Text);//Добавление атрибута info с тексот из tBAdditInfo
            elem.SetAttribute("lat", Convert.ToString(flat));//Добавление атрибута lat
            elem.SetAttribute("lng", Convert.ToString(flng));//Добавление атрибута lng                
            xn.AppendChild(elem);//Добавляем узел в документе.
            xml.Save(uri);//Сохранение введенных данных в XML файл GeoXML
            tSSLInfo.Text = "Подразделение " + tBSubdiv.Text + " успешно добавлено!";//Вывод сообщения о сохранении
            int n = dGVListOFUnits.Rows.Add();//Загрузка столбцов
            dGVListOFUnits.Rows[n].Cells[0].Value = cBDepart.Text; // Отдел
            dGVListOFUnits.Rows[n].Cells[1].Value = tBSubdiv.Text; // Подразделение
            dGVListOFUnits.Rows[n].Cells[2].Value = tBAddress.Text; // Адрес
            dGVListOFUnits.Rows[n].Cells[3].Value = tBAdditInfo.Text; // Доп.инфа
            tBSubdiv.Text = "";//Очищение текстбокса
            tBAdditInfo.Text = "";//Очищение текстбокса
            tBAddress.Text = "";//Очищение текстбокса
            cBDepart.Text = "";//Очищение комбобокса
        }
        private void DeleteToStMI_Click(object sender, EventArgs e)//Обработчик на удаление 
        {
            int ind = dGVListOFUnits.SelectedCells[0].RowIndex;//Задаем переменной ind, что всегда брать нулевуб ячейку
            IEnumerable<XElement> elementsdeletedgv = dom.XPathSelectElements("/Отделы/Отдел/Подотдел");//Путь к узлу Подотдел в XML файле GeoXML
            foreach (XElement el in elementsdeletedgv)//Перебор всех элментов в XML файле GeoXML в узле Подотдел
            {
                //Проверка: если атрибут Name в XML файле GeoXML равен нулевой ячейки из dGVListAbDepart, то удалить элемент из GeoXML
                if (el.Attribute("Name").Value == dGVListOFUnits.CurrentRow.Cells[1].Value.ToString())
                {
                    el.Remove();//удаление элемента из XML файла GeoXML
                }
            }
            dom.Save(infoxml);//Сохранение изменений
            FillDataGrid(dGVListOFUnits);//Заполнение строк в dGVListAbDepart 
        }
        private void dGVListOFUnits_MouseDown(object sender, MouseEventArgs e)//обработчик на клик по мыши
        {
            if (e.Button == MouseButtons.Right)//Если было нажатие правого клика
            {
                var h = dGVListOFUnits.HitTest(e.X, e.Y);//Попадание коорд. точек в dGVListAbDepart
                var n = dGVListOFUnits.HitTest(e.X, e.Y);//Попадание коорд. точек в dGVListAbDepart
                if (h.Type == DataGridViewHitTestType.Cell)
                {
                    dGVListOFUnits.ClearSelection();//снять выделение всех выбранных ячеек 
                    dGVListOFUnits.Rows[h.RowIndex].Cells[n.ColumnIndex].Selected = true;}}}}}

//Код формы DobavleniePolzovatelya.cs:
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GMap.NET;
using GMap.NET.WindowsForms;
using GMap.NET.WindowsForms.ToolTips;
using GMap.NET.WindowsForms.Markers;
//Пространсво имен для работа с XML
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.XPath;
using System.IO;
//Пространство имен для работы со стилем формы
using MetroFramework.Components;
using MetroFramework.Forms;

namespace mapsgto
{
    public partial class DobavleniePolzovatelya : MetroForm
    {
        public double flat = 0, flng = 0;//Присваивание переменным широты и долготы нулевых значений 
        public static string infoxml = Application.StartupPath + "\\GeoXML.xml";//Переменная путь к XML файлу GeoXML              
        public XDocument dom = XDocument.Load(infoxml);// Создание документа DOM и загрузка данных XML в него.
        public int ncount = 0;//переменная для индекса элемента в документе GeoXML
        public double lat//метод для доступа к lat
        {
            get { return lat; }
            set { flat = value; }
        }
        public double lng//метод для доступа к lng  
        {
            get { return lng; }
            set { flng = value; }
        }        
        public DobavleniePolzovatelya()
        {
            InitializeComponent();
        }
        private void DobavleniePolzovatelya_Load(object sender, EventArgs e)
        {
            FillDataGrid(dGVListUsers);//Заполнения датагрида 
            LoadCmb();//Заполнение комбобокса 
        }
        private void FillDataGrid(DataGridView dgv)//Метод на заполнение датагрида
        {
            //Получаем координаты курсора
            int CursorX = Cursor.Position.X;
            int CursorY = Cursor.Position.Y;
            //Присваиваем tBflat и tBflng значения flat и flng
            textBox6.Text = Convert.ToString(flat);
            textBox5.Text = Convert.ToString(flng);            
            dGVListUsers.Rows.Clear();//Очищение датагрида
            int n = 0;//Переменная путь к XML файлу GeoXML
            IEnumerable<XElement> elementsdata = dom.XPathSelectElements("/Отделы/Отдел/Подотдел/РМ//Пользователь");//Путь к узлу Подотдел в XML файле GeoXML
            ncount = elementsdata.Count();//Подсчет элементов
            foreach (XElement el in elementsdata)//Перебор всех элементов в файле GeoXML 
            {
                dGVListUsers.Rows.Add();//Загрузка строк в датагрид
                if (el.Attribute("ФИО") != null && el.Attribute("ФИО").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVListUsers.Rows[n].Cells[0].Value = el.Attribute("ФИО").Value;//Загрузка соответствующего атрибута в датагрид
                }
                if (el.Ancestors("Отдел").Single().Attribute("Name") != null && el.Ancestors("Отдел").Single().Attribute("Name").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVListUsers.Rows[n].Cells[1].Value = el.Ancestors("Отдел").Single().Attribute("Name").Value;//Загрузка соответствующего атрибута в датагрид
                }
                if (el.Ancestors("Подотдел").Single().Attribute("Name") != null && el.Ancestors("Подотдел").Single().Attribute("Name").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVListUsers.Rows[n].Cells[2].Value = el.Ancestors("Подотдел").Single().Attribute("Name").Value;//Загрузка соответствующего атрибута в датагрид
                }
                if (el.Ancestors("РМ").Single().Attribute("Name") != null && el.Ancestors("РМ").Single().Attribute("Name").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVListUsers.Rows[n].Cells[3].Value = el.Ancestors("РМ").Single().Attribute("Name").Value;//Загрузка соответствующего атрибута в датагрид
                }
                if (el.Attribute("Сетевое_имя") != null && el.Attribute("Сетевое_имя").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVListUsers.Rows[n].Cells[4].Value = el.Attribute("Сетевое_имя").Value;//Загрузка соответствующего атрибута в датагрид
                }
                if (el.Attribute("Должность") != null && el.Attribute("Должность").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVListUsers.Rows[n].Cells[5].Value = el.Attribute("Должность").Value;//Загрузка соответствующего атрибута в датагрид
                }
                if (el.Attribute("Телефон") != null && el.Attribute("Телефон").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVListUsers.Rows[n].Cells[6].Value = el.Attribute("Телефон").Value;//Загрузка соответствующего атрибута в датагрид
                }
                if (el.Attribute("Местонахождение") != null && el.Attribute("Местонахождение").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVListUsers.Rows[n].Cells[7].Value = el.Attribute("Местонахождение").Value;//Загрузка соответствующего атрибута в датагрид
                }
                n++;//прибавление к n
            }
        }
        private void LoadCmb()//Метод на заполнение комбобокса
        {
            cBDepart.Items.Clear();//Очищение комбобокса cBDepart
            cBSubdiv.Items.Clear();//Очищение комбобокса cBSubdiv
            cBRM.Items.Clear();//Очищение комбобокса cBRM
            IEnumerable<XElement> elementscomboO = dom.XPathSelectElements("/Отделы/Отдел");//Путь к узлу Отдел в XML файле GeoXML
            foreach (XElement el in elementscomboO)//Перебор всех элементов в файле GeoXML 
            {
                if (el.Attribute("Name") != null && el.Attribute("Name").Value != "")//проверка того, что атрибуты не пустые
                {
                    cBDepart.Items.Add(el.Attribute("Name").Value.Trim());//Загрузка соответствующего атрибута в комбобокс
                }
            }
            IEnumerable<XElement> elementscomboP = dom.XPathSelectElements("/Отделы/Отдел/Подотдел");//Путь к узлу Подотдел в XML файле GeoXML
            foreach (XElement el in elementscomboP)//Перебор всех элементов в файле GeoXML 
            {
                if (el.Parent.Attribute("Name").Value.Trim() == Convert.ToString(cBDepart.SelectedItem))//проверка того, что атрибуты не пустые
                {
                    cBSubdiv.Items.Add(el.Attribute("Name").Value.Trim());//Загрузка соответствующего атрибута в комбобокс
                }
            }
            IEnumerable<XElement> elementscomboRabochee = dom.XPathSelectElements("/Отделы/Отдел/Подотдел/РМ");//Путь к узлу РМ в XML файле GeoXML
            foreach (XElement el in elementscomboRabochee)//Перебор всех элементов в файле GeoXML 
            {
                if (el.Attribute("Name") != null && el.Attribute("Name").Value != "")//проверка того, что атрибуты не пустые
                {
                    cBRM.Items.Add(el.Attribute("Name").Value.Trim());//Загрузка соответствующего атрибута в комбобокс
                }
            }
        }
        private void btnSave_Click(object sender, EventArgs e)//Сохранение
        {
            if (tBFIO.Text == "")//Если текстбокс пустой
            {
                tSSLInfo.ForeColor = Color.Red;//цвет шрифтра 
                tSSLInfo.Text = "Заполните все поля!";//сообщение
            }
            else
            {
                addAll();//сохранение
            }
            this.Close();//Закрытие формы
        }
        private void addAll()//Метод на сохранение
        {
            string otdel = cBDepart.SelectedItem.ToString();//присваивание переменной значения выбранного элемента в комбобокс
            string podotdel = cBSubdiv.SelectedItem.ToString();//присваивание переменной значения выбранного элемента в комбобокс
            string rm = cBRM.SelectedItem.ToString();//присваивание переменной значения выбранного элемента в комбобокс
            string uri = Application.StartupPath + @"\GeoXML.xml";//Переменная путь к XML файлу GeoXML
            XmlDocument xml = new XmlDocument();//Создание нового документа
            xml.Load(uri);//Загрузка документа
            XmlNode root = xml.DocumentElement;            
            XmlNode xn = xml.SelectSingleNode("/Отделы/Отдел[@Name='" + otdel + "']/Подотдел[@Name='" + podotdel + "']/РМ[@Name='" + rm + "']");//Создаем новый узел.
            XmlElement elem = xml.CreateElement("Пользователь");//Создаем новаый элемент
            elem.SetAttribute("id", Convert.ToString(ncount + 1));//Добавление атрибута id с присваиванием n+1
            elem.SetAttribute("ФИО", tBFIO.Text);//Добавление атрибута ФИО с тексот из tBFIO
            elem.SetAttribute("Сетевое_имя", tBSetName.Text);//Добавление атрибута Сетевое_имя с тексот из tBSetName
            elem.SetAttribute("Должность", tBDol.Text);//Добавление атрибута Должность с тексот из tBDol
            elem.SetAttribute("Телефон", tBPhone.Text);//Добавление атрибута Телефон с тексот из tBPhone
            elem.SetAttribute("Местонахождение", tBLocation.Text);//Добавление атрибута Местонахождение с тексот из tBLocation
            elem.SetAttribute("lat", Convert.ToString(flat));//Добавление атрибута lat
            elem.SetAttribute("lng", Convert.ToString(flng));//Добавление атрибута lng
            xn.AppendChild(elem);//Добавляем узел в документе.
            xml.Save(uri);//Сохранение введенных данных в XML файл GeoXML
            tSSLInfo.Text = "Пользователь " + tBFIO.Text + " успешно добавлен!";//Вывод сообщения о сохранении
            int n = dGVListUsers.Rows.Add();//Загрузка столбцов
            dGVListUsers.Rows[n].Cells[0].Value = tBFIO.Text;  // ФИО
            dGVListUsers.Rows[n].Cells[1].Value = cBDepart.Text; // Отдел
            dGVListUsers.Rows[n].Cells[2].Value = cBSubdiv.Text; // Подразделение
            dGVListUsers.Rows[n].Cells[3].Value = cBRM.Text; // Рабочее место
            dGVListUsers.Rows[n].Cells[4].Value = tBSetName.Text;  // Сетевое имя
            dGVListUsers.Rows[n].Cells[5].Value = tBDol.Text;  // Должность
            dGVListUsers.Rows[n].Cells[6].Value = tBPhone.Text;  // Телефон
            dGVListUsers.Rows[n].Cells[7].Value = tBLocation.Text;  // Местонахождение
            tBFIO.Text = "";//Очищение текстбокса
            tBSetName.Text = "";//Очищение текстбокса
            tBDol.Text = "";//Очищение текстбокса
            tBPhone.Text = "";//Очищение текстбокса
            tBLocation.Text = "";//Очищение текстбокса
            cBDepart.Text = "";//Очищение текстбокса
            cBSubdiv.Text = "";//Очищение текстбокса
            cBRM.Text = "";//Очищение текстбокса
        }
        private void cBDepart_SelectedIndexChanged(object sender, EventArgs e)//Метод на выбранный элемент в комбобокс cBDepart
        {
            cBSubdiv.Enabled = true;//Включаем элемент
            cBSubdiv.Items.Clear();//Очищение списка cBSubdiv
            cBSubdiv.Text = "";// Пустой текст cBSubdiv
            IEnumerable<XElement> elementscomboP = dom.XPathSelectElements("/Отделы/Отдел/Подотдел");//Путь к узлу Подотдел в XML файле GeoXML
            foreach (XElement el in elementscomboP)//Перебор всех элементов в файле GeoXML 
            {
                if (el.Parent.Attribute("Name").Value.Trim() == Convert.ToString(cBDepart.SelectedItem))//проверка того, что атрибуты не пустые                
                    cBSubdiv.Items.Add(el.Attribute("Name").Value.Trim());//Загрузка соответствующего атрибута в комбобокс                
            }
        }
        private void cBSubdiv_SelectedIndexChanged(object sender, EventArgs e)//Метод на выбранный элемент в комбобокс cBSubdiv
        {
            cBRM.Enabled = true;//Включаем элемент
            cBRM.Items.Clear();//Очищение списка cBSubdiv
            cBRM.Text = "";// Пустой текст cBSubdiv
            IEnumerable<XElement> elementscomboP = dom.XPathSelectElements("/Отделы/Отдел/Подотдел/РМ");//Путь к узлу Подотдел в XML файле GeoXML
            foreach (XElement el in elementscomboP)//Перебор всех элементов в файле GeoXML 
            {
                if (el.Parent.Attribute("Name").Value.Trim() == Convert.ToString(cBSubdiv.SelectedItem))//проверка того, что атрибуты не пустые
                {
                    cBRM.Items.Add(el.Attribute("Name").Value.Trim());//Загрузка соответствующего атрибута в комбобокс
                }
            }
        }
        private void dGVListUser_MouseDown(object sender, MouseEventArgs e)//обработчик на клик по мыши
        {
            if (e.Button == MouseButtons.Right)//Если было нажатие правого клика
            {
                var h = dGVListUsers.HitTest(e.X, e.Y);//Попадание коорд. точек в dGVListUsers
                var n = dGVListUsers.HitTest(e.X, e.Y);//Попадание коорд. точек в dGVListUsers
                if (h.Type == DataGridViewHitTestType.Cell)
                {
                    dGVListUsers.ClearSelection();//снять выделение всех выбранных ячеек 
                    dGVListUsers.Rows[h.RowIndex].Cells[n.ColumnIndex].Selected = true;                     
                }
            } 
        }
        private void DeleteToStMI_Click(object sender, EventArgs e)//Обработчик на удаление 
        {
            int ind = dGVListUsers.SelectedCells[0].RowIndex;//Задаем переменной ind, что всегда брать нулевуб ячейку
            IEnumerable<XElement> elementsdeletedgv = dom.XPathSelectElements("/Отделы/Отдел/Подотдел/РМ/Пользователь");//Путь к узлу Подотдел в XML файле GeoXML
            foreach (XElement el in elementsdeletedgv)//Перебор всех элментов в XML файле GeoXML в узле Подотдел
            {
                //Проверка: если атрибут Name в XML файле GeoXML равен нулевой ячейки из dGVListAbDepart, то удалить элемент из GeoXML
                if (el.Attribute("ФИО").Value == dGVListUsers.CurrentRow.Cells[0].Value.ToString())
                {
                    el.Remove();//удаление элемента из XML файла GeoXML
                }
            }
            dom.Save(infoxml);//Сохранение изменений
            FillDataGrid(dGVListUsers);//Заполнение строк в dGVListAbDepart}}}

//Код формы DobavlenieRM.cs:
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GMap.NET;
using GMap.NET.WindowsForms;
using GMap.NET.WindowsForms.ToolTips;
using GMap.NET.WindowsForms.Markers;
//Пространсво имен для работа с XML
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.XPath;
using System.IO;
//Пространство имен для работы со стилем формы
using MetroFramework.Components;
using MetroFramework.Forms;

namespace mapsgto
{
    public partial class DobavlenieRM : MetroForm
    {
        public double flat = 0, flng = 0;//Присваивание переменным широты и долготы нулевых значений 
        public static string infoxml = Application.StartupPath + "\\GeoXML.xml";//Переменная путь к XML файлу GeoXML               
        public XDocument dom = XDocument.Load(infoxml);// Создание документа DOM и загрузка данных XML в него.
        public int ncount = 0;//переменная для индекса элемента в документе GeoXML
        public double lat//метод для доступа к lat
        {
            get { return lat; }
            set { flat = value; }
        }
        public double lng
        {
            get { return lng; }//метод для доступа к lng  
            set { flng = value; }
        }        
        public DobavlenieRM()
        {
            InitializeComponent();
        }
        private void DobavlenieRM_Load(object sender, EventArgs e)
        {
            FillDataGrid(dGVListOFWorkpla);//Заполнение датагрида
        }            
        private void FillDataGrid(DataGridView dgv)
        {
            //Получаем координаты курсора
            int CursorX = Cursor.Position.X;
            int CursorY = Cursor.Position.Y;
            //Присваиваем tBflat и tBflng значения flat и flng
            tBLat.Text = Convert.ToString(flat);
            tBLon.Text = Convert.ToString(flng);
            IEnumerable<XElement> elementscomboO = dom.XPathSelectElements("/Отделы/Отдел");//Путь к узлу Отдел в XML файле GeoXML
            foreach (XElement el in elementscomboO)//Перебор всех элементов в файле GeoXML 
            {
                if (el.Attribute("Name") != null && el.Attribute("Name").Value != "")//проверка того, что атрибуты не пустые
                {
                    cBDepart.Items.Add(el.Attribute("Name").Value.Trim());//Загрузка соответствующего атрибута в комбобокс
                }
            }
            IEnumerable<XElement> elementscomboP = dom.XPathSelectElements("/Отделы/Отдел/Подотдел");//Путь к узлу Подотдел в XML файле GeoXML
            foreach (XElement el in elementscomboP)//Перебор всех элементов в файле GeoXML 
            {
                if (el.Attribute("Name") != null && el.Attribute("Name").Value != "")//проверка того, что атрибуты не пустые
                {
                    cBSubdiv.Items.Add(el.Attribute("Name").Value.Trim());//Загрузка соответствующего атрибута в комбобокс
                }
            }
            dGVListOFWorkpla.Rows.Clear();//Очищение датагрида
            IEnumerable<XElement> elementsdata = dom.XPathSelectElements("/Отделы/Отдел/Подотдел/РМ");//Путь к узлу Подотдел в XML файле GeoXML
            ncount = elementsdata.Count();//Подсчет элементов
            int n = 0;
            foreach (XElement el in elementsdata)//Перебор всех элементов в файле GeoXML 
            {
                dGVListOFWorkpla.Rows.Add();//Загрузка строк в датагрид
                if (el.Ancestors("Отдел").Single().Attribute("Name") != null && el.Ancestors("Отдел").Single().Attribute("Name").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVListOFWorkpla.Rows[n].Cells[0].Value = el.Ancestors("Отдел").Single().Attribute("Name").Value;//Загрузка соответствующего атрибута в комбобокс
                }

                if (el.Ancestors("Подотдел").Single().Attribute("Name") != null && el.Ancestors("Подотдел").Single().Attribute("Name").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVListOFWorkpla.Rows[n].Cells[1].Value = el.Ancestors("Подотдел").Single().Attribute("Name").Value;//Загрузка соответствующего атрибута в комбобокс
                }

                if (el.Attribute("Name") != null && el.Attribute("Name").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVListOFWorkpla.Rows[n].Cells[2].Value = el.Attribute("Name").Value;//Загрузка соответствующего атрибута в комбобокс
                }
                n++;//прибавление к n
            }
        }
        private void btnSave_Click(object sender, EventArgs e)//Сохранение
        {
            if (tBLat.Text == "")//Если текстбокс пустой
            {
                tSSLInfo.ForeColor = Color.Red;//цвет шрифтра 
                tSSLInfo.Text = "Заполните все поля!";//сообщение
            }
            else
            {
                addAll();//сохранение
            }
            this.Close();//Закрытие формы
        }
        private void addAll()//Метод на сохранение
        {
            string otdel = cBDepart.SelectedItem.ToString();//присваивание переменной значения выбранного элемента в комбобокс
            string podotdel = cBSubdiv.SelectedItem.ToString();//присваивание переменной значения выбранного элемента в комбобокс
            string uri = Application.StartupPath + @"\GeoXML.xml";//Переменная путь к XML файлу GeoXML
            XmlDocument xml = new XmlDocument();//Создание нового документа
            xml.Load(uri);//Загрузка документа
            XmlNode root = xml.DocumentElement;
            XmlNode xn = xml.SelectSingleNode("/Отделы/Отдел[@Name='" + otdel + "']/Подотдел[@Name='" + podotdel + "']");//Создаем новый узел.
            XmlElement elem = xml.CreateElement("РМ");//Создаем новаый элемент
            elem.SetAttribute("id", Convert.ToString(ncount + 1));//Добавление атрибута id с присваиванием n+1
            elem.SetAttribute("Name", tBNaim.Text);//Добавление атрибута Name с тексот из tBNaim
            elem.SetAttribute("lat", Convert.ToString(flat));//Добавление атрибута lat
            elem.SetAttribute("lng", Convert.ToString(flng));//Добавление атрибута lng            
            xn.AppendChild(elem);//Добавляем узел в документе.
            xml.Save(uri);//Сохранение введенных данных в XML файл GeoXML
            tSSLInfo.Text = "Рабочее место " + tBNaim.Text + " успешно добавлено!";//Вывод сообщения о сохранении
            int n = dGVListOFWorkpla.Rows.Add();//Загрузка столбцов
            dGVListOFWorkpla.Rows[n].Cells[0].Value = cBDepart.Text; // Отдел
            dGVListOFWorkpla.Rows[n].Cells[1].Value = cBSubdiv.Text; // Подразделение
            dGVListOFWorkpla.Rows[n].Cells[2].Value = tBNaim.Text; // Рабочее место
            tBNaim.Text = "";//Очищение текстбокса
            cBDepart.Text = "";//Очищение комбобокса
            cBSubdiv.Text = "";//Очищение комбобокса
        }
        private void cBDepart_SelectedIndexChanged(object sender, EventArgs e)//Метод на выбранный элемент в комбобокс cBDepart
        {
            cBSubdiv.Enabled = true;//Включаем элемент
            cBSubdiv.Items.Clear();//Очищение списка cBSubdiv
            cBSubdiv.Text = "";// Пустой текст cBSubdiv
            IEnumerable<XElement> elementscomboP = dom.XPathSelectElements("/Отделы/Отдел/Подотдел");//Путь к узлу Подотдел в XML файле GeoXML
            foreach (XElement el in elementscomboP)//Перебор всех элементов в файле GeoXML 
            {
                if (el.Parent.Attribute("Name").Value.Trim() == Convert.ToString(cBDepart.SelectedItem))//проверка того, что атрибуты не пустые                
                    cBSubdiv.Items.Add(el.Attribute("Name").Value.Trim());//Загрузка соответствующего атрибута в комбобокс                
            }
        }
        private void dGVListOFWorkpla_MouseDown(object sender, MouseEventArgs e)//обработчик на клик по мыши
        {
            if (e.Button == MouseButtons.Right)//Если было нажатие правого клика
            {
                var h = dGVListOFWorkpla.HitTest(e.X, e.Y);//Попадание коорд. точек в dGVListUsers
                var n = dGVListOFWorkpla.HitTest(e.X, e.Y);//Попадание коорд. точек в dGVListUsers
                if (h.Type == DataGridViewHitTestType.Cell)
                {
                    dGVListOFWorkpla.ClearSelection();//снять выделение всех выбранных ячеек 
                    dGVListOFWorkpla.Rows[h.RowIndex].Cells[n.ColumnIndex].Selected = true;
                }
            } 
        }
        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)//Обработчик на удаление 
        {
            int ind = dGVListOFWorkpla.SelectedCells[0].RowIndex;//Задаем переменной ind, что всегда брать нулевуб ячейку
            IEnumerable<XElement> elementsdeletedgv = dom.XPathSelectElements("/Отделы/Отдел/Подотдел/РМ");//Путь к узлу Подотдел в XML файле GeoXML
            foreach (XElement el in elementsdeletedgv)//Перебор всех элментов в XML файле GeoXML в узле Подотдел
            {
                //Проверка: если атрибут Name в XML файле GeoXML равен нулевой ячейки из dGVListAbDepart, то удалить элемент из GeoXML
                if (el.Attribute("Name").Value == dGVListOFWorkpla.CurrentRow.Cells[2].Value.ToString())
                {
                    el.Remove();//удаление элемента из XML файла GeoXML
                }
            }
            dom.Save(infoxml);//Сохранение изменений
            FillDataGrid(dGVListOFWorkpla);//Заполнение строк в dGVListAbDepart}}}

//Код формы InformationAboutMarkers.cs:
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GMap.NET;
using GMap.NET.WindowsForms;
using GMap.NET.WindowsForms.ToolTips;
using GMap.NET.WindowsForms.Markers;
//Пространсво имен для работа с XML
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.XPath;
using System.IO;
//Пространство имен для работы со стилем формы
using MetroFramework.Components;
using MetroFramework.Forms;
//Пространство имен для работы с Microsoft Office Word
using Word = Microsoft.Office.Interop.Word;

namespace mapsgto
{
    public partial class InformationAboutMarkers : MetroForm
    {
        public int ncount = 0;
        public static string infoxml = Application.StartupPath + "\\GeoXML.xml";//Переменная путь к XML файлу GeoXML            
        public XDocument dom = XDocument.Load(infoxml);// Создание документа DOM и загрузка данных XML в него.
        public InformationAboutMarkers()
        {
            InitializeComponent();
        }
        private void InformationAboutMarkers_Load(object sender, EventArgs e)
        {
            FillDataGrid(dGVContInfoUser);//Заполнение датагрида
        }
        private void FillDataGrid(DataGridView dgv)
        {            
            dGVContInfoUser.Rows.Clear();//Очищение датагрида
            int n = 0;
            IEnumerable<XElement> elementsdata = dom.XPathSelectElements("/Отделы/Отдел/Подотдел/РМ//Пользователь");//Путь к узлу Отдел в XML файле GeoXML
            //Отображение инф. о пользователе
            foreach (XElement el in elementsdata)
            {               
                dGVContInfoUser.Rows.Add();
                if (el.Ancestors("РМ").Single().Attribute("Name") != null && el.Ancestors("РМ").Single().Attribute("Name").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVContInfoUser.Rows[n].Cells[0].Value = el.Ancestors("РМ").Single().Attribute("Name").Value;//Загрузка соответствующего атрибута в комбобокс
                }

                if (el.Attribute("ФИО") != null && el.Attribute("ФИО").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVContInfoUser.Rows[n].Cells[1].Value = el.Attribute("ФИО").Value;//Загрузка соответствующего атрибута в комбобокс
                }

                if (el.Ancestors("Отдел").Single().Attribute("Name") != null && el.Ancestors("Отдел").Single().Attribute("Name").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVContInfoUser.Rows[n].Cells[2].Value = el.Ancestors("Отдел").Single().Attribute("Name").Value;//Загрузка соответствующего атрибута в комбобокс
                }

                if (el.Ancestors("Подотдел").Single().Attribute("Name") != null && el.Ancestors("Подотдел").Single().Attribute("Name").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVContInfoUser.Rows[n].Cells[3].Value = el.Ancestors("Подотдел").Single().Attribute("Name").Value;//Загрузка соответствующего атрибута в комбобокс
                }

                if (el.Attribute("Сетевое_имя") != null && el.Attribute("Сетевое_имя").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVContInfoUser.Rows[n].Cells[4].Value = el.Attribute("Сетевое_имя").Value;//Загрузка соответствующего атрибута в комбобокс
                }

                if (el.Attribute("Должность") != null && el.Attribute("Должность").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVContInfoUser.Rows[n].Cells[5].Value = el.Attribute("Должность").Value;//Загрузка соответствующего атрибута в комбобокс
                }
                if (el.Attribute("Телефон") != null && el.Attribute("Телефон").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVContInfoUser.Rows[n].Cells[6].Value = el.Attribute("Телефон").Value;//Загрузка соответствующего атрибута в комбобокс
                }

                if (el.Attribute("Местонахождение") != null && el.Attribute("Местонахождение").Value != "")//проверка того, что атрибуты не пустые
                {
                    dGVContInfoUser.Rows[n].Cells[7].Value = el.Attribute("Местонахождение").Value;//Загрузка соответствующего атрибута в комбобокс
                }
                n++;//прибавление к n           
            }
            SearhOnName();
        }
        private void SearhOnName()
        {
            for (int i = 0; i<=dGVContInfoUser.RowCount -1;i++)
            {
                if (dGVContInfoUser.Rows[i].Cells[0].Value.ToString()==labInfo.Text)
                {
                        labRM.Text = dGVContInfoUser.Rows[i].Cells[0].Value.ToString();
                        
                        labFIO.Text = dGVContInfoUser.Rows[i].Cells[1].Value.ToString();
                    
                        labDepart.Text = dGVContInfoUser.Rows[i].Cells[2].Value.ToString();
                        
                        labSubdiv.Text = dGVContInfoUser.Rows[i].Cells[3].Value.ToString();
                        
                        labNetWName.Text = dGVContInfoUser.Rows[i].Cells[4].Value.ToString();
                    
                        labPosition.Text = dGVContInfoUser.Rows[i].Cells[5].Value.ToString();
                        
                        labPhone.Text = dGVContInfoUser.Rows[i].Cells[6].Value.ToString();
                        
                        labLocation.Text = dGVContInfoUser.Rows[i].Cells[7].Value.ToString();
                    }
                }
            }
        private void pictBoOUT_Click(object sender, EventArgs e)//Вывод инофрмации в ворд
        {
            //Запуск нового экземпляра Word
            Word.Application app = new Word.Application();
            //Загрузка документа
            Word.Document doc = app.Documents.Add();
            //Текст, который будет выводиться в ворд
            doc.Paragraphs[1].Range.Text = "                             Информация о пользователе" + "\n" + "Рабочее место:  " + labRM.Text + "\n" + "ФИО:  " + labFIO.Text +
            "\n" + "Отдел:  " + labDepart.Text +
            "\n" + "Подразделение:  " + labSubdiv.Text +
            "\n" + "Сетевое имя:  " + labNetWName.Text +
            "\n" + "Должность:  " + labPosition.Text +
            "\n" + "Телефон:  " + labPhone.Text +
            "\n" + "Местонахождение:  " + labLocation.Text;
            //Вкл. видимость 
            app.Visible = true;
            this.Close();//Закрытие формы}}}

//Код формы AB.cs:
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework.Components;
using MetroFramework.Forms;
namespace mapsgto
{
    partial class AB : MetroForm
    {
        public AB()
        {
            InitializeComponent();
            MetroStyleManager.Default.Style = MetroFramework.MetroColorStyle.Orange;
            MetroStyleManager.Default.Theme = MetroFramework.MetroThemeStyle.Default;
            this.Text = String.Format("О программе Геолокация пользователей АО БКО", AssemblyTitle);
            this.labelProductName.Text = AssemblyProduct;
            this.labelVersion.Text = String.Format("Версия {0}", AssemblyVersion);
            this.labelCopyright.Text = AssemblyCopyright;
            this.labelCompanyName.Text = AssemblyCompany;
            this.textBoxDescription.Text = AssemblyDescription;
            this.labelProductName.Text = "Геолокация парка АО БКО";
            this.labelVersion.Text = "0.0.1";
            this.labelCopyright.Text = "Авторские права закрепелены за Дроботенко К.С.";
            this.labelCompanyName.Text = "Christina Drobotenko";
            this.textBoxDescription.Text = "Данный программный продукт предназначен для получения информации о местонахождении цехов, отделов, подразделений и пользователей организации.";
        }
        #region 
        public string AssemblyTitle
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != "")
                    {
                        return titleAttribute.Title;
                    }
                }
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }
        public string AssemblyVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }
        public string AssemblyDescription
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }
        public string AssemblyProduct
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }
        public string AssemblyCopyright
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }
        public string AssemblyCompany
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";}
                return ((AssemblyCompanyAttribute)attributes[0]).Company;}} #endregion 
rivate void okButton_Click(object sender, EventArgs e){Close();}}}
