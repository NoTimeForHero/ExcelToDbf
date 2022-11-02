using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
using ExcelToDbf.Sources.View;
using ExcelToDbf.ViewModels;

namespace ExcelToDbf
{
    /// <summary>
    /// Логика взаимодействия для MainView.xaml
    /// </summary>
    public partial class MainView : Window
    {
        public MainView(MainViewModel model = null)
        {
            InitializeComponent();
            if (model != null) DataContext = model;

            // Так как логика всегда одинаковая, используем хардкод
            BtnAbout.Click += (o, ev) => new AboutBox().ShowDialog();
            BtnExit.Click += (o, ev) => Close();
        }
    }
}
