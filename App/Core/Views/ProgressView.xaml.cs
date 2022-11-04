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
using System.Windows.Threading;
using ExcelToDbf.Core.ViewModels;

namespace ExcelToDbf.Core.Views
{
    /// <summary>
    /// Логика взаимодействия для ProgressVM.xaml
    /// </summary>
    public partial class ProgressView : UserControl
    {
        public ProgressView()
        {
            InitializeComponent();
            var timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromMilliseconds(50);
            timer.Tick += UpdateAll;
            timer.Start();
        }

        private void UpdateAll(object sender, EventArgs e)
        {
            var ctx = DataContext;
            DataContext = null;
            DataContext = ctx;
        }

    }
}
