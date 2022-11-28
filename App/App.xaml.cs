using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using ExcelToDbf.Core;

namespace ExcelToDbf
{
    /// <summary>
    /// Логика взаимодействия для App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            new Program().Run(e.Args);
        }

        public static string Version => Assembly.GetExecutingAssembly().GetName().Version.ToString();
    }
}
