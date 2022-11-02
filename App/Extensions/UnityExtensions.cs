using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ReactiveUI;
using Unity;

namespace ExcelToDbf.Extensions
{
    public static class UnityExtensions
    {
        public static IUnityContainer RegisterSingletonMVVM<TView, TViewModel>(this IUnityContainer container)
            where TView : Window
            where TViewModel : ReactiveObject
        {
            container.RegisterSingleton<TView>();
            container.RegisterSingleton<TViewModel>();
            return container;
        }
    }
}
