using System;
using System.Activities;
using System.Activities.XamlIntegration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            Run("Activity1.xaml");
        }
        public static void Run(string fileName)
        {
            var workflow = ActivityXamlServices.Load(fileName, new ActivityXamlServicesSettings() { CompileExpressions = true });
            WorkflowInvoker.Invoke(workflow);
        }
    }
}
