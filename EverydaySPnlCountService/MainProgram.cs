using System.ServiceProcess;

namespace EverydaySPnlCountService
{
    public partial class MainProgram : ServiceBase
    {
        MainMethod mm = new MainMethod();
        public MainProgram()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            mm.Start();
        }

        protected override void OnStop()
        {
            mm.Stop();
        }
    }
}
