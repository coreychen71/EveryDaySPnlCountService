using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;

namespace EverydaySPnlCountService
{
    [RunInstaller(true)]
    public partial class MainProgramInstaller : System.Configuration.Install.Installer
    {
        public MainProgramInstaller()
        {
            InitializeComponent();
        }
    }
}
