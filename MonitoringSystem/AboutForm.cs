using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MonitoringSystem
{
    public partial class AboutForm : Form
    {
        public AboutForm()
        {
            InitializeComponent();
            label2.Text = "Розрядність: x86 та x64.\r\n\r\nОпис: \r\nПрограмний застосунок розроблений для вирішення задачі\r\nмоніторингу забруднення атмосферного повітря в Україні.\r\nНадається можливість моніторингу забруднення повітря по \r\nнаселеним пунктам та моделювання потоків забрудненого \r\nповітря між населеними пунктами та їх вплив один на \r\nодного.";
        }

    }
}
