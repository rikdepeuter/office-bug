using System;
using Microsoft.Office.Tools.Ribbon;

namespace BugReproduction
{
    public partial class BugReproductionRibbon
    {
        private void bTest_Click(object sender, RibbonControlEventArgs e)
        {
            new TestButton().Run();
        }
    }
}
