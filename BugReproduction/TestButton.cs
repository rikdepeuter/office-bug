using System;
using System.Windows;
using Microsoft.Office.Interop.Word;

namespace BugReproduction
{
    public class TestButton
    {
        private Document document;
        private System.Windows.Forms.Timer timer;

        private bool SetDocumentToReadOnlyImmediatlyAfterOpening = true; // When you set this to false, then the bug will not happen
        private bool UseTimer = true; // When the timer is not used, then the bug will not happen

        public void Run()
        {
            document = Globals.BugReproductionAddIn.Application.ActiveDocument;

            if (SetDocumentToReadOnlyImmediatlyAfterOpening)
            {
                document.SetReadOnly(true);
            }

            WaitAndLoadContent();
        }

        void WaitAndLoadContent()
        {
            if (UseTimer)
            {
                timer = new System.Windows.Forms.Timer();
                timer.Tick += timer_Tick;
                timer.Interval = 1000;
                timer.Start();
            }
            else
            {
                timer_Tick(null, EventArgs.Empty);
            }
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            timer.Stop();
            try
            {
                document.SetReadOnly(false);
                var res = MessageBox.Show($"Document ProtectionType: {document.ProtectionType}\nWill now try to add content.", "Office bug", MessageBoxButton.OKCancel);
                if (res == MessageBoxResult.Cancel)
                {
                    return;
                }

                document.AddContent(); // EXCEPTION HAPPENS HERE
            }
            finally
            {
                timer.Start();
            }
        }
    }
}
