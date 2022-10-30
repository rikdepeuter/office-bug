using System;
using Microsoft.Office.Interop.Word;

namespace BugReproduction
{
    public class TestButton
    {
        private Document document;

        private bool SetDocumentToReadOnlyImmediatlyAfterOpening = true; // When you set this to false, then the bug will not happen
        private bool UseTimer = true; // When the timer is not used, then the bug will not happen

        public void Run()
        {
            OpenDocument(Globals.BugReproductionAddIn.Application);
        }

        void OpenDocument(Application application)
        {
            document = application.Documents.Add();

            if (SetDocumentToReadOnlyImmediatlyAfterOpening)
            {
                document.SetReadOnly(true);
            }

            ResetDocumentAndLoadContentAfterSomeTime();
        }

        void ResetDocumentAndLoadContentAfterSomeTime()
        {
            ResetDocument();

            if (UseTimer)
            {
                var timer = new System.Windows.Forms.Timer();
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
            document.SetReadOnly(false);
            document.AddContent(); // EXCEPTION HAPPENS HERE
        }

        void ResetDocument()
        {
            var isReadOnly = document.IsReadOnly();
            if (isReadOnly)
            {
                document.SetReadOnly(false);
            }

            document.Content.Delete();

            if (isReadOnly)
            {
                document.SetReadOnly(true);
            }
        }
    }
}
