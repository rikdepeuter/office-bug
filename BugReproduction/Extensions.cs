using Microsoft.Office.Interop.Word;

namespace BugReproduction
{
    internal static class Extensions
    {
        internal static bool IsReadOnly(this Document document)
        {
            return document.ProtectionType != WdProtectionType.wdNoProtection;
        }

        internal static void SetReadOnly(this Document document, bool readOnly)
        {
            object password = "test";

            if (readOnly)
            {
                if (document.ProtectionType != WdProtectionType.wdAllowOnlyReading)
                {
                    object noReset = false;
                    object useIRM = false;
                    object enforceStyleLock = false;

                    document.Protect(WdProtectionType.wdAllowOnlyReading, noReset, password, useIRM, enforceStyleLock);
                }
            }
            else
            {
                if (document.ProtectionType != WdProtectionType.wdNoProtection)
                {
                    document.Unprotect(password);
                }
            }
        }

        internal static void AddContent(this Document document)
        {
            var range = document.Content;
            range.Collapse(WdCollapseDirection.wdCollapseEnd);
            range.Text = "Test\r";
        }
    }
}
