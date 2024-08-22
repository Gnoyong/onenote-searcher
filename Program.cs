using Newtonsoft.Json;
using ScipBe.Common.Office.OneNote;


namespace OneNoteSearcher
{
    public class Utils
    {
        public static string search(string keyword)
        {
            if (string.IsNullOrEmpty(keyword))
            {
                return "Missing keyword parameter";
            }

            var pages = OneNoteProvider.FindPages(keyword);
            return JsonConvert.SerializeObject(pages);
        }

        public static void open(string id)
        {
            if (string.IsNullOrEmpty(id))
            {
                return "Missing id parameter";
            }

            Microsoft.Office.Interop.OneNote.Application oneNote;
            oneNote = new Microsoft.Office.Interop.OneNote.Application();
            oneNote.NavigateTo(id);
        }
    }
}

