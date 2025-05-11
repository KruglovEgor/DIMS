namespace DIMS.Utils
{
    using HtmlAgilityPack;

    public static class HtmlUtils
    {
        public static string ConvertHtmlToPlainText(string html)
        {
            if (string.IsNullOrWhiteSpace(html))
            {
                return string.Empty;
            }

            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(html);

            // Извлекаем текст без HTML-тегов
            return htmlDoc.DocumentNode.InnerText;
        }
    }
}
