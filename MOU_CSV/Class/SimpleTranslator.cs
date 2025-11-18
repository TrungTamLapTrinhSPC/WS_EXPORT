using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Windows.Forms;
using System.Xml.Linq;
using MOU_CSV;
using System.Reflection;

public static class SimpleTranslator
{
    // Đường dẫn file xml
    //public static string TranslationFilePath = Path.Combine(Application.StartupPath, "Translations.xml");

    // Kết hợp với file XML
    public static string TranslationFilePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Translations.xml");

    // Các biến static cần thiết
    public static string NgonNgu { get; set; } = cls_Bien.NGONNGU;
    public static Dictionary<string, string> dicTranslations { get; set; } = new Dictionary<string, string>();

    /// <summary>
    /// Hàm dịch: input là tiếng Anh (mặc định), trả về tiếng Nhật nếu tìm thấy
    /// </summary>
    public static string T(string textENG)
    {
        if (NgonNgu == "JAP")
        {
            // Nếu ngôn ngữ là tiếng Nhật, tìm bản dịch tiếng Nhật
            if (dicTranslations.TryGetValue(textENG, out string japText) && !string.IsNullOrEmpty(japText))
            {
                return japText;
            }
        }
        // Mặc định trả về tiếng Anh
        return textENG;
    }

    /// <summary>
    /// Chuyển đổi ngôn ngữ của toàn bộ controls trong form theo biến NgonNgu
    /// </summary>
    public static void ApplyLanguageSettings(Control parent)
    {
        // Duyệt toàn bộ controls trong form
        foreach (Control ctrl in GetAllControls(parent))
        {
            if (!string.IsNullOrEmpty(ctrl.Text))
            {
                if (NgonNgu == "JAP")
                {
                    // Nếu là tiếng Nhật
                    if (dicTranslations.ContainsKey(ctrl.Text))
                    {
                        ctrl.Text = dicTranslations[ctrl.Text];
                    }
                }
                else
                {
                    // Nếu là tiếng Anh (mặc định)
                    var engEntry = dicTranslations.FirstOrDefault(x => x.Value == ctrl.Text);
                    if (!engEntry.Equals(default(KeyValuePair<string, string>)) && !string.IsNullOrEmpty(engEntry.Key))
                    {
                        ctrl.Text = engEntry.Key;
                    }
                }
            }
        }
    }

    /// <summary>
    /// Hàm đệ quy lấy tất cả controls trong form (bao gồm control con)
    /// </summary>
    public static IEnumerable<Control> GetAllControls(Control parent)
    {
        var controls = new List<Control>();
        foreach (Control ctrl in parent.Controls)
        {
            controls.Add(ctrl);
            controls.AddRange(GetAllControls(ctrl));
        }
        return controls;
    }

    /// <summary>
    /// Hàm dịch: input là tiếng Anh, trả về tiếng Nhật nếu tìm thấy.
    /// Nếu không tìm thấy sẽ trả về chính input (hoặc bạn có thể trả Nothing)
    /// </summary>
    public static string TranslateToJapanese(string eng,
                                            Dictionary<string, string> translations,
                                            bool fallbackToInput = true)
    {
        if (eng == null)
        {
            return fallbackToInput ? string.Empty : null;
        }

        string key = eng.Trim();

        if (translations == null)
        {
            return fallbackToInput ? key : null;
        }

        // Tra dictionary theo ignore-case (đã khởi tạo với StringComparer.OrdinalIgnoreCase)
        if (translations.ContainsKey(key))
        {
            return translations[key];
        }

        return fallbackToInput ? key : null;
    }

    /// <summary>
    /// Ví dụ helper: loại bỏ dấu (nếu bạn muốn dùng) - không bắt buộc
    /// </summary>
    public static string RemoveDiacritics(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text;
        }

        string normalized = text.Normalize(NormalizationForm.FormD);
        var sb = new StringBuilder();

        foreach (char ch in normalized)
        {
            UnicodeCategory cat = CharUnicodeInfo.GetUnicodeCategory(ch);
            if (cat != UnicodeCategory.NonSpacingMark)
            {
                sb.Append(ch);
            }
        }

        return sb.ToString().Normalize(NormalizationForm.FormC);
    }

    /// <summary>
    /// Load dictionary từ tiếng Anh sang tiếng Nhật từ file XML
    /// </summary>
    public static Dictionary<string, string> LoadEngToJapDictionary()
    {
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        if (!File.Exists(TranslationFilePath))
        {
            return dict;
        }

        XDocument doc = XDocument.Load(TranslationFilePath);
        XElement root = doc.Root;

        if (root == null)
        {
            return dict;
        }

        // Duyệt từng <Item>3
        foreach (var itemNode in root.Elements("Item"))
        {
            var engNode = itemNode.Element("ENG");
            var japNode = itemNode.Element("JAP");

            if (engNode != null && japNode != null)
            {
                string engText = engNode.Value.Trim();
                string japText = japNode.Value.Trim();

                if (!dict.ContainsKey(engText))
                {
                    dict.Add(engText, japText);
                }
            }
        }

        return dict;
    }
}