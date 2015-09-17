namespace MidnightDevelopers.VisualStudio.VsRestart
{
    public static class StringExtension
    {
        public static string ReplaceSmart(this string value, string oldValue, string newValue)
        {
            if (string.IsNullOrEmpty(oldValue))
            {
                return value;
            }

            return value.Replace(oldValue, newValue);
        }
    }
}