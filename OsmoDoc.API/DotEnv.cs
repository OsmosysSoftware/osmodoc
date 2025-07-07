namespace OsmoDoc.API;

public static class DotEnv
{
    public static void LoadEnvFile(string fileName = ".env")
    {
        DirectoryInfo? dir = new DirectoryInfo(Directory.GetCurrentDirectory());

        while (dir != null)
        {
            string envPath = Path.Combine(dir.FullName, fileName);
            if (File.Exists(envPath))
            {
                Load(envPath);
                return;
            }

            dir = dir.Parent;
        }

        throw new FileNotFoundException($"{fileName} file not found");
    }
    
    public static void Load(string filePath)
    {
        foreach (string line in File.ReadAllLines(filePath))
        {
            // Check if the line contains '='
            int equalsIndex = line.IndexOf('=');
            if (equalsIndex == -1)
            {
                continue; // Skip lines without '='
            }

            string key = line.Substring(0, equalsIndex).Trim();
            string value = line.Substring(equalsIndex + 1).Trim();

            // Check if the value starts and ends with double quotation marks
            if (value.StartsWith("\"") && value.EndsWith("\""))
            {
                // Remove the double quotation marks
                value = value[1..^1];
            }

            Environment.SetEnvironmentVariable(key, value);
        }
    }
}
