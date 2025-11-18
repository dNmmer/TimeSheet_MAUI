using System.Text.Json;

namespace Timesheet.Core.Configuration;

public sealed class ConfigService
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true,
    };

    private readonly string _configPath;

    public ConfigService(string? customPath = null)
    {
        if (!string.IsNullOrWhiteSpace(customPath))
        {
            _configPath = customPath!;
            Directory.CreateDirectory(Path.GetDirectoryName(_configPath)!);
            return;
        }

        var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var directory = Path.Combine(appData, "Timesheet_WinUI");
        Directory.CreateDirectory(directory);
        _configPath = Path.Combine(directory, "config.json");
    }

    public string ConfigPath => _configPath;

    public async Task<AppConfig> LoadAsync(CancellationToken cancellationToken = default)
    {
        if (!File.Exists(_configPath))
        {
            return AppConfig.Default;
        }

        try
        {
            await using var stream = File.OpenRead(_configPath);
            var config = await JsonSerializer.DeserializeAsync<AppConfig>(stream, SerializerOptions, cancellationToken);
            return config ?? AppConfig.Default;
        }
        catch (JsonException)
        {
            return AppConfig.Default;
        }
        catch (IOException)
        {
            return AppConfig.Default;
        }
    }

    public async Task SaveAsync(AppConfig config, CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_configPath)!);
        await using var stream = new FileStream(_configPath, FileMode.Create, FileAccess.Write, FileShare.None);
        await JsonSerializer.SerializeAsync(stream, config, SerializerOptions, cancellationToken);
    }
}
