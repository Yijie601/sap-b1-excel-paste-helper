using System.Net.Http.Headers;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace SapB1ExcelHelper.Services;

public sealed record AppUpdateInfo(
    SemanticVersion Version,
    string TagName,
    string ReleaseName,
    string ReleaseNotes,
    Uri ReleasePage,
    Uri InstallerDownload,
    string InstallerFileName,
    long InstallerSize,
    string Sha256Digest);

public sealed class UpdateService
{
    private const string ReleasesApi =
        "https://api.github.com/repos/Yijie601/sap-b1-excel-paste-helper/releases?per_page=20";

    private static readonly HttpClient Client = CreateHttpClient();

    public static SemanticVersion CurrentVersion { get; } = ReadCurrentVersion();

    public async Task<AppUpdateInfo?> CheckForUpdateAsync(CancellationToken cancellationToken = default)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, ReleasesApi);
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
        request.Headers.Add("X-GitHub-Api-Version", "2026-03-10");

        using var response = await Client.SendAsync(
            request,
            HttpCompletionOption.ResponseHeadersRead,
            cancellationToken);
        response.EnsureSuccessStatusCode();
        var json = await response.Content.ReadAsStringAsync(cancellationToken);
        return SelectAvailableUpdate(json, CurrentVersion);
    }

    public static AppUpdateInfo? SelectAvailableUpdate(string releasesJson, SemanticVersion currentVersion)
    {
        var releases = JsonSerializer.Deserialize<List<GitHubRelease>>(releasesJson) ?? new List<GitHubRelease>();
        var candidates = new List<AppUpdateInfo>();

        foreach (var release in releases)
        {
            if (release.Draft ||
                (!currentVersion.IsPreRelease && release.PreRelease) ||
                !SemanticVersion.TryParse(release.TagName, out var version) ||
                version! <= currentVersion)
            {
                continue;
            }

            var candidateVersion = version!;
            var expectedFileName = $"SapB1ExcelHelper-Setup-{candidateVersion}-win-x64.exe";
            var asset = release.Assets.FirstOrDefault(candidate =>
                string.Equals(candidate.Name, expectedFileName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(candidate.State, "uploaded", StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(candidate.Digest) &&
                candidate.Digest.StartsWith("sha256:", StringComparison.OrdinalIgnoreCase));

            if (asset is null ||
                !Uri.TryCreate(asset.BrowserDownloadUrl, UriKind.Absolute, out var downloadUri) ||
                !downloadUri.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ||
                !downloadUri.Host.Equals("github.com", StringComparison.OrdinalIgnoreCase) ||
                !Uri.TryCreate(release.HtmlUrl, UriKind.Absolute, out var releaseUri))
            {
                continue;
            }

            candidates.Add(new AppUpdateInfo(
                candidateVersion,
                release.TagName,
                string.IsNullOrWhiteSpace(release.Name) ? release.TagName : release.Name,
                release.Body ?? string.Empty,
                releaseUri,
                downloadUri,
                asset.Name,
                asset.Size,
                asset.Digest["sha256:".Length..]));
        }

        return candidates.OrderByDescending(candidate => candidate.Version).FirstOrDefault();
    }

    public async Task<string> DownloadInstallerAsync(
        AppUpdateInfo update,
        IProgress<int>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var updateDirectory = Path.Combine(
            Path.GetTempPath(),
            "SapB1ExcelHelper",
            "Updates",
            update.Version.ToString());
        Directory.CreateDirectory(updateDirectory);

        var installerPath = Path.Combine(updateDirectory, Path.GetFileName(update.InstallerFileName));
        if (File.Exists(installerPath) &&
            await VerifySha256Async(installerPath, update.Sha256Digest, cancellationToken))
        {
            progress?.Report(100);
            return installerPath;
        }

        var partialPath = installerPath + ".download";
        try
        {
            using var response = await Client.GetAsync(
                update.InstallerDownload,
                HttpCompletionOption.ResponseHeadersRead,
                cancellationToken);
            response.EnsureSuccessStatusCode();

            var totalBytes = response.Content.Headers.ContentLength ?? update.InstallerSize;
            await using var source = await response.Content.ReadAsStreamAsync(cancellationToken);
            await using var destination = new FileStream(
                partialPath,
                FileMode.Create,
                FileAccess.Write,
                FileShare.None,
                1024 * 128,
                FileOptions.Asynchronous | FileOptions.SequentialScan);

            var buffer = new byte[1024 * 128];
            long downloadedBytes = 0;
            while (true)
            {
                var bytesRead = await source.ReadAsync(buffer, cancellationToken);
                if (bytesRead == 0)
                {
                    break;
                }

                await destination.WriteAsync(buffer.AsMemory(0, bytesRead), cancellationToken);
                downloadedBytes += bytesRead;
                if (totalBytes > 0)
                {
                    progress?.Report((int)Math.Clamp(downloadedBytes * 100 / totalBytes, 0, 100));
                }
            }

            await destination.FlushAsync(cancellationToken);
            destination.Close();

            if (!await VerifySha256Async(partialPath, update.Sha256Digest, cancellationToken))
            {
                throw new InvalidDataException("The downloaded installer failed SHA-256 verification.");
            }

            File.Move(partialPath, installerPath, true);
            progress?.Report(100);
            return installerPath;
        }
        catch
        {
            try
            {
                File.Delete(partialPath);
            }
            catch
            {
                // A later update attempt can replace the partial file.
            }

            throw;
        }
    }

    public static async Task<bool> VerifySha256Async(
        string filePath,
        string expectedDigest,
        CancellationToken cancellationToken = default)
    {
        await using var stream = new FileStream(
            filePath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            1024 * 128,
            FileOptions.Asynchronous | FileOptions.SequentialScan);
        var digest = await SHA256.HashDataAsync(stream, cancellationToken);
        return Convert.ToHexString(digest).Equals(expectedDigest, StringComparison.OrdinalIgnoreCase);
    }

    private static HttpClient CreateHttpClient()
    {
        var client = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(20)
        };
        client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue(
            "SapB1ExcelHelper",
            "1.0"));
        return client;
    }

    private static SemanticVersion ReadCurrentVersion()
    {
        var assembly = typeof(UpdateService).Assembly;
        var informationalVersion = assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;
        if (SemanticVersion.TryParse(informationalVersion, out var version))
        {
            return version!;
        }

        var assemblyVersion = assembly.GetName().Version;
        return SemanticVersion.Parse(
            assemblyVersion is null
                ? "0.0.0"
                : $"{assemblyVersion.Major}.{assemblyVersion.Minor}.{Math.Max(0, assemblyVersion.Build)}");
    }

    private sealed class GitHubRelease
    {
        [JsonPropertyName("tag_name")]
        public string TagName { get; init; } = string.Empty;

        [JsonPropertyName("name")]
        public string Name { get; init; } = string.Empty;

        [JsonPropertyName("body")]
        public string? Body { get; init; }

        [JsonPropertyName("html_url")]
        public string HtmlUrl { get; init; } = string.Empty;

        [JsonPropertyName("draft")]
        public bool Draft { get; init; }

        [JsonPropertyName("prerelease")]
        public bool PreRelease { get; init; }

        [JsonPropertyName("assets")]
        public List<GitHubAsset> Assets { get; init; } = new();
    }

    private sealed class GitHubAsset
    {
        [JsonPropertyName("name")]
        public string Name { get; init; } = string.Empty;

        [JsonPropertyName("state")]
        public string State { get; init; } = string.Empty;

        [JsonPropertyName("browser_download_url")]
        public string BrowserDownloadUrl { get; init; } = string.Empty;

        [JsonPropertyName("size")]
        public long Size { get; init; }

        [JsonPropertyName("digest")]
        public string Digest { get; init; } = string.Empty;
    }
}
