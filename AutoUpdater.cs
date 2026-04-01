using Velopack;
using Velopack.Sources;

namespace SuneungMarker;

public static class AutoUpdater
{
    public const string CurrentVersion = "1.2.0";
    private const string RepoUrl = "https://github.com/ISG-kris79/SpecialEnglish";

    public static async Task CheckAndApply(Action<string>? onStatus = null)
    {
        try
        {
            var mgr = new UpdateManager(new GithubSource(RepoUrl, null, false));

            onStatus?.Invoke("업데이트 확인 중...");
            var updateInfo = await mgr.CheckForUpdatesAsync();

            if (updateInfo == null)
            {
                onStatus?.Invoke($"v{CurrentVersion} 최신 버전입니다.");
                return;
            }

            onStatus?.Invoke($"새 버전 v{updateInfo.TargetFullRelease.Version} 다운로드 중...");
            await mgr.DownloadUpdatesAsync(updateInfo);

            onStatus?.Invoke("업데이트 적용 중... 앱이 재시작됩니다.");
            mgr.ApplyUpdatesAndRestart(updateInfo);
        }
        catch (Exception ex)
        {
            onStatus?.Invoke($"업데이트 확인 실패: {ex.Message}");
        }
    }
}
