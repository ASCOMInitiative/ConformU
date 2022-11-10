using System;
using System.Collections.Generic;

using Semver;
using System.Reflection;
using ASCOM.Common.Interfaces;
using ASCOM.Common;
using System.Threading.Tasks;
//using Octokit;

namespace ConformU
{
    internal class Update
    {
        #region Internal properties

        /// <summary>
        /// True if a newer release verison is available
        /// </summary>
        internal static bool HasNewerRelease { get; set; }

        /// <summary>
        /// Latest release name
        /// </summary>
        internal static string LatestReleaseName { get; set; } = "";
        /// <summary>
        /// Latest release version
        /// </summary>
        internal static string LatestReleaseVersion { get; set; } = "";

        /// <summary>
        /// Download URL for the latest release version
        /// </summary>
        internal static string ReleaseUrl { get; set; }

        /// <summary>
        /// True if a new preview version is available
        /// </summary>
        internal static bool HasNewerPreview { get; set; }

        /// <summary>
        /// Latest preview version
        /// </summary>
        internal static string LatestPreviewName { get; set; } = "";

        /// <summary>
        /// Latest preview version
        /// </summary>
        internal static string LatestPreviewVersion { get; set; } = "";

        /// <summary>
        /// Download URL for the latest preview version
        /// </summary>
        internal static string PreviewURL { get; set; } = "";

        /// <summary>
        /// True if the client is running the latest release version
        /// </summary>
        internal static bool UpToDate { get; set; }

        /// <summary>
        /// True if the client has a verison that is ahead of the latest preview release
        /// </summary>
        internal static bool AheadOfPreview { get; set; } = false;

        /// <summary>
        /// True if some releases have been retrieved from GitHub
        /// </summary>
        internal static bool HasReleases { get => Releases.Count > 0; }

        /// <summary>
        /// List of releases
        /// </summary>
        internal static IReadOnlyList<Octokit.Release> Releases { get; set; } = new List<Octokit.Release>(); //null;

        #endregion

        #region Internal methods

        internal static async Task<IReadOnlyList<Octokit.Release>> GetReleases(ConformLogger logger = null)
        {
            try
            {
                logger?.LogMessage("CheckForUpdatesSync", MessageLevel.Debug, "Getting release details");
                //Task<IReadOnlyList<Octokit.Release>> getReleasesTask = Task.Run(() => { return GitHubReleases.GetReleases("ASCOMInitiative", "ConformU"); });
                //logger?.LogMessage("CheckForUpdatesSync", MessageLevel.Debug, "Waiting for task to complete");

                //Task.WaitAll(getReleasesTask);
                //Releases = getReleasesTask.Result;
                Releases = await GitHubReleases.GetReleases("ASCOMInitiative", "ConformU");
                SetProperties(logger);
                logger?.LogMessage("CheckForUpdatesSync", MessageLevel.Debug, $"Found {Releases.Count} releases");

                foreach (Octokit.Release release in Releases)
                {
                    logger?.LogMessage("CheckForUpdatesSync", MessageLevel.Debug, $"Found release: {release.Name}, ReleaseSemVersionFromTag: {release.ReleaseSemVersionFromTag().ToString()}, Published on: {release.PublishedAt.GetValueOrDefault()}, Major: {release.ReleaseSemVersionFromTag().Major}, Minor: {release.ReleaseSemVersionFromTag().Minor}, Build: {release.ReleaseSemVersionFromTag().Patch}, Pre-release: {release.Prerelease}");
                }

                return Releases;
            }
            catch (Exception ex)
            {
                logger?.LogMessage("CheckForUpdatesSync", MessageLevel.Debug, $"Exception: {ex}");
                throw;
            }
        }

        internal async static Task CheckForUpdates(ConformLogger logger = null)
        {
            try
            {
                logger?.LogMessage("CheckForUpdates", MessageLevel.Debug, "Getting release details");
                Releases = await Task.Run(() => { return GitHubReleases.GetReleases("ASCOMInitiative", "ConformU"); });
                SetProperties(logger);
                logger?.LogMessage("CheckForUpdates", MessageLevel.Debug, $"Found {Releases.Count} releases");

                foreach (Octokit.Release release in Releases)
                {
                    logger?.LogMessage("CheckForUpdates", MessageLevel.Debug, $"Found release: {release.Name}, ReleaseSemVersionFromTag: {release.ReleaseSemVersionFromTag().ToString()}, Published on: {release.PublishedAt.GetValueOrDefault()}, Major: {release.ReleaseSemVersionFromTag().Major}, Minor: {release.ReleaseSemVersionFromTag().Minor}, Build: {release.ReleaseSemVersionFromTag().Patch}, Pre-release: {release.Prerelease}");
                }
            }
            catch (Exception ex)
            {
                logger?.LogMessage("CheckForUpdates", MessageLevel.Debug, $"Exception: {ex}");
                throw;
            }
        }

        internal static bool UpdateAvailable(ConformLogger logger = null)
        {
            try
            {
                if (Releases != null)
                {
                    if (Releases.Count > 0)
                    {
                        Version appVersion = Assembly.GetExecutingAssembly().GetName().Version;
                        if (SemVersion.TryParse($"{appVersion.Major}.{appVersion.Minor}.{appVersion.Build}", SemVersionStyles.AllowV, out SemVersion currentversion))
                        {
                            var Release = Releases?.Latest();

                            if (Release != null)
                            {
                                if (SemVersion.TryParse(Release.TagName, SemVersionStyles.AllowV, out SemVersion latestrelease))
                                {
                                    return SemVersion.CompareSortOrder(currentversion, latestrelease) == -1;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger?.LogMessage("CheckForUpdates", MessageLevel.Debug, $"Exception: {ex}");
            }
            return false;
        }

        #endregion

        #region Support code
        /// <summary>
        /// Set properties according to the releases returned
        /// </summary>
        /// <param name="logger">ConfgormLogger instance to record operational messages</param>
        private static void SetProperties(ConformLogger logger)
        {
            try
            {
                logger?.LogMessage("Update.SetProperties", MessageLevel.Debug, $"Running...");
                Version appVersion = Assembly.GetExecutingAssembly().GetName().Version;
                if (SemVersion.TryParse($"{appVersion.Major}.{appVersion.Minor}.{appVersion.Build}", SemVersionStyles.AllowV, out SemVersion installedRelease))
                {
                    logger?.LogMessage("Update.SetProperties", MessageLevel.Debug, $"Installed version: {installedRelease}");

                    var LatestRelease = Update.Releases?.LatestRelease();
                    var LatestPrerelease = Update.Releases?.LatestPrerelease();
                    if ((LatestRelease is not null) & (LatestPrerelease is not null))
                    {

                        SemVersion.TryParse(LatestRelease.TagName, SemVersionStyles.AllowV, out SemVersion latestrelease);

                        SemVersion.TryParse(LatestPrerelease.TagName, SemVersionStyles.AllowV, out SemVersion latestprerelease);
                        logger?.LogMessage("Update.SetProperties", MessageLevel.Debug, $"latestrelease: {latestrelease}, latestprerelease: {latestprerelease}");

                        if (installedRelease == latestrelease || installedRelease == latestprerelease)
                        {
                            UpToDate = true;
                        }

                        if (latestrelease != null)
                        {
                            logger?.LogMessage("Update.SetProperties", MessageLevel.Debug, $"SemVersion.CompareSortOrder(currentversion, latestrelease) == -1: {SemVersion.CompareSortOrder(installedRelease, latestrelease) == -1}");
                            if (SemVersion.CompareSortOrder(installedRelease, latestrelease) == -1)  //(installedRelease < latestrelease)
                            {
                                HasNewerRelease = true;
                                LatestReleaseVersion = LatestRelease.TagName;
                                LatestReleaseName = LatestRelease.Name;
                                ReleaseUrl = LatestRelease.HtmlUrl;
                            }
                        }
                        else
                        {
                            latestrelease = new SemVersion(0);
                        }

                        if (latestprerelease != null)
                        {
                            logger?.LogMessage("Update.SetProperties", MessageLevel.Debug, $"(SemVersion.CompareSortOrder(currentversion, latestprerelease) == -1) && (SemVersion.CompareSortOrder(latestrelease, latestprerelease) == -1): {(SemVersion.CompareSortOrder(installedRelease, latestprerelease) == -1) && (SemVersion.CompareSortOrder(latestrelease, latestprerelease) == -1)}");
                            if ((SemVersion.CompareSortOrder(installedRelease, latestprerelease) == -1) && (SemVersion.CompareSortOrder(latestrelease, latestprerelease) == -1)) //installedRelease < latestprerelease && latestrelease < latestprerelease
                            {
                                HasNewerPreview = true;
                                LatestPreviewVersion = LatestPrerelease.TagName;
                                LatestPreviewName = LatestPrerelease.Name;
                                PreviewURL = LatestPrerelease.HtmlUrl;
                            }

                            logger?.LogMessage("Update.SetProperties", MessageLevel.Debug, $"(SemVersion.CompareSortOrder(currentversion, latestprerelease) == -1) && (SemVersion.CompareSortOrder(latestrelease, latestprerelease) == 1): {(SemVersion.CompareSortOrder(installedRelease, latestprerelease) == 1) && (SemVersion.CompareSortOrder(latestrelease, latestprerelease) == -1)}");
                            if ((SemVersion.CompareSortOrder(installedRelease, latestprerelease) == 1) && (SemVersion.CompareSortOrder(latestrelease, latestprerelease) == -1)) //(installedRelease > latestprerelease && latestrelease < latestprerelease)
                            {
                                AheadOfPreview = true;
                            }
                        }
                        logger?.LogMessage("Update.SetProperties", MessageLevel.Debug, $"UpToDate: {UpToDate}, HasNewerRelease: {HasNewerRelease}, HasNewerPreview: {HasNewerPreview}, AheadOfPreview: {AheadOfPreview}, LatestVersion: {LatestReleaseVersion}, URL: {ReleaseUrl}, LatestPreviewVersion: {LatestPreviewVersion}, PreviewURL: {PreviewURL}");
                    }
                }
                else
                {
                    logger?.LogMessage("Update.SetProperties", MessageLevel.Debug, $"Failed to parse {Assembly.GetExecutingAssembly().GetName().Version}");

                }
            }
            catch (Exception ex)
            {
                logger?.LogMessage("Update.SetProperties", MessageLevel.Debug, $"Exception: {ex}");
            }
        }


        #endregion
    }
}
