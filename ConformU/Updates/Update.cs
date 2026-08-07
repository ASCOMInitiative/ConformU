using System;
using System.Collections.Generic;

using Semver;
using System.Reflection;
using System.Threading.Tasks;
using ASCOM;
using Octokit;
//using Octokit;

namespace ConformU
{
    internal class Update
    {
        #region Update state

        /// <summary>
        /// True if the client is running the latest release version
        /// </summary>
        public static bool UpToDate { get; set; } = true;

        /// <summary>
        /// True if a newer release version is available
        /// </summary>
        public static bool HasNewerRelease { get; set; } = false;

        /// <summary>
        /// True if the client has a version that is ahead of the latest main release
        /// </summary>
        public static bool AheadOfRelease { get; set; } = false;

        /// <summary>
        /// True if a new preview version is available
        /// </summary>
        public static bool HasNewerPreview { get; set; } = false;

        /// <summary>
        /// True if the client has a version that is ahead of the latest preview release
        /// </summary>
        public static bool AheadOfPreview { get; set; } = false;

        #endregion

        #region Update metadata

        /// <summary>
        /// Latest release name
        /// </summary>
        public static string LatestReleaseName { get; set; } = "";
        /// <summary>
        /// Latest release version
        /// </summary>
        public static string LatestReleaseVersion { get; set; } = "";

        /// <summary>
        /// Download URL for the latest release version
        /// </summary>
        public static string ReleaseUrl { get; set; } = "";

        /// <summary>
        /// Latest preview version
        /// </summary>
        public static string LatestPreviewName { get; set; } = "";

        /// <summary>
        /// Latest preview version
        /// </summary>
        public static string LatestPreviewVersion { get; set; } = "";

        /// <summary>
        /// Download URL for the latest preview version
        /// </summary>
        public static string PreviewURL { get; set; } = "";

        /// <summary>
        /// List of releases
        /// </summary>
        public static IReadOnlyList<Octokit.Release> Releases { get; set; } = [];

        /// <summary>
        /// True if some releases have been retrieved from GitHub
        /// </summary>
        public static bool HasReleases { get => Releases.Count > 0; }

        public static string VersionString
        {
            get
            {
                return $"{Assembly.GetEntryAssembly()?.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion}";
            }
        }

        public static string VersionDisplayString
        {
            get
            {
                string? informationalVersion = Assembly.GetEntryAssembly()?.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;

                SemVersion.TryParse(informationalVersion, SemVersionStyles.AllowV, out SemVersion? semver);
                if (semver is not null)
                    return $"{semver.Major}.{semver.Minor}.{semver.Patch}{(semver.Prerelease == "" ? "" : "-")}{semver.Prerelease} (Build {semver.Metadata})";
                else
                    return $"Unable to parse version string: '{informationalVersion}'";
            }
        }

        #endregion

        #region Public methods

        public static async Task<IReadOnlyList<Octokit.Release>> GetReleases(IAppLogger? logger = null)
        {
            try
            {
                logger?.LogDebug("CheckForUpdatesSync", "Update - Getting release details");
                Releases = await GitHubReleases.GetReleases(Globals.GITHUB_OWNER, Globals.GITHUB_REPOSITORY);
                SetProperties(logger);
                logger?.LogDebug("CheckForUpdatesSync", $"Update - Found {Releases.Count} releases");

                foreach (Octokit.Release release in Releases)
                {
                    logger?.LogDebug("CheckForUpdatesSync", $"Update - Found release: {release.Name}, ReleaseSemVersionFromTag: {release.ReleaseSemVersionFromTag()}, Published on: {release.PublishedAt.GetValueOrDefault()}, Major: {release.ReleaseSemVersionFromTag().Major}, Minor: {release.ReleaseSemVersionFromTag().Minor}, Patch: {release.ReleaseSemVersionFromTag().Patch}, Pre-release: {release.Prerelease}");
                }

                return Releases;
            }
            catch (Exception ex)
            {
                logger?.LogDebug("CheckForUpdatesSync", $"Update - Exception: {ex}");
                throw;
            }
        }

        public static async Task CheckForUpdates(IAppLogger? logger = null)
        {
            try
            {
                logger?.LogDebug("CheckForUpdates", "Update - Getting release details");
                Releases = await Task.Run(() => GitHubReleases.GetReleases(Globals.GITHUB_OWNER, Globals.GITHUB_REPOSITORY));
                SetProperties(logger);
                logger?.LogDebug("CheckForUpdates", $"Update - Found {Releases.Count} releases");

                foreach (Octokit.Release release in Releases)
                {
                    logger?.LogDebug("CheckForUpdates", $"Update - Found release: {release.Name}, ReleaseSemVersionFromTag: {release.ReleaseSemVersionFromTag()}, Published on: {release.PublishedAt.GetValueOrDefault()}, Major: {release.ReleaseSemVersionFromTag().Major}, Minor: {release.ReleaseSemVersionFromTag().Minor}, Patch: {release.ReleaseSemVersionFromTag().Patch}, Pre-release: {release.Prerelease}");
                }
            }
            catch (Exception ex)
            {
                logger?.LogDebug("CheckForUpdates", $"Update - Exception: {ex}");
                throw;
            }
        }

        public static bool UpdateAvailable(IAppLogger? logger = null)
        {
            try
            {
                if (Releases != null)
                {
                    if (Releases.Count > 0)
                    {
                        if (SemVersion.TryParse(Update.VersionString, SemVersionStyles.AllowV, out SemVersion? currentversion))
                        {
                            logger?.LogDebug("UpdateAvailable", $"Update - Application semver - Major: {currentversion.Major}, Minor: {currentversion.Minor}, Patch: {currentversion.Patch}, Pre-release: {currentversion.Prerelease}, Metadata: {currentversion.Metadata}");
                            Octokit.Release? Release = Releases?.Latest();

                            if (Release != null)
                            {
                                if (SemVersion.TryParse(Release.TagName, SemVersionStyles.AllowV, out SemVersion? latestrelease))
                                {
                                    logger?.LogDebug("UpdateAvailable", $"Update - Found release semver - Major: {latestrelease.Major}, Minor: {latestrelease.Minor}, Patch: {latestrelease.Patch}, Pre-release: {latestrelease.Prerelease}, Metadata: {latestrelease.Metadata}");
                                    return SemVersion.ComparePrecedence(currentversion, latestrelease) == -1;
                                }
                            }
                        }
                        else
                        {
                            throw new InvalidValueException($"Update - The informational product version set in the project file is not a valid SEMVER string: {Update.VersionString}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger?.LogDebug("UpdateAvailable", $"Update - Exception: {ex}");
            }
            return false;
        }

        #endregion

        #region Support code

        /// <summary>
        /// Set properties according to the releases returned
        /// </summary>
        /// <param name="logger">ILogger instance to record operational messages</param>
        private static void SetProperties(IAppLogger? logger)
        {
            try
            {
                logger?.LogDebug("Update.SetProperties", $"Update - SetProperties Running...");
                if (SemVersion.TryParse(Update.VersionString, SemVersionStyles.AllowV, out SemVersion? installedVersion))
                {
                    logger?.LogDebug("Update.SetProperties", $"Update - Installed version: {installedVersion}");

                    Octokit.Release? latestRelease = Update.Releases?.LatestRelease();
                    Octokit.Release? latestPreRelease = Update.Releases?.LatestPrerelease();
                    if ((latestRelease is not null) & (latestPreRelease is not null))
                    {

                        bool latesOk = SemVersion.TryParse(latestRelease?.TagName, SemVersionStyles.AllowV, out SemVersion? latestVersion);

                        bool latestPreOk = SemVersion.TryParse(latestPreRelease?.TagName, SemVersionStyles.AllowV, out SemVersion? latestPreReleaseVersion);

                        logger?.LogDebug("Update.SetProperties", $"Update - Installed version: {installedVersion}, Latest release: {latestVersion}, Latest pre-release: {latestPreReleaseVersion}");
                        logger?.LogDebug("Update.SetProperties", $"Update - ComparePrecedence(installedVersion, latestVersion): {SemVersion.ComparePrecedence(installedVersion, latestVersion)}");
                        logger?.LogDebug("Update.SetProperties", $"Update - ComparePrecedence(installedVersion, latestPreReleaseVersion): {SemVersion.ComparePrecedence(installedVersion, latestPreReleaseVersion)}");

                        UpToDate = (SemVersion.ComparePrecedence(installedVersion, latestVersion) == 0) || (SemVersion.ComparePrecedence(installedVersion, latestPreReleaseVersion) == 0);

                        if (latestVersion != null)
                        {
                            if (SemVersion.ComparePrecedence(installedVersion, latestVersion) == -1)  //(installedRelease < latestRelease)
                            {
                                HasNewerRelease = true;
                                LatestReleaseVersion = latestRelease?.TagName ?? "";
                                LatestReleaseName = latestRelease?.Name ?? "";
                                ReleaseUrl = latestRelease?.HtmlUrl ?? "";
                            }
                            else
                                HasNewerRelease = false;

                            if (SemVersion.ComparePrecedence(installedVersion, latestVersion) == 1)  //(installedRelease > latestRelease)
                                AheadOfRelease = true;
                            else
                                AheadOfRelease = false;
                        }
                        else
                        {
                            latestVersion = new SemVersion(0);
                        }

                        if (latestPreReleaseVersion != null)
                        {
                            logger?.LogDebug("Update.SetProperties", $"Update - Installed Release < Latest PreRelease: {SemVersion.ComparePrecedence(installedVersion, latestPreReleaseVersion) == -1}, Latest Release < Latest PreRelease: {SemVersion.ComparePrecedence(latestVersion, latestPreReleaseVersion) == -1}");
                            if ((SemVersion.ComparePrecedence(installedVersion, latestPreReleaseVersion) == -1) && (SemVersion.ComparePrecedence(latestVersion, latestPreReleaseVersion) == -1)) //installedRelease < latestPreRelease && latestRelease < latestPreRelease
                            {
                                HasNewerPreview = true;
                                LatestPreviewVersion = latestPreRelease?.TagName ?? "";
                                LatestPreviewName = latestPreRelease?.Name ?? "";
                                PreviewURL = latestPreRelease?.HtmlUrl ?? "";
                            }
                            else
                                HasNewerPreview = false;


                            logger?.LogDebug("Update.SetProperties", $"Update - Installed Release > Latest PreRelease: {SemVersion.ComparePrecedence(installedVersion, latestPreReleaseVersion) == 1}, Latest Release < Latest PreRelease: {SemVersion.ComparePrecedence(latestVersion, latestPreReleaseVersion) == -1}");
                            if ((SemVersion.ComparePrecedence(installedVersion, latestPreReleaseVersion) == 1) && (SemVersion.ComparePrecedence(latestVersion, latestPreReleaseVersion) == -1)) //(installedRelease > latestPreRelease && latestRelease < latestPreRelease)
                                AheadOfPreview = true;
                            else
                                AheadOfPreview = false;
                        }
                        logger?.LogDebug("Update.SetProperties", $"Update - UpToDate: {UpToDate}, HasNewerRelease: {HasNewerRelease}, HasNewerPreview: {HasNewerPreview}, AheadOfPreview: {AheadOfPreview}, LatestVersion: {LatestReleaseVersion}, URL: {ReleaseUrl}, LatestPreviewVersion: {LatestPreviewVersion}, PreviewURL: {PreviewURL}");
                    }
                }
                else
                {
                    logger?.LogDebug("Update.SetProperties", $"Update - Failed to parse {Update.VersionString}");
                }
            }
            catch (Exception ex)
            {
                logger?.LogDebug("Update.SetProperties", $"Update - Exception: {ex}");
            }
        }

        #endregion
    }
}
