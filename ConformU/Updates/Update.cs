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
        /// True if the client has a version that is ahead of the latest preview release
        /// </summary>
        internal static bool AheadOfPreview { get; set; } = false;

        /// <summary>
        /// True if the client has a version that is ahead of the latest main release
        /// </summary>
        internal static bool AheadOfRelease { get; set; } = false;

        /// <summary>
        /// True if some releases have been retrieved from GitHub
        /// </summary>
        internal static bool HasReleases { get => Releases.Count > 0; }

        /// <summary>
        /// List of releases
        /// </summary>
        internal static IReadOnlyList<Octokit.Release> Releases { get; set; } = new List<Octokit.Release>(); //null;

        internal static string ConformuVersion
        {
            get
            {
                return $"{Assembly.GetEntryAssembly().GetCustomAttribute<AssemblyInformationalVersionAttribute>().InformationalVersion}";
            }
        }

        internal static string ConformuVersionDisplayString
        {
            get
            {
                string informationalVersion = Assembly.GetEntryAssembly().GetCustomAttribute<AssemblyInformationalVersionAttribute>().InformationalVersion;

                SemVersion.TryParse(informationalVersion, SemVersionStyles.AllowV, out SemVersion semver);
                if (semver is not null)
                    return $"{semver.Major}.{semver.Minor}.{semver.Patch}{(semver.Prerelease == "" ? "" : "-")}{semver.Prerelease} (Build {semver.Metadata})";
                else
                    return $"Unable to parse version string: '{informationalVersion}'";
            }
        }
        #endregion

        #region Internal methods

        internal static async Task<IReadOnlyList<Octokit.Release>> GetReleases(ConformLogger logger = null)
        {
            try
            {
                logger?.LogMessage("CheckForUpdatesSync", MessageLevel.Debug, "Getting release details");
                Releases = await GitHubReleases.GetReleases("ASCOMInitiative", "ConformU");
                SetProperties(logger);
                logger?.LogMessage("CheckForUpdatesSync", MessageLevel.Debug, $"Found {Releases.Count} releases");

                foreach (Octokit.Release release in Releases)
                {
                    logger?.LogMessage("CheckForUpdatesSync", MessageLevel.Debug, $"Found release: {release.Name}, ReleaseSemVersionFromTag: {release.ReleaseSemVersionFromTag()}, Published on: {release.PublishedAt.GetValueOrDefault()}, Major: {release.ReleaseSemVersionFromTag().Major}, Minor: {release.ReleaseSemVersionFromTag().Minor}, Patch: {release.ReleaseSemVersionFromTag().Patch}, Pre-release: {release.Prerelease}");
                }

                return Releases;
            }
            catch (Exception ex)
            {
                logger?.LogMessage("CheckForUpdatesSync", MessageLevel.Debug, $"Exception: {ex}");
                throw;
            }
        }

        internal static async Task CheckForUpdates(ConformLogger logger = null)
        {
            try
            {
                logger?.LogMessage("CheckForUpdates", MessageLevel.Debug, "Getting release details");
                Releases = await Task.Run(() => GitHubReleases.GetReleases("ASCOMInitiative", "ConformU"));
                SetProperties(logger);
                logger?.LogMessage("CheckForUpdates", MessageLevel.Debug, $"Found {Releases.Count} releases");

                foreach (Octokit.Release release in Releases)
                {
                    logger?.LogMessage("CheckForUpdates", MessageLevel.Debug, $"Found release: {release.Name}, ReleaseSemVersionFromTag: {release.ReleaseSemVersionFromTag()}, Published on: {release.PublishedAt.GetValueOrDefault()}, Major: {release.ReleaseSemVersionFromTag().Major}, Minor: {release.ReleaseSemVersionFromTag().Minor}, Patch: {release.ReleaseSemVersionFromTag().Patch}, Pre-release: {release.Prerelease}");
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
                        if (SemVersion.TryParse(Update.ConformuVersion, SemVersionStyles.AllowV, out SemVersion currentversion))
                        {
                            logger?.LogMessage("UpdateAvailable", MessageLevel.Debug, $"Application semver - Major: {currentversion.Major}, Minor: {currentversion.Minor}, Patch: {currentversion.Patch}, Pre-release: {currentversion.Prerelease}, Metadata: {currentversion.Metadata}");
                            Octokit.Release Release = Releases?.Latest();

                            if (Release != null)
                            {
                                if (SemVersion.TryParse(Release.TagName, SemVersionStyles.AllowV, out SemVersion latestrelease))
                                {
                                    logger?.LogMessage("UpdateAvailable", MessageLevel.Debug, $"Found release semver - Major: {latestrelease.Major}, Minor: {latestrelease.Minor}, Patch: {latestrelease.Patch}, Pre-release: {latestrelease.Prerelease}, Metadata: {latestrelease.Metadata}");
                                    return SemVersion.ComparePrecedence(currentversion, latestrelease) == -1;
                                }
                            }
                        }
                        else
                        {
                            throw new InvalidValueException($"The informational product version set in the project file is not a valid SEMVER string: {Update.ConformuVersion}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger?.LogMessage("UpdateAvailable", MessageLevel.Debug, $"Exception: {ex}");
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
                if (SemVersion.TryParse(Update.ConformuVersion, SemVersionStyles.AllowV, out SemVersion installedVersion))
                {
                    logger?.LogMessage("Update.SetProperties", MessageLevel.Debug, $"Installed version: {installedVersion}");

                    Release latestRelease = Update.Releases?.LatestRelease();
                    Release latestPreRelease = Update.Releases?.LatestPrerelease();
                    if ((latestRelease is not null) & (latestPreRelease is not null))
                    {

                        bool latesOk = SemVersion.TryParse(latestRelease.TagName, SemVersionStyles.AllowV, out SemVersion latestVersion);

                        bool latestPreOk = SemVersion.TryParse(latestPreRelease.TagName, SemVersionStyles.AllowV, out SemVersion latestPreReleaseVersion);

                        if (true)
                        {

                        }
                        else
                        {

                        }
                        logger?.LogMessage("Update.SetProperties", MessageLevel.Debug, $"latestrelease: {latestVersion}, latestprerelease: {latestPreReleaseVersion}");

                        if ((SemVersion.ComparePrecedence(installedVersion, latestVersion) == 0) || (SemVersion.ComparePrecedence(installedVersion, latestPreReleaseVersion) == 0))
                        {
                            UpToDate = true;
                        }

                        if (latestVersion != null)
                        {
                            if (SemVersion.ComparePrecedence(installedVersion, latestVersion) == -1)  //(installedRelease < latestrelease)
                            {
                                HasNewerRelease = true;
                                LatestReleaseVersion = latestRelease.TagName;
                                LatestReleaseName = latestRelease.Name;
                                ReleaseUrl = latestRelease.HtmlUrl;
                            }

                            if (SemVersion.ComparePrecedence(installedVersion, latestVersion) == 1)  //(installedRelease > latestrelease)
                            {
                                logger?.LogMessage("Update.SetProperties", MessageLevel.Debug, $"Setting AheadOfRelease True");
                                AheadOfRelease = true;
                            }
                        }
                        else
                        {
                            latestVersion = new SemVersion(0);
                        }

                        if (latestPreReleaseVersion != null)
                        {
                            logger?.LogMessage("Update.SetProperties", MessageLevel.Debug, $"(SemVersion.ComparePrecedence(currentversion, latestprerelease) == -1) && (SemVersion.ComparePrecedence(latestrelease, latestprerelease) == -1): {(SemVersion.ComparePrecedence(installedVersion, latestPreReleaseVersion) == -1) && (SemVersion.ComparePrecedence(latestVersion, latestPreReleaseVersion) == -1)}");
                            if ((SemVersion.ComparePrecedence(installedVersion, latestPreReleaseVersion) == -1) && (SemVersion.ComparePrecedence(latestVersion, latestPreReleaseVersion) == -1)) //installedRelease < latestprerelease && latestrelease < latestprerelease
                            {
                                HasNewerPreview = true;
                                LatestPreviewVersion = latestPreRelease.TagName;
                                LatestPreviewName = latestPreRelease.Name;
                                PreviewURL = latestPreRelease.HtmlUrl;
                            }

                            logger?.LogMessage("Update.SetProperties", MessageLevel.Debug, $"(SemVersion.ComparePrecedence(currentversion, latestprerelease) == -1) && (SemVersion.ComparePrecedence(latestrelease, latestprerelease) == 1): {(SemVersion.ComparePrecedence(installedVersion, latestPreReleaseVersion) == 1) && (SemVersion.ComparePrecedence(latestVersion, latestPreReleaseVersion) == -1)}");
                            if ((SemVersion.ComparePrecedence(installedVersion, latestPreReleaseVersion) == 1) && (SemVersion.ComparePrecedence(latestVersion, latestPreReleaseVersion) == -1)) //(installedRelease > latestprerelease && latestrelease < latestprerelease)
                            {
                                AheadOfPreview = true;
                            }
                        }
                        logger?.LogMessage("Update.SetProperties", MessageLevel.Debug, $"UpToDate: {UpToDate}, HasNewerRelease: {HasNewerRelease}, HasNewerPreview: {HasNewerPreview}, AheadOfPreview: {AheadOfPreview}, LatestVersion: {LatestReleaseVersion}, URL: {ReleaseUrl}, LatestPreviewVersion: {LatestPreviewVersion}, PreviewURL: {PreviewURL}");
                    }
                }
                else
                {
                    logger?.LogMessage("Update.SetProperties", MessageLevel.Debug, $"Failed to parse {Update.ConformuVersion}");
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
