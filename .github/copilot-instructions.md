# ConformU repository instructions

## Build, test, and run commands

This repository is a single-project .NET solution targeting `net10.0` and `net10.0-windows`.

- Build the solution: `dotnet build ConformU.sln -c Debug -p:Platform="Any CPU"`
- Linting: there is no separate lint command; code-style analysis runs as part of `dotnet build` because `ConformU.csproj` enables `EnforceCodeStyleInBuild` and `AnalysisLevel=latest`.
- Run the GUI locally: `dotnet run --project ConformU\ConformU.csproj -- gui`
- Run a conformance check from the CLI: `dotnet run --project ConformU\ConformU.csproj -- conformance <COM_PROGID_or_ALPACA_URI>`
- Run an Alpaca protocol check from the CLI: `dotnet run --project ConformU\ConformU.csproj -- alpacaprotocol <Alpaca_URI>`
- Publish/release packaging is script-driven:
  - `publish.cmd` builds cross-platform packages and the Windows installer
  - `publishWindows.cmd` builds Windows publish outputs and the installer
  - `publishArm.cmd` builds Windows ARM publish outputs

There is currently no dedicated automated test project in this repository, so there is no full-suite or single-test command to run with `dotnet test`.

## High-level architecture

ConformU is one executable with two entry modes:

- **CLI mode**: `Program.cs` defines `conformance`, `alpacaprotocol`, `conformance-settings`, `alpacaprotocol-settings`, and `gui` commands with `System.CommandLine`.
- **GUI mode**: the same executable starts a localhost Kestrel server and serves a Blazor Server UI. `Startup.cs` configures Razor Pages, Server-Side Blazor, Radzen services, and a `CircuitHandler` used to react to browser disconnects. In non-debug runs, the app opens the default browser automatically.

The key runtime services are shared singletons created in `Program.CreateHostBuilder()`:

- `ConformConfiguration` owns persisted settings and validation.
- `ConformLogger` is the central log sink for console, file, and UI output.
- `SessionState` holds transient UI state such as the current screen log, protocol log, scroll position, and UI refresh events.

The Blazor pages are thin orchestration layers over those shared services:

- `Pages\Index.razor` starts and stops ASCOM device conformance runs.
- `Pages\CheckAlpacaProtocol.razor` starts and stops Alpaca protocol runs.
- `Pages\DeviceSelect.razor` selects either discovered/manual Alpaca devices or Windows COM drivers.
- `Pages\Settings.razor` edits the persisted `Settings` object, including test scope, safety gates, discovery behavior, and protocol options.

Test execution is split into two manager classes:

- `ConformanceTestManager` selects a device-specific tester based on `Settings.DeviceType` and runs the standard lifecycle: initialize, create device, pre-connect checks, connect, common methods, can-properties, pre-run checks, properties, methods, performance, post-run checks, disconnect, then configuration checks. It also writes the machine-readable JSON report file.
- `AlpacaProtocolTestManager` drives direct HTTP-level protocol validation with `HttpClient`, explicit request parameter permutations, and response/status-code checks.

Both managers accumulate test outcomes into a `ConformResults` instance (tracking `Errors`, `Issues`, `ConfigurationAlerts`, and `Timings`). The CLI process exit code is the total count of errors plus issues; zero means a clean run.

Device-specific ASCOM validation lives under `Conform\`:

- Each device type has a dedicated tester such as `CameraTester`, `TelescopeTester`, or `DomeTester`.
- Common behavior sits in `Conform\DeviceTesterBaseClass\DeviceTesterBaseClass.cs`.
- The tester layer talks to devices through facade classes under `Conform\Facades\`.

The facade layer is important for technology differences:

- Alpaca access uses client-side device objects directly.
- Windows COM access is wrapped in facade classes that host drivers on a dedicated STA thread via `DriverHostForm`, then marshal calls through `FacadeBaseClass`.

## Key conventions

- Treat `ConformConfiguration`, `ConformLogger`, and `SessionState` as the shared application backbone. New UI work should usually integrate with those existing singletons instead of creating parallel state or logging abstractions.
- Preserve the split between **device conformance** and **Alpaca protocol** paths. Conformance work belongs in `ConformanceTestManager` plus the device tester/facade hierarchy; protocol work belongs in `AlpacaProtocolTestManager`.
- New ASCOM device support is not just a new tester class. It requires wiring the device type into `ConformanceTestManager.AssignTestDevice()`, adding any settings surface needed in `Settings` / `Pages\Settings.razor`, and adding a facade when device access needs technology-specific wrapping.
- Device tester classes follow the `DeviceTesterBaseClass` lifecycle and capability flags instead of inventing custom execution order. Reuse the existing `HasPreConnectCheck`, `HasProperties`, `HasMethods`, `HasPerformanceCheck`, and related hooks.
- Repo style favors block-scoped namespaces and explicit class/method bodies, but also uses modern C# features already present in the codebase such as target-typed `new()`, collection expressions (`[]`), range/index operators, and expression-bodied members where they stay readable.
- Settings compatibility is versioned. Changes to persisted configuration should respect `SettingsCompatibilityVersion` handling in `ConformConfiguration`; do not silently change the on-disk schema without updating the compatibility flow.
- Safety-related defaults are intentional. Examples include disabled dome shutter opening and disabled switch write tests until explicitly enabled in settings. Preserve those guardrails when changing default behavior or test scope.
- The UI log panes are fed from logger events into `SessionState` (`ConformLog` and `ProtocolLog`). If you add a long-running action that should be visible in the browser, wire it into the existing logger/event flow rather than updating the UI ad hoc.
- Log calls use `TL.LogMessage(id, MessageLevel, message)`. The `MessageLevel` enum values are: `Debug` (filtered unless debug mode), `Info`, `OK`, `Issue`, `Error`, `TestOnly` (prints `id` as a section header with no level prefix), and `TestAndMessage` (prints a line with no level prefix). Match the level to the semantic outcome.
- Settings properties that must be forced to a fixed value during CLI full-test runs are decorated with `[MandatoryInFullTest(bool)]`. Annotate new settings that belong to test scope rather than user preference.
- Windows-only code is conditionally compiled with `#if WINDOWS` / `#if !WINDOWS`. The project dual-targets `net10.0-windows` (Windows, including COM/WinForms STA host) and `net10.0` (cross-platform). Keep platform-specific blocks explicit rather than relying on runtime OS checks when a compile-time guard is sufficient.
- The Blazor UI uses **Radzen** components (`RadzenButton`, `RadzenTextArea`, `NotificationService`, `DialogService`, etc.). Radzen services are registered `AddScoped` (per Blazor circuit), not singleton. `Settings.OperationInProgress` is the canonical flag that enables/disables Start/Stop buttons; set it at the start and end of any long-running operation exposed through the UI.
